"""
Taxi Dashboard - Automatisation du remplissage Google Sheets
==============================================================
Le taximan envoie un lot de 5 photos de rapports hebdomadaires manuscrits
(un mois complet). L'application :
  1. Analyse les 5 images via Gemini Vision (recettes JOUR PAR JOUR + dépenses détaillées)
  2. Fait confirmer le mois calendaire couvert et signale les périodes manquantes
  3. Affiche, pour chaque semaine, une page de vérification/correction ne
     retenant que les jours appartenant au mois confirmé
  4. Établit un bilan mensuel consolidé (page dédiée, navigable librement)
  5. Enregistre uniquement ce bilan mensuel (recettes/dépenses/solde) dans
     la feuille Google Sheets principale — il n'y a plus de feuille secondaire.

Interface organisée en 3 sections (barre latérale) :
  - 📤 Nouveau rapport : le flux principal d'analyse/enregistrement
  - 🔃 Maintenance     : outils divers
  - ⚙️ Paramètres      : choix du fichier Google Sheets + préférences
"""

import base64
import datetime
import hashlib
import json
import os
import re
import time
import unicodedata

import gspread
import requests
import streamlit as st
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession

# ============================================================
# CONFIGURATION GÉNÉRALE
# ============================================================
st.set_page_config(page_title="Taxi Dashboard - Automatisation", page_icon="🚖", layout="wide")

# Masque les vignettes natives (icône + nom + taille + croix) que Streamlit
# affiche sous chaque file_uploader : la revue visuelle des photos (Étape 1)
# remplace déjà ce rôle de façon plus soignée.
st.markdown(
    """
    <style>
    [data-testid="stFileUploaderFile"] { display: none; }
    </style>
    """,
    unsafe_allow_html=True,
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Valeur par défaut (utilisée tant que rien n'est configuré dans ⚙️ Paramètres)
SHEET_PRINCIPALE_ID_DEFAUT = "1BNlV17OasazXtFPLbp64xHwJ2RqBjPqs97_Adi4Cbuo"

GEMINI_MODEL = "gemini-3.1-flash-lite"  # confirmé via l'endpoint ListModels : GA, supporte generateContent

CONFIG_PATH = "config_app.json"

NOMS_MOIS = ["", "janvier", "février", "mars", "avril", "mai", "juin", "juillet",
             "août", "septembre", "octobre", "novembre", "décembre"]

MAX_IMAGES_PAR_LOT = 5

# Feuille principale : une colonne par mois (mois + décalage), 3 lignes de bilan
# Colonne D = janvier => colonne = numéro du mois + 3
OFFSET_COLONNE_MOIS = 3
LIGNE_RECETTES_MENSUEL = 4
LIGNE_DEPENSES_MENSUEL = 5
LIGNE_SOLDE_MENSUEL = 6

# --- Onglet "Expenses" de la feuille du client ---
# Structure observée : colonne A = nom de la rubrique, colonne C = intitulé du
# poste de dépense, colonnes D+ = un mois chacune. Les lignes dont la colonne C
# vaut "Monthly totals:" sont des lignes de TOTAUX (formules) : l'app ne doit
# jamais écrire dedans.
NOM_ONGLET_DEPENSES_DEFAUT = "Expenses"
COL_RUBRIQUE_DEPENSES = 1   # A
COL_POSTE_DEPENSES = 3      # C
LIBELLE_TOTAUX_MENSUELS = "monthly totals"

# Rubriques récurrentes (mois après mois) : privilégiées lors de l'appariement
# automatique, au détriment des rubriques ponctuelles (achat du taxi, lancement).
PRIORITE_RUBRIQUES = {
    "depenses fixes": 0,
    "depenses diverses": 1,
}

VALUE_INPUT_OPTION = "USER_ENTERED"  # évite que Sheets force les nombres en texte (bug de l'apostrophe)


# ============================================================
# CONFIGURATION UTILISATEUR (persistée sur disque)
# ============================================================
def charger_config() -> dict:
    defaut = {
        "sheet_principale_id": SHEET_PRINCIPALE_ID_DEFAUT,
        "nom_onglet_depenses": NOM_ONGLET_DEPENSES_DEFAUT,
        "nom_utilisateur": "Pascal",
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                defaut.update(json.load(f))
        except Exception:
            pass
    return defaut


def sauvegarder_config(config: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def extraire_id_depuis_url(texte: str) -> str:
    """Accepte soit un ID brut, soit une URL Google Sheets complète."""
    texte = texte.strip()
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", texte)
    return match.group(1) if match else texte


def obtenir_email_service_account() -> str:
    try:
        if os.path.exists("credentials.json"):
            with open("credentials.json", "r", encoding="utf-8") as f:
                data = json.load(f)
        else:
            data = st.secrets["google_credentials"]
        return data.get("client_email", "inconnu")
    except Exception:
        return "inconnu"


def numero_colonne_vers_lettre(n: int) -> str:
    lettres = ""
    while n > 0:
        n, reste = divmod(n - 1, 26)
        lettres = chr(65 + reste) + lettres
    return lettres


# ============================================================
# CONNEXIONS (Google Sheets + clé Gemini)
# ============================================================
def charger_credentials():
    """Charge les identifiants du compte de service, depuis le fichier local
    en développement ou depuis les secrets Streamlit en production."""
    if os.path.exists("credentials.json"):
        return Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return Credentials.from_service_account_info(st.secrets["google_credentials"], scopes=SCOPES)


@st.cache_resource(show_spinner=False)
def get_gspread_client():
    """Authentifie le compte de service Google - mis en cache car les
    identifiants ne changent jamais en cours de session."""
    return gspread.authorize(charger_credentials())


def diagnostiquer_acces_feuille(sheet_id: str) -> dict:
    """Vérifie, SANS RIEN MODIFIER, si le compte de service peut lire et
    écrire dans le classeur du client.

    On interroge l'API Drive plutôt que de tenter une écriture d'essai :
    écrire puis annuler risquerait d'effacer une formule existante dans la
    feuille d'un client. Drive renvoie directement le droit `canEdit`."""
    session = AuthorizedSession(charger_credentials())
    reponse = session.get(
        f"https://www.googleapis.com/drive/v3/files/{sheet_id}",
        params={"fields": "name,capabilities/canEdit", "supportsAllDrives": "true"},
        timeout=30,
    )

    if reponse.status_code == 404:
        return {
            "statut": "introuvable",
            "message": (
                "Le classeur est introuvable pour le compte de service. Soit l'identifiant est "
                "incorrect, soit la feuille n'a pas encore été partagée avec l'adresse ci-dessus."
            ),
        }
    if reponse.status_code == 403:
        return {
            "statut": "refuse",
            "message": "Accès refusé : la feuille n'est pas partagée avec le compte de service.",
        }
    if reponse.status_code != 200:
        return {
            "statut": "erreur",
            "message": f"Réponse inattendue de Google Drive (code {reponse.status_code}).",
        }

    infos = reponse.json()
    peut_editer = infos.get("capabilities", {}).get("canEdit", False)
    return {
        "statut": "editeur" if peut_editer else "lecture_seule",
        "titre": infos.get("name", "(sans titre)"),
        "message": (
            "Le compte de service peut lire ET modifier ce classeur."
            if peut_editer
            else "Le compte de service peut lire ce classeur, mais PAS y écrire : "
                 "le partage est en « Lecteur » au lieu de « Éditeur »."
        ),
    }


def get_clients(sheet_principale_id: str):
    """Récupère la feuille de calcul à CHAQUE appel (volontairement PAS mise
    en cache) : si on gardait l'objet Worksheet en mémoire, un onglet
    supprimé/recréé manuellement dans Google Sheets casserait l'app avec une
    erreur 'No grid with id ...' jusqu'au redémarrage. Cet appel API est
    rapide et peu coûteux vu la fréquence d'utilisation de l'app."""
    gc = get_gspread_client()
    api_key = st.secrets["GEMINI_API_KEY"]
    feuille_principale = gc.open_by_key(sheet_principale_id).get_worksheet(0)
    return api_key, feuille_principale


def get_clients_config():
    """Raccourci qui lit l'identifiant depuis la config utilisateur en session."""
    config = st.session_state.config
    return get_clients(config["sheet_principale_id"])


def get_onglet(nom_onglet: str):
    """Ouvre un onglet précis (par son nom) du classeur du client.
    Lève une erreur explicite si l'onglet n'existe pas, plutôt que l'erreur
    gspread brute peu compréhensible pour l'utilisateur."""
    gc = get_gspread_client()
    classeur = gc.open_by_key(st.session_state.config["sheet_principale_id"])
    try:
        return classeur.worksheet(nom_onglet)
    except gspread.WorksheetNotFound:
        onglets = ", ".join(f"« {f.title} »" for f in classeur.worksheets())
        raise RuntimeError(
            f"L'onglet « {nom_onglet} » est introuvable dans le classeur du client. "
            f"Onglets disponibles : {onglets}. Corrige le nom dans ⚙️ Paramètres."
        )


# ============================================================
# APPEL GEMINI VISION
# ============================================================
def construire_prompt() -> str:
    return (
        "Tu es un assistant comptable spécialisé dans la lecture de rapports "
        "manuscrits de recettes de taxi. Analyse cette image et renvoie "
        "UNIQUEMENT un objet JSON strict (sans texte autour, sans balises "
        "markdown), avec exactement cette structure :\n\n"
        "{\n"
        '  "periode_hebdo": "JJ/MM/AA - JJ/MM/AA",\n'
        '  "recettes_journalieres": [\n'
        '    {"date": "JJ/MM/AA", "montant": nombre}\n'
        "  ],\n"
        '  "depenses": [\n'
        '    {"titre": "texte court", "montant": nombre, "date": "JJ/MM/AA ou vide"}\n'
        "  ]\n"
        "}\n\n"
        "Règles strictes :\n"
        "- 'recettes_journalieres' doit lister CHAQUE ligne de recette "
        "journalière visible sur le cahier, avec sa date exacte (JJ/MM/AA) "
        "et son montant. N'en saute aucune, même si un total est aussi écrit "
        "à la main.\n"
        "- 'depenses' doit lister CHAQUE dépense mentionnée individuellement "
        "(assurance, vidange, garage, huile de frein, fournitures, main "
        "d'œuvre, déplacement, etc.) avec un titre court et son montant en "
        "chiffres uniquement. Indique la 'date' à laquelle la dépense a été "
        "notée si elle est identifiable sur le cahier (sinon laisse ce champ "
        "vide, l'application utilisera une date par défaut).\n"
        "- Le 'solde antérieur' ou 'solde à ce jour' n'est PAS une dépense, "
        "ne l'inclus pas dans la liste.\n"
        "- 'periode_hebdo' = date de la première et de la dernière ligne de "
        "recette journalière (format JJ/MM/AA).\n"
        "- N'inclus PAS de champ total_recettes, total_depenses ni "
        "solde_net : ils sont recalculés séparément par l'application.\n"
        "- Si un montant est peu lisible, donne ta meilleure estimation "
        "plutôt que de l'omettre."
    )


def encoder_image(fichier) -> tuple[str, str]:
    mime_type = fichier.type or "image/jpeg"
    fichier.seek(0)
    contenu = fichier.read()
    fichier.seek(0)  # remis à zéro : le fichier peut être relu ensuite (st.image, etc.)
    return base64.b64encode(contenu).decode("utf-8"), mime_type


def masquer_cle_api(texte: str) -> str:
    """Retire toute clé API visible d'un message d'erreur (ex: dans une URL
    '...?key=AQ.xxxx'), pour qu'elle ne puisse plus jamais être exposée par
    accident (capture d'écran, logs, partage...)."""
    if not texte:
        return texte
    return re.sub(r"key=[^&\s\"'\)]+", "key=***MASQUÉE***", texte)


def appeler_gemini(api_key: str, image_b64: str, mime_type: str, tentatives: int = 3) -> dict:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    payload = {
        "contents": [{
            "parts": [
                {"text": construire_prompt()},
                {"inline_data": {"mime_type": mime_type, "data": image_b64}},
            ]
        }],
        "generationConfig": {
            "temperature": 0,
            "response_mime_type": "application/json",
        },
    }

    for essai in range(1, tentatives + 1):
        try:
            reponse = requests.post(url, json=payload, timeout=60)
        except requests.RequestException as exc:
            raise RuntimeError(masquer_cle_api(str(exc))) from exc

        if reponse.status_code == 401 or reponse.status_code == 403:
            raise RuntimeError(
                f"Authentification refusée par Gemini (erreur {reponse.status_code}). "
                "La clé API est invalide, désactivée, ou le projet Google Cloud associé "
                "est suspendu. Vérifie l'état de ton projet sur Google AI Studio / Google Cloud Console."
            )

        if reponse.status_code == 429:
            message_api = ""
            try:
                message_api = reponse.json().get("error", {}).get("message", "")
            except Exception:
                pass

            if "PerDay" in message_api or "per day" in message_api.lower():
                raise RuntimeError(
                    "Quota GRATUIT journalier Gemini atteint pour aujourd'hui (erreur 429). "
                    "Inutile de réessayer maintenant : ce quota se réinitialise automatiquement "
                    "le lendemain. Pour lever cette limite tout de suite, active la facturation "
                    "sur Google AI Studio."
                )

            attente = int(reponse.headers.get("Retry-After", 15 * essai))
            if essai < tentatives:
                st.warning(f"⏳ Limite de débit Gemini atteinte, nouvelle tentative dans {attente}s...")
                time.sleep(attente)
                continue
            raise RuntimeError(
                "Limite de débit Gemini atteinte (erreur 429) après plusieurs tentatives. "
                "Attends 1 à 2 minutes avant de réessayer, ou vérifie ton quota sur Google AI Studio."
            )

        if reponse.status_code in (500, 502, 503, 504):
            attente = 5 * essai
            if essai < tentatives:
                st.warning(
                    f"⏳ Le service Gemini est momentanément indisponible (erreur {reponse.status_code}), "
                    f"nouvelle tentative dans {attente}s..."
                )
                time.sleep(attente)
                continue
            raise RuntimeError(
                f"Le service Gemini est indisponible (erreur {reponse.status_code}) après plusieurs tentatives. "
                "C'est un problème temporaire côté Google, pas un bug de l'application : réessaie dans "
                "quelques minutes."
            )

        if reponse.status_code >= 400:
            message_detaille = ""
            try:
                message_detaille = reponse.json().get("error", {}).get("message", "")
            except Exception:
                pass
            if not message_detaille:
                message_detaille = reponse.text[:400] or f"Erreur HTTP {reponse.status_code} sans détail."
            raise RuntimeError(
                masquer_cle_api(f"Erreur Gemini {reponse.status_code} : {message_detaille}")
            )

        return reponse.json()


def extraire_donnees(res_json: dict) -> dict:
    if "error" in res_json:
        raise RuntimeError(res_json["error"].get("message", "Erreur inconnue de l'API Gemini."))

    if not res_json.get("candidates"):
        raise RuntimeError("Gemini n'a renvoyé aucun résultat. L'image est peut-être illisible ou trop floue.")

    candidat = res_json["candidates"][0]
    if candidat.get("finishReason") == "SAFETY":
        raise RuntimeError("La demande a été bloquée par les filtres de sécurité de Gemini.")

    texte = candidat["content"]["parts"][0]["text"]
    texte = re.sub(r"^```(json)?|```$", "", texte.strip(), flags=re.MULTILINE).strip()

    try:
        donnees, _ = json.JSONDecoder().raw_decode(texte)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Réponse Gemini invalide (JSON incorrect) : {exc}") from exc

    donnees.setdefault("periode_hebdo", "")
    donnees.setdefault("recettes_journalieres", [])
    donnees.setdefault("depenses", [])
    return donnees


# ============================================================
# CALCULS / FORMATAGE
# ============================================================
def parser_date(date_str: str):
    """Tolère plusieurs formats, car l'IA (selon le modèle utilisé) peut
    renvoyer l'année sur 2 ou 4 chiffres, et parfois un autre séparateur
    que '/' (ex: '27/04/26', '27/04/2026', '27-04-2026', '27.04.26')."""
    if not date_str:
        return None
    texte = re.sub(r"[.\-]", "/", date_str.strip())
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(texte, fmt).date()
        except ValueError:
            continue
    return None


def parser_periode(periode_str: str):
    try:
        debut_str, fin_str = [p.strip() for p in periode_str.split(" - ")]
        debut, fin = parser_date(debut_str), parser_date(fin_str)
        if debut and fin:
            return debut, fin
    except Exception:
        pass
    return None


def periode_semaine_courante() -> str:
    aujourdhui = datetime.date.today()
    debut = aujourdhui - datetime.timedelta(days=aujourdhui.weekday())
    fin = debut + datetime.timedelta(days=6)
    return f"{debut.strftime('%d/%m/%y')} - {fin.strftime('%d/%m/%y')}"


def dernier_jour_du_mois(annee: int, mois: int) -> datetime.date:
    if mois == 12:
        return datetime.date(annee, 12, 31)
    return datetime.date(annee, mois + 1, 1) - datetime.timedelta(days=1)


def determiner_bornes_mois(lot_donnees: list[dict]):
    """À partir des recettes journalières des 5 rapports du lot, détermine
    le mois calendaire majoritairement couvert (celui qui contient le plus
    de jours de recette), puis renvoie ses bornes (1er jour -> dernier jour).
    Renvoie (None, None) si aucune date exploitable n'a été trouvée."""
    compteur_par_mois: dict = {}
    for donnees in lot_donnees:
        if not isinstance(donnees, dict):
            continue
        for j in donnees.get("recettes_journalieres", []):
            d = parser_date(str(j.get("date", "")))
            if d:
                cle = (d.year, d.month)
                compteur_par_mois[cle] = compteur_par_mois.get(cle, 0) + 1

    if not compteur_par_mois:
        return None, None

    annee, mois = max(compteur_par_mois.items(), key=lambda kv: kv[1])[0]
    return datetime.date(annee, mois, 1), dernier_jour_du_mois(annee, mois)


def detecter_trous_hebdomadaires(lot_donnees: list[dict]) -> list:
    """Vérifie l'enchaînement des périodes hebdomadaires des rapports du lot
    et renvoie la liste des trous détectés, sous forme de tuples
    (debut_manquant, fin_manquant)."""
    periodes = []
    for donnees in lot_donnees:
        if not isinstance(donnees, dict):
            continue
        p = parser_periode(donnees.get("periode_hebdo", ""))
        if p:
            periodes.append(p)
    periodes.sort(key=lambda p: p[0])

    trous = []
    for (debut1, fin1), (debut2, fin2) in zip(periodes, periodes[1:]):
        if debut2 > fin1 + datetime.timedelta(days=1):
            trous.append((fin1 + datetime.timedelta(days=1), debut2 - datetime.timedelta(days=1)))
    return trous


def hachage_fichier(fichier) -> str:
    fichier.seek(0)
    empreinte = hashlib.md5(fichier.read()).hexdigest()
    fichier.seek(0)
    return empreinte


def estimer_ordre_upload(fichiers: list) -> list:
    """Avant toute analyse IA, estime un ordre chronologique plausible à
    partir de l'horodatage souvent présent dans le nom du fichier (ex :
    captures d'écran 'Capture d'écran 2026-07-03 023326.png'). Purement
    indicatif pour l'affichage de l'étape de confirmation : l'ordre
    définitif est de toute façon recalculé après analyse, à partir des
    dates réellement lues sur les photos elles-mêmes."""
    motif = re.compile(r"(\d{4})-(\d{2})-(\d{2})[ _-](\d{2})(\d{2})(\d{2})")

    def cle(i):
        m = motif.search(fichiers[i].name)
        if m:
            return (0, tuple(int(x) for x in m.groups()))
        return (1, i)

    return sorted(range(len(fichiers)), key=cle)


def calculer_ordre_chronologique(resultats_lot: list[dict]) -> list[int]:
    """Détermine automatiquement l'ordre chronologique (du plus ancien au
    plus récent) des rapports du lot, à partir de la période détectée par
    l'IA sur chaque photo — l'utilisateur n'a plus besoin de les uploader
    dans le bon ordre. Les rapports en erreur (période inconnue) sont
    placés à la fin, dans leur ordre d'origine."""
    indices = list(range(len(resultats_lot)))

    def cle(i):
        donnees = resultats_lot[i]
        if isinstance(donnees, dict) and "erreur" not in donnees:
            p = parser_periode(donnees.get("periode_hebdo", ""))
            if p:
                return (0, p[0])
        return (1, i)

    return sorted(indices, key=cle)


def construire_fichiers_ordonnes(fichiers_avec_remplacements: list) -> list:
    """Renvoie la liste des fichiers dans le même ordre que
    st.session_state.lot_donnees : ordre chronologique pour les rapports
    initiaux, puis les éventuelles photos ajoutées pour combler des trous."""
    ordre = st.session_state.get("ordre_chronologique")
    if ordre and len(ordre) == len(fichiers_avec_remplacements):
        base = [fichiers_avec_remplacements[j] for j in ordre]
    else:
        base = list(fichiers_avec_remplacements)
    return base + list(st.session_state.fichiers_combles.values())


def detecter_doublons(fichiers: list, lot_donnees: list[dict]) -> list[dict]:
    """Repère les anomalies de doublon dans le lot :
    - deux fichiers strictement identiques (même image envoyée deux fois)
    - deux rapports dont les périodes sont identiques ou se chevauchent
    Renvoie une liste d'alertes {"gravite": "erreur"|"avertissement", "message": str}."""
    alertes = []

    empreintes: dict = {}
    for i, f in enumerate(fichiers):
        empreintes.setdefault(hachage_fichier(f), []).append(i)
    for indices in empreintes.values():
        if len(indices) > 1:
            noms = ", ".join(f"Semaine {i + 1}" for i in indices)
            alertes.append({
                "gravite": "erreur",
                "message": f"{noms} sont EXACTEMENT la même image (fichier identique) — doublon très probable.",
            })

    periodes = []
    for i, donnees in enumerate(lot_donnees):
        if isinstance(donnees, dict) and "erreur" not in donnees:
            p = parser_periode(donnees.get("periode_hebdo", ""))
            if p:
                periodes.append((i, p[0], p[1]))

    for a in range(len(periodes)):
        for b in range(a + 1, len(periodes)):
            i1, d1, f1 = periodes[a]
            i2, d2, f2 = periodes[b]
            if d1 <= f2 and d2 <= f1:
                if d1 == d2 and f1 == f2:
                    alertes.append({
                        "gravite": "erreur",
                        "message": (
                            f"Semaine {i1 + 1} et Semaine {i2 + 1} couvrent EXACTEMENT la même période "
                            f"({d1.strftime('%d/%m/%y')} - {f1.strftime('%d/%m/%y')}) : doublon probable."
                        ),
                    })
                else:
                    alertes.append({
                        "gravite": "avertissement",
                        "message": (
                            f"Semaine {i1 + 1} et Semaine {i2 + 1} ont des périodes qui se chevauchent "
                            f"({d1.strftime('%d/%m/%y')}-{f1.strftime('%d/%m/%y')} et "
                            f"{d2.strftime('%d/%m/%y')}-{f2.strftime('%d/%m/%y')})."
                        ),
                    })
    return alertes


def formater_montant(valeur: float) -> str:
    if float(valeur).is_integer():
        return str(int(valeur))
    return str(valeur)


def calculer_total_recettes(recettes: list[dict]) -> float:
    total = 0.0
    for r in recettes:
        try:
            total += float(r.get("montant", 0) or 0)
        except (TypeError, ValueError):
            pass
    return total


def calculer_total_depenses(depenses: list[dict]) -> float:
    total = 0.0
    for d in depenses:
        try:
            total += float(d.get("montant", 0) or 0)
        except (TypeError, ValueError):
            pass
    return total


# ============================================================
# FILTRAGE PAR MOIS + BILAN MENSUEL
# (tout se calcule désormais en mémoire à partir du lot de 5 rapports :
# il n'y a plus de feuille secondaire à lire ni à écrire)
# ============================================================
def filtrer_donnees_par_mois(donnees: dict, debut_mois, fin_mois) -> dict:
    """Ne conserve, pour un rapport hebdomadaire donné, que les jours de
    recette appartenant au mois confirmé. Les dépenses datées hors du mois
    sont également écartées ; celles sans date explicite sont rattachées
    au premier jour du mois présent dans cette semaine (et non plus
    forcément au 1er jour brut de la période, qui peut être dans le mois
    précédent)."""
    recettes_filtrees = []
    for r in donnees.get("recettes_journalieres", []):
        d = parser_date(str(r.get("date", "")))
        if d and debut_mois <= d <= fin_mois:
            recettes_filtrees.append(r)

    premier_jour_semaine_dans_mois = None
    if recettes_filtrees:
        dates_valides = [parser_date(str(r.get("date", ""))) for r in recettes_filtrees]
        dates_valides = [d for d in dates_valides if d]
        if dates_valides:
            premier_jour_semaine_dans_mois = min(dates_valides)

    depenses_filtrees = []
    for dep in donnees.get("depenses", []):
        d = parser_date(str(dep.get("date", ""))) if dep.get("date") else None
        if d is None:
            if premier_jour_semaine_dans_mois:
                depenses_filtrees.append({**dep, "date": premier_jour_semaine_dans_mois.strftime("%d/%m/%y")})
            # sinon : cette semaine n'a aucun jour dans le mois confirmé -> dépense ignorée pour ce mois
        elif debut_mois <= d <= fin_mois:
            depenses_filtrees.append(dep)

    return {**donnees, "recettes_journalieres": recettes_filtrees, "depenses": depenses_filtrees}


def calculer_bilan_mensuel_agrege(recettes_combinees: list[dict], depenses_combinees: list[dict], annee: int, mois: int) -> dict:
    """Construit le rapport détaillé du mois à partir des données déjà
    filtrées/éditées en mémoire (plus besoin de relire une feuille)."""
    jours_recette = set()
    for r in recettes_combinees:
        d = parser_date(str(r.get("date", "")))
        if d:
            jours_recette.add(d)

    recette_totale = calculer_total_recettes(recettes_combinees)

    depenses_par_titre: dict = {}
    for dep in depenses_combinees:
        titre = str(dep.get("titre", "")).strip() or "Dépense"
        try:
            montant = float(dep.get("montant", 0) or 0)
        except (TypeError, ValueError):
            montant = 0.0
        depenses_par_titre[titre] = depenses_par_titre.get(titre, 0.0) + montant

    total_depenses = sum(depenses_par_titre.values())
    return {
        "annee": annee,
        "mois": mois,
        "jours_travailles": len(jours_recette),
        "recette_totale": recette_totale,
        "depenses_par_titre": depenses_par_titre,
        "total_depenses": total_depenses,
        "solde_net": recette_totale - total_depenses,
    }


def enregistrer_bilan_mensuel(feuille_principale, rapport: dict) -> None:
    """Écrit le bilan (recettes, dépenses, solde) du mois dans la feuille
    principale, dans la colonne correspondant à ce mois."""
    colonne_mois = rapport["mois"] + OFFSET_COLONNE_MOIS
    lettre_col = numero_colonne_vers_lettre(colonne_mois)
    feuille_principale.update(
        f"{lettre_col}{LIGNE_RECETTES_MENSUEL}", [[rapport["recette_totale"]]], value_input_option=VALUE_INPUT_OPTION
    )
    feuille_principale.update(
        f"{lettre_col}{LIGNE_DEPENSES_MENSUEL}", [[rapport["total_depenses"]]], value_input_option=VALUE_INPUT_OPTION
    )
    feuille_principale.update(
        f"{lettre_col}{LIGNE_SOLDE_MENSUEL}", [[rapport["solde_net"]]], value_input_option=VALUE_INPUT_OPTION
    )


# ============================================================
# ONGLET "EXPENSES" DE LA FEUILLE DU CLIENT
# Lecture de la structure existante, appariement des dépenses du
# bilan avec les postes déjà présents, puis écriture ciblée.
# ============================================================
def message_erreur_sheets(exc: Exception) -> str:
    """Traduit les erreurs Google Sheets courantes en consignes actionnables,
    plutôt que de laisser remonter un code HTTP brut."""
    texte = str(exc)
    email_service = obtenir_email_service_account()

    if "403" in texte or "PERMISSION_DENIED" in texte or "permission" in texte.lower():
        return (
            "Accès refusé par Google : le compte de service n'a pas le droit d'écrire dans "
            f"cette feuille. Demande au client de la partager avec « {email_service} » "
            "en rôle **Éditeur**, puis relance l'opération."
        )
    if "404" in texte or "not found" in texte.lower():
        return (
            "Classeur ou onglet introuvable. Vérifie l'identifiant du classeur et le nom de "
            "l'onglet dans ⚙️ Paramètres, et que la feuille est bien partagée avec "
            f"« {email_service} »."
        )
    if "429" in texte or "quota" in texte.lower():
        return (
            "Trop de requêtes envoyées à Google Sheets en peu de temps. Patiente une minute "
            "puis réessaie."
        )
    return texte


def normaliser_libelle(texte: str) -> str:
    """Minuscules, sans accents, sans ponctuation : permet de rapprocher
    « Vidange » (feuille du client) de « vidange » ou « VIDANGE » (lu par l'IA
    sur un rapport manuscrit)."""
    texte = unicodedata.normalize("NFD", str(texte or ""))
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", texte.lower()).strip()


def lire_structure_depenses(feuille) -> dict:
    """Analyse l'onglet Expenses et en extrait la structure réelle :
    les rubriques (colonne A) et, pour chacune, la liste des postes de
    dépense (colonne C) avec leur numéro de ligne.
    Les lignes « Monthly totals: » servent de séparateurs de rubrique et
    sont exclues des destinations possibles (ce sont des formules)."""
    valeurs = feuille.get_all_values()
    rubriques = []
    rubrique_courante = None

    for numero_ligne, ligne in enumerate(valeurs, start=1):
        col_a = ligne[COL_RUBRIQUE_DEPENSES - 1].strip() if len(ligne) >= COL_RUBRIQUE_DEPENSES else ""
        col_c = ligne[COL_POSTE_DEPENSES - 1].strip() if len(ligne) >= COL_POSTE_DEPENSES else ""

        if normaliser_libelle(col_c).startswith(LIBELLE_TOTAUX_MENSUELS):
            rubrique_courante = {
                "nom": col_a or f"Rubrique (ligne {numero_ligne})",
                "ligne_totaux": numero_ligne,
                "postes": [],
            }
            rubriques.append(rubrique_courante)
        elif col_c and rubrique_courante is not None:
            rubrique_courante["postes"].append({"ligne": numero_ligne, "nom": col_c})

    return {"rubriques": rubriques, "valeurs": valeurs}


def lister_postes(structure: dict) -> list:
    """Aplatit la structure en une liste de destinations possibles."""
    postes = []
    for rubrique in structure["rubriques"]:
        for poste in rubrique["postes"]:
            postes.append({
                "ligne": poste["ligne"],
                "nom": poste["nom"],
                "rubrique": rubrique["nom"],
            })
    return postes


def priorite_rubrique(nom_rubrique: str) -> int:
    return PRIORITE_RUBRIQUES.get(normaliser_libelle(nom_rubrique), 9)


def trouver_appariement_auto(titre_depense: str, postes: list):
    """Propose la ligne la plus plausible pour une dépense du bilan.
    En cas d'homonymes dans plusieurs rubriques (ex. « Assurance » existe à la
    fois en dépenses de lancement et en dépenses fixes), on privilégie les
    rubriques récurrentes. Renvoie None si rien de convaincant."""
    cible = normaliser_libelle(titre_depense)
    if not cible:
        return None

    exacts = [p for p in postes if normaliser_libelle(p["nom"]) == cible]
    if exacts:
        return min(exacts, key=lambda p: priorite_rubrique(p["rubrique"]))

    partiels = [
        p for p in postes
        if cible in normaliser_libelle(p["nom"]) or normaliser_libelle(p["nom"]) in cible
    ]
    if partiels:
        return min(
            partiels,
            key=lambda p: (priorite_rubrique(p["rubrique"]), len(normaliser_libelle(p["nom"]))),
        )

    return None


def valeur_actuelle_cellule(structure: dict, ligne: int, mois: int) -> str:
    """Contenu actuellement affiché dans la cellule visée (pour prévenir
    l'utilisateur qu'une écriture va écraser une valeur existante)."""
    colonne = mois + OFFSET_COLONNE_MOIS
    valeurs = structure["valeurs"]
    if ligne - 1 < len(valeurs):
        ligne_valeurs = valeurs[ligne - 1]
        if colonne - 1 < len(ligne_valeurs):
            return ligne_valeurs[colonne - 1].strip()
    return ""


def ecrire_depenses_client(feuille, montants_par_ligne: dict, mois: int) -> None:
    """Écrit, en une seule requête, les montants dans la colonne du mois.
    montants_par_ligne : {numero_de_ligne: montant}."""
    lettre_col = numero_colonne_vers_lettre(mois + OFFSET_COLONNE_MOIS)
    donnees = [
        {"range": f"{lettre_col}{ligne}", "values": [[montant]]}
        for ligne, montant in sorted(montants_par_ligne.items())
    ]
    if donnees:
        feuille.batch_update(donnees, value_input_option=VALUE_INPUT_OPTION)


# ============================================================
# ÉTAT DE SESSION
# ============================================================
for cle, defaut in {
    "config": None,
    "cle_uploader": 0,
    "signature_lot": None,
    "page": "📤 Nouveau rapport",
    "lot_donnees": None,
    "donnees_filtrees": None,
    "mois_confirme": False,
    "sous_page": 0,
    "bilan_enregistre": False,
    "trous_ignores": set(),
    "fichiers_combles": {},
    "ordre_chronologique": None,
    "ordre_upload_estime": None,
    "structure_depenses": None,
    "depenses_ecrites": False,
    "semaines_validees": set(),
    "bilan_etabli": False,
    "rapport_fige": None,
    "donnees_editees": {},
}.items():
    if cle not in st.session_state:
        st.session_state[cle] = defaut

if st.session_state.config is None:
    st.session_state.config = charger_config()


def reinitialiser_lot() -> None:
    """Efface tout l'état lié au lot courant pour repartir sur un nouveau
    lot de 5 photos (nouveau mois)."""
    for cle in list(st.session_state.keys()):
        if cle.startswith("editeur_recettes_") or cle.startswith("editeur_depenses_"):
            del st.session_state[cle]
    st.session_state.cle_uploader += 1
    st.session_state.signature_lot = None
    st.session_state.lot_donnees = None
    st.session_state.donnees_filtrees = None
    st.session_state.mois_confirme = False
    st.session_state.sous_page = 0
    st.session_state.bilan_enregistre = False
    st.session_state.trous_ignores = set()
    st.session_state.fichiers_combles = {}
    st.session_state.ordre_chronologique = None
    st.session_state.ordre_upload_estime = None
    st.session_state.structure_depenses = None
    st.session_state.depenses_ecrites = False
    st.session_state.semaines_validees = set()
    st.session_state.bilan_etabli = False
    st.session_state.rapport_fige = None
    st.session_state.donnees_editees = {}
    for cle in list(st.session_state.keys()):
        if cle.startswith("dest_depense_"):
            del st.session_state[cle]


def normaliser_lignes_editeur(valeur) -> list:
    """Convertit ce que renvoie st.data_editor en une liste de dictionnaires.
    Selon le type d'entrée, Streamlit renvoie un DataFrame ou une liste : on
    normalise pour que le reste du code manipule toujours des dictionnaires."""
    if valeur is None:
        return []
    if hasattr(valeur, "to_dict"):  # DataFrame
        return valeur.to_dict("records")
    return list(valeur)


def memoriser_donnees_semaine(i: int, recettes, depenses) -> None:
    """Mémorise les données RÉELLEMENT affichées/éditées pour la semaine i.

    Indispensable : pour un st.data_editor, st.session_state[cle] ne contient
    que le journal des modifications (lignes ajoutées/modifiées/supprimées),
    et non les données résultantes. Sans cette mémorisation, les corrections
    manuelles ne seraient pas reprises dans le bilan mensuel."""
    st.session_state.donnees_editees[i] = {
        "recettes": normaliser_lignes_editeur(recettes),
        "depenses": normaliser_lignes_editeur(depenses),
    }


def donnees_initiales_semaine(i: int, champ: str) -> list:
    """Valeur de DÉPART du tableau éditable : toujours les données brutes
    issues de l'analyse, jamais les données déjà éditées.

    C'est volontaire : st.data_editor applique ses modifications par-dessus la
    valeur qu'on lui fournit. Si on lui redonnait le résultat déjà modifié, une
    ligne ajoutée par l'utilisateur serait réappliquée à chaque rechargement et
    se dupliquerait indéfiniment. En gardant une base fixe, l'affichage reste
    exact et les corrections sont conservées via l'état interne du widget."""
    if st.session_state.donnees_filtrees and i in st.session_state.donnees_filtrees:
        return st.session_state.donnees_filtrees[i][champ]
    return []


def obtenir_donnees_semaine(i: int, champ: str) -> list:
    """Renvoie les données de la semaine i telles qu'elles serviront au bilan :
    la version corrigée par l'utilisateur si la page a déjà été ouverte, sinon
    les données issues de l'analyse (non modifiées)."""
    editees = st.session_state.donnees_editees.get(i)
    if editees is not None:
        return editees[champ]
    return donnees_initiales_semaine(i, champ)


def afficher_bilan_mensuel(rapport: dict) -> None:
    """Affiche le détail d'un bilan mensuel (métriques + dépenses par titre)."""
    nom_mois = NOMS_MOIS[rapport["mois"]]
    st.markdown(f"### 🎉 Bilan du mois de {nom_mois} {rapport['annee']}")

    c1, c2, c3 = st.columns(3)
    c1.metric("Jours travaillés", rapport["jours_travailles"])
    c2.metric("Recette totale", f"{formater_montant(rapport['recette_totale'])} FCFA")
    c3.metric("Solde net", f"{formater_montant(rapport['solde_net'])} FCFA")
    st.metric("Total dépenses", f"{formater_montant(rapport['total_depenses'])} FCFA")

    st.subheader("📋 Détail des dépenses du mois")
    if rapport["depenses_par_titre"]:
        tableau_depenses = [
            {"Dépense": titre, "Montant total": f"{formater_montant(montant)} FCFA"}
            for titre, montant in sorted(rapport["depenses_par_titre"].items(), key=lambda x: -x[1])
        ]
        tableau_depenses.append(
            {"Dépense": "TOTAL", "Montant total": f"{formater_montant(rapport['total_depenses'])} FCFA"}
        )
        st.table(tableau_depenses)
    else:
        st.caption("Aucune dépense enregistrée ce mois-ci.")


# ============================================================
# BARRE LATÉRALE
# ============================================================
with st.sidebar:
    st.markdown("## 🚖 Taxi Dashboard")
    st.caption("Automatisation des rapports hebdomadaires")

    heure = datetime.datetime.now().hour
    if heure < 5:
        salutation = "Bonne nuit"
    elif heure < 12:
        salutation = "Bonjour"
    elif heure < 18:
        salutation = "Bon après-midi"
    else:
        salutation = "Bonsoir"
    st.markdown(f"### {salutation}, {st.session_state.config['nom_utilisateur']} 👋")

    st.divider()
    st.session_state.page = st.radio(
        "Navigation",
        ["📤 Nouveau rapport", "🔃 Maintenance", "⚙️ Paramètres"],
        label_visibility="collapsed",
    )
    st.divider()

    st.caption("📧 Compte de service Google de l'app :")
    st.code(obtenir_email_service_account(), language=None)
    st.caption(
        "C'est avec cette identité que l'app accède aux fichiers. Le client doit partager "
        "sa feuille avec cette adresse en rôle **Éditeur** (détails dans ⚙️ Paramètres)."
    )


# ============================================================
# PAGE : PARAMÈTRES
# ============================================================
if st.session_state.page == "⚙️ Paramètres":
    st.title("⚙️ Paramètres")

    st.subheader("Préférences")
    nouveau_nom = st.text_input("Ton prénom (utilisé dans la salutation)", value=st.session_state.config["nom_utilisateur"])

    st.subheader("Fichier Google Sheets du client")
    st.caption(
        "Colle l'URL complète du fichier Google Sheets partagé par le client, ou juste "
        "son identifiant (la partie entre `/d/` et `/edit` dans l'URL). Pense à vérifier "
        "que le compte de service ci-contre y a bien un accès en écriture."
    )
    nouveau_principale = st.text_input(
        "Classeur du client",
        value=st.session_state.config["sheet_principale_id"],
    )
    nouvel_onglet_depenses = st.text_input(
        "Nom de l'onglet des dépenses",
        value=st.session_state.config.get("nom_onglet_depenses", NOM_ONGLET_DEPENSES_DEFAUT),
        help="Doit correspondre exactement au nom de l'onglet dans Google Sheets (majuscules comprises).",
    )

    if st.button("💾 Enregistrer les paramètres", type="primary"):
        st.session_state.config = {
            "nom_utilisateur": nouveau_nom.strip() or "Pascal",
            "sheet_principale_id": extraire_id_depuis_url(nouveau_principale),
            "nom_onglet_depenses": nouvel_onglet_depenses.strip() or NOM_ONGLET_DEPENSES_DEFAUT,
        }
        sauvegarder_config(st.session_state.config)
        st.session_state.structure_depenses = None
        st.success("✅ Paramètres enregistrés. Ils seront utilisés pour tous les prochains rapports.")

    st.divider()
    st.subheader("🔐 Accès à la feuille du client")
    email_service = obtenir_email_service_account()
    st.write(
        "L'application ne se connecte pas avec ton compte Google personnel, mais avec un "
        "**compte de service** : une identité Google dédiée au robot. Pour qu'elle puisse "
        "écrire dans la feuille, le client doit la partager avec cette adresse, "
        "exactement comme il partagerait avec un collègue."
    )
    st.code(email_service, language=None)

    with st.expander("📋 Marche à suivre à envoyer au client"):
        st.markdown(
            f"""
1. Ouvrir le fichier Google Sheets.
2. Cliquer sur **Partager** (en haut à droite).
3. Coller cette adresse dans le champ des destinataires :
   `{email_service}`
4. Choisir le rôle **Éditeur** (et non « Lecteur »).
5. Décocher « Envoyer une notification » si proposé, puis cliquer sur **Envoyer**.
"""
        )
        st.caption(
            "Un partage « Tout utilisateur disposant du lien » fonctionne aussi, mais rend le "
            "fichier accessible à quiconque possède l'URL : le partage nominatif ci-dessus est "
            "nettement plus sûr pour les données financières du client."
        )

    if st.button("🔌 Vérifier l'accès et la structure du classeur", type="primary"):
        with st.spinner("Vérification en cours..."):
            sheet_id = st.session_state.config["sheet_principale_id"]
            try:
                diagnostic = diagnostiquer_acces_feuille(sheet_id)
            except Exception as e:
                diagnostic = {"statut": "erreur", "message": str(e)}

            if diagnostic["statut"] == "editeur":
                st.success(f"✅ Classeur « {diagnostic['titre']} » : accès en écriture confirmé.")
            elif diagnostic["statut"] == "lecture_seule":
                st.error(f"🚨 {diagnostic['message']}")
                st.info("Demande au client de repasser le partage de « Lecteur » à « Éditeur ».")
            else:
                st.error(f"🚨 {diagnostic['message']}")
                st.info("Vérifie l'identifiant du classeur ci-dessus, puis le partage avec le compte de service.")

            if diagnostic["statut"] in ("editeur", "lecture_seule"):
                try:
                    classeur = get_gspread_client().open_by_key(sheet_id)
                    titres = [f.title for f in classeur.worksheets()]
                    st.write("**Onglets trouvés :** " + ", ".join(f"`{t}`" for t in titres))

                    onglet_attendu = st.session_state.config.get("nom_onglet_depenses", NOM_ONGLET_DEPENSES_DEFAUT)
                    if onglet_attendu in titres:
                        st.success(f"✅ L'onglet des dépenses « {onglet_attendu} » est bien présent.")
                    else:
                        st.error(
                            f"🚨 L'onglet des dépenses « {onglet_attendu} » est introuvable. "
                            "Corrige son nom ci-dessus (orthographe et majuscules doivent correspondre exactement)."
                        )
                except Exception as e:
                    st.error(f"🚨 Lecture des onglets impossible : {message_erreur_sheets(e)}")

    st.divider()
    st.subheader("Quelle clé Gemini l'app utilise-t-elle réellement ?")
    st.caption(
        "Utile après un changement de clé API : confirme que l'app a bien pris en compte "
        "la nouvelle clé (sans jamais afficher la clé en entier)."
    )
    if st.button("🔑 Afficher les derniers caractères de la clé chargée"):
        try:
            cle_active = st.secrets["GEMINI_API_KEY"]
            masque = f"{cle_active[:6]}...{cle_active[-4:]} (longueur : {len(cle_active)} caractères)"
            st.info(f"Clé actuellement chargée par l'app : `{masque}`")
            st.caption(
                "Compare ces derniers caractères avec ceux de ta nouvelle clé dans Google AI Studio. "
                "S'ils ne correspondent pas, l'app utilise encore l'ancienne clé : il faut mettre à jour "
                "le secret puis redémarrer/redéployer l'application."
            )
        except Exception as e:
            st.error(f"🚨 Impossible de lire la clé configurée : {e}")

# ============================================================
# PAGE : MAINTENANCE
# ============================================================
elif st.session_state.page == "🔃 Maintenance":
    st.title("🔃 Maintenance")
    st.caption(
        "Depuis le passage à une feuille Google Sheets unique, il n'y a plus de "
        "tri ni de nettoyage de lignes à faire ici : seul le bilan mensuel (une "
        "colonne par mois) est écrit, directement depuis la page 📊 Bilan mensuel."
    )

    st.divider()
    st.subheader("En cas de problème")
    st.caption(
        "Si tu obtiens une erreur du type « No grid with id » ou une erreur de "
        "connexion inhabituelle après avoir modifié la structure de ton fichier "
        "Google Sheets, clique ici."
    )
    if st.button("🔄 Forcer une reconnexion complète"):
        get_gspread_client.clear()
        st.success("✅ Connexion réinitialisée.")

# ============================================================
# PAGE : NOUVEAU RAPPORT
# ============================================================
else:
    st.title("📤 Nouveau rapport")
    st.write(f"Uploade exactement {MAX_IMAGES_PAR_LOT} rapports hebdomadaires (un mois complet, glisser-déposer possible).")

    fichiers = st.file_uploader(
        f"Sélectionner les {MAX_IMAGES_PAR_LOT} images des rapports",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.cle_uploader}",
    )

    if fichiers:
        if len(fichiers) > MAX_IMAGES_PAR_LOT:
            st.warning(f"⚠️ Maximum {MAX_IMAGES_PAR_LOT} images à la fois. Seules les {MAX_IMAGES_PAR_LOT} premières seront prises en compte.")
            fichiers = fichiers[:MAX_IMAGES_PAR_LOT]

        if len(fichiers) < MAX_IMAGES_PAR_LOT:
            st.info(
                f"📸 L'analyse ne se lance qu'à partir d'un lot complet de {MAX_IMAGES_PAR_LOT} rapports "
                f"hebdomadaires (un mois entier). Il t'en manque {MAX_IMAGES_PAR_LOT - len(fichiers)}."
            )
            st.stop()

        signature = tuple((f.name, f.size) for f in fichiers)
        if st.session_state.signature_lot != signature:
            for cle in list(st.session_state.keys()):
                if cle.startswith("editeur_recettes_") or cle.startswith("editeur_depenses_"):
                    del st.session_state[cle]
            st.session_state.signature_lot = signature
            st.session_state.lot_donnees = None
            st.session_state.donnees_filtrees = None
            st.session_state.mois_confirme = False
            st.session_state.sous_page = 0
            st.session_state.bilan_enregistre = False
            st.session_state.trous_ignores = set()
            st.session_state.fichiers_combles = {}
            st.session_state.ordre_chronologique = None
            st.session_state.ordre_upload_estime = estimer_ordre_upload(fichiers)
            st.session_state.semaines_validees = set()
            st.session_state.bilan_etabli = False
            st.session_state.rapport_fige = None
            st.session_state.donnees_editees = {}
            st.session_state.structure_depenses = None
            st.session_state.depenses_ecrites = False

        total_fichiers = len(fichiers)
        fichiers_effectifs = [fichiers[j] for j in st.session_state.ordre_upload_estime]

        if st.session_state.lot_donnees is not None:
            # L'analyse a déjà eu lieu : on retrouve les fichiers dans l'ordre
            # chronologique déterminé automatiquement (+ éventuelles photos
            # ajoutées pour combler des trous).
            fichiers = construire_fichiers_ordonnes(fichiers_effectifs)
        else:
            # Avant analyse : ordre d'upload brut (l'ordre chronologique n'est pas encore connu).
            fichiers = fichiers_effectifs

        # ============================================================
        # ÉTAPE 1 : analyse groupée des 5 rapports (Gemini)
        # ============================================================
        if st.session_state.lot_donnees is None:
            st.divider()
            st.subheader("🗓️ Étape 1 — Analyse du lot mensuel")
            st.write("Les 5 rapports vont être lus par l'IA afin d'identifier le mois couvert, avant tout enregistrement.")
            st.caption(
                f"Modèle utilisé : `{GEMINI_MODEL}`. Les 5 images sont envoyées avec quelques secondes "
                "d'écart entre chacune pour rester dans les limites du palier gratuit Gemini."
            )

            if st.button("🔍 Analyser le lot (5 rapports)", type="primary"):
                with st.spinner("Analyse Gemini des 5 rapports en cours..."):
                    try:
                        api_key, _ = get_clients_config()
                        resultats_lot = []
                        for idx_f, f in enumerate(fichiers):
                            try:
                                if idx_f > 0:
                                    time.sleep(6)  # laisse respirer le quota requêtes/minute entre 2 images (palier gratuit)
                                image_b64, mime_type = encoder_image(f)
                                res_json = appeler_gemini(api_key, image_b64, mime_type)
                                donnees = extraire_donnees(res_json)

                                if not donnees.get("periode_hebdo"):
                                    jours_dates = [
                                        parser_date(str(j.get("date", "")))
                                        for j in donnees.get("recettes_journalieres", [])
                                    ]
                                    jours_dates = [d for d in jours_dates if d]
                                    if jours_dates:
                                        donnees["periode_hebdo"] = (
                                            f"{min(jours_dates).strftime('%d/%m/%y')} - "
                                            f"{max(jours_dates).strftime('%d/%m/%y')}"
                                        )
                                    else:
                                        donnees["periode_hebdo"] = periode_semaine_courante()
                                resultats_lot.append(donnees)
                            except Exception as e:
                                resultats_lot.append({"erreur": masquer_cle_api(str(e))})

                        ordre = calculer_ordre_chronologique(resultats_lot)
                        st.session_state.ordre_chronologique = ordre
                        st.session_state.lot_donnees = [resultats_lot[j] for j in ordre]
                        st.rerun()
                    except Exception as e:
                        st.error(f"🚨 Erreur lors de la connexion : {masquer_cle_api(str(e))}")
            st.stop()

        # ============================================================
        # ÉTAPE 2 : confirmation du mois + détection des anomalies
        # (doublons) + gestion interactive des périodes manquantes
        # ============================================================
        if not st.session_state.mois_confirme:
            st.divider()
            st.subheader("🗓️ Étape 2 — Vérification du mois avant enregistrement")

            nb_erreurs = sum(1 for d in st.session_state.lot_donnees if "erreur" in d)
            if nb_erreurs:
                st.warning(
                    f"⚠️ {nb_erreurs} rapport(s) sur {total_fichiers} n'ont pas pu être lus par l'IA. "
                    "Ils seront à passer ou traiter manuellement lors de la revue à l'étape suivante."
                )

            debut_mois, fin_mois = determiner_bornes_mois(st.session_state.lot_donnees)

            if debut_mois is None:
                st.error(
                    "🚨 Impossible de déterminer le mois couvert : aucune date exploitable n'a été "
                    "trouvée dans les 5 rapports."
                )
                with st.expander("🔍 Voir les données brutes reçues (diagnostic)"):
                    for i, d in enumerate(st.session_state.lot_donnees, start=1):
                        st.write(f"**Rapport {i}**")
                        st.json(d)
                if st.button("↩️ Recommencer l'analyse du lot"):
                    st.session_state.lot_donnees = None
                    st.session_state.ordre_chronologique = None
                    st.rerun()
                st.stop()

            nom_utilisateur = st.session_state.config["nom_utilisateur"]
            nom_mois = NOMS_MOIS[debut_mois.month]
            st.markdown(
                f"**{nom_utilisateur}, ce rapport mensuel semble couvrir le mois de "
                f"{nom_mois} {debut_mois.year}, du {debut_mois.strftime('%d/%m/%y')} "
                f"au {fin_mois.strftime('%d/%m/%y')}.**"
            )
            remarque = st.text_area(
                "Si tu y vois un défaut, communique-le ici (sinon laisse vide) :",
                key="remarque_mois",
            )
            if remarque.strip():
                st.info("📝 Remarque notée : tu pourras corriger chaque date lors de la revue détaillée, rapport par rapport.")

            # --------------------------------------------------------
            # 1) Détection d'anomalies : doublons (fichier ou période)
            # --------------------------------------------------------
            st.divider()
            st.markdown("#### 🔎 Vérification des anomalies")
            alertes_doublons = detecter_doublons(fichiers, st.session_state.lot_donnees)
            doublons_bloquants = [a for a in alertes_doublons if a["gravite"] == "erreur"]

            if alertes_doublons:
                for alerte in alertes_doublons:
                    if alerte["gravite"] == "erreur":
                        st.error(f"🚨 {alerte['message']}")
                    else:
                        st.warning(f"⚠️ {alerte['message']}")
            else:
                st.success("✅ Aucun doublon détecté (fichiers et périodes tous distincts).")

            poursuivre_malgre_doublon = True
            if doublons_bloquants:
                poursuivre_malgre_doublon = st.checkbox(
                    "Je confirme que ce n'est pas une erreur : poursuivre quand même malgré le(s) doublon(s) ci-dessus.",
                    key="confirmer_doublon",
                )

            # --------------------------------------------------------
            # 2) Périodes manquantes : combler avec une photo, ou ignorer
            # --------------------------------------------------------
            st.divider()
            st.markdown("#### 🗓️ Vérification des périodes manquantes")
            trous = detecter_trous_hebdomadaires(st.session_state.lot_donnees)
            trous_a_traiter = [
                t for t in trous
                if f"{t[0].isoformat()}_{t[1].isoformat()}" not in st.session_state.trous_ignores
            ]

            if trous_a_traiter:
                for debut_trou, fin_trou in trous_a_traiter:
                    cle_trou = f"{debut_trou.isoformat()}_{fin_trou.isoformat()}"
                    st.warning(
                        f"⚠️ La période du **{debut_trou.strftime('%d/%m/%y')}** au "
                        f"**{fin_trou.strftime('%d/%m/%y')}** manque."
                    )
                    with st.expander(f"Gérer la période manquante du {debut_trou.strftime('%d/%m/%y')} au {fin_trou.strftime('%d/%m/%y')}"):
                        fichier_combler = st.file_uploader(
                            "Ajoute la photo de cette semaine si tu l'as (elle sera vérifiée automatiquement)",
                            type=["jpg", "jpeg", "png"],
                            key=f"upload_comblement_{cle_trou}",
                        )
                        col_verif, col_ignorer = st.columns(2)
                        with col_verif:
                            if st.button(
                                "🔍 Vérifier et ajouter cette photo",
                                key=f"verifier_comblement_{cle_trou}",
                                disabled=fichier_combler is None,
                                use_container_width=True,
                            ):
                                with st.spinner("Analyse de la photo ajoutée..."):
                                    try:
                                        api_key, _ = get_clients_config()
                                        image_b64, mime_type = encoder_image(fichier_combler)
                                        res_json = appeler_gemini(api_key, image_b64, mime_type)
                                        donnees_comblees = extraire_donnees(res_json)
                                        p = parser_periode(donnees_comblees.get("periode_hebdo", ""))
                                        chevauche = p and not (p[1] < debut_trou or p[0] > fin_trou)
                                        if chevauche:
                                            st.session_state.lot_donnees.append(donnees_comblees)
                                            st.session_state.fichiers_combles[cle_trou] = fichier_combler
                                            st.success("✅ Photo ajoutée : elle correspond bien à la période manquante.")
                                            st.rerun()
                                        else:
                                            periode_trouvee = donnees_comblees.get("periode_hebdo") or "non détectée"
                                            st.error(
                                                f"🚨 Cette photo couvre la période « {periode_trouvee} », qui ne "
                                                f"correspond pas à la période manquante ({debut_trou.strftime('%d/%m/%y')} - "
                                                f"{fin_trou.strftime('%d/%m/%y')}). Vérifie que c'est la bonne photo."
                                            )
                                    except Exception as e:
                                        st.error(f"🚨 Erreur lors de l'analyse : {masquer_cle_api(str(e))}")
                        with col_ignorer:
                            if st.button(
                                "⏭️ Enregistrer sans cette période",
                                key=f"ignorer_comblement_{cle_trou}",
                                use_container_width=True,
                            ):
                                st.session_state.trous_ignores.add(cle_trou)
                                st.rerun()
                st.caption(
                    "Pour chaque période manquante ci-dessus : ajoute la photo correspondante (elle sera "
                    "vérifiée automatiquement avant d'être intégrée), ou choisis d'enregistrer sans elle."
                )
            else:
                st.success("✅ Aucune période manquante en attente (comblée ou ignorée).")

            tous_les_trous_geres = len(trous_a_traiter) == 0

            col_go, col_retour = st.columns(2)
            with col_go:
                if st.button(
                    "➡️ Continuer vers la vérification des rapports",
                    type="primary",
                    disabled=not (tous_les_trous_geres and poursuivre_malgre_doublon),
                ):
                    for cle in list(st.session_state.keys()):
                        if cle.startswith("editeur_recettes_") or cle.startswith("editeur_depenses_"):
                            del st.session_state[cle]
                    filtres = {}
                    for i, donnees in enumerate(st.session_state.lot_donnees):
                        if "erreur" in donnees:
                            continue
                        filtre = filtrer_donnees_par_mois(donnees, debut_mois, fin_mois)
                        filtres[i] = {
                            "recettes": filtre["recettes_journalieres"],
                            "depenses": filtre["depenses"],
                        }
                    st.session_state.donnees_filtrees = filtres
                    st.session_state.mois_confirme = True
                    st.session_state.sous_page = 0
                    st.rerun()
            with col_retour:
                if st.button("↩️ Recommencer l'analyse du lot"):
                    st.session_state.lot_donnees = None
                    st.session_state.trous_ignores = set()
                    st.session_state.fichiers_combles = {}
                    st.session_state.ordre_chronologique = None
                    st.rerun()
            st.stop()

        # ============================================================
        # ÉTAPE 3 : pages navigables — une par semaine (filtrée sur le
        # mois confirmé), puis le bilan mensuel consolidé
        # ============================================================
        debut_mois, fin_mois = determiner_bornes_mois(st.session_state.lot_donnees)
        tous_les_fichiers = list(fichiers)
        nb_rapports_total = len(st.session_state.lot_donnees)

        def libelle_semaine(i: int) -> str:
            if "erreur" in st.session_state.lot_donnees[i]:
                return f"Semaine {i + 1} ⚠️"
            return f"Semaine {i + 1}" + (" ✅" if i in st.session_state.semaines_validees else "")

        noms_pages = [libelle_semaine(i) for i in range(nb_rapports_total)] + ["📊 Bilan mensuel"]
        nb_pages = len(noms_pages)

        st.divider()

        nav1, nav2, nav3 = st.columns([1, 3, 1])
        with nav1:
            if st.button("◀ Précédent", disabled=st.session_state.sous_page == 0, use_container_width=True, key="nav_prec_semaine"):
                st.session_state.sous_page -= 1
                st.rerun()
        with nav2:
            # Clé dynamique (inclut sous_page) : le widget se recrée à chaque
            # changement de page et respecte alors 'index=', au lieu de garder
            # un état figé qui écraserait la navigation par boutons.
            choix_page = st.selectbox(
                "Aller à la page",
                options=list(range(nb_pages)),
                format_func=lambda i: noms_pages[i],
                index=st.session_state.sous_page,
                label_visibility="collapsed",
                key=f"select_nav_semaine_{st.session_state.sous_page}",
            )
            if choix_page != st.session_state.sous_page:
                st.session_state.sous_page = choix_page
                st.rerun()
        with nav3:
            if st.button("Suivant ▶", disabled=st.session_state.sous_page == nb_pages - 1, use_container_width=True, key="nav_suiv_semaine"):
                st.session_state.sous_page += 1
                st.rerun()

        page_actuelle = st.session_state.sous_page

        # --------------------------------------------------------
        # PAGES 0..N-1 : une semaine, filtrée sur le mois confirmé
        # --------------------------------------------------------
        if page_actuelle < nb_rapports_total:
            i = page_actuelle
            fichier_i = tous_les_fichiers[i]
            donnees_brutes = st.session_state.lot_donnees[i]

            st.divider()
            st.subheader(f"📄 Étape 3 — Vérification : Semaine {i + 1} / {nb_rapports_total}")

            if "erreur" in donnees_brutes:
                st.error(f"🚨 Ce rapport n'a pas pu être analysé : {donnees_brutes['erreur']}")
                st.caption("Il est exclu du bilan mensuel. Tu peux relancer l'analyse du lot si besoin (bouton à l'étape précédente).")
            else:
                semaine_validee = i in st.session_state.semaines_validees

                if st.session_state.bilan_etabli:
                    st.info(
                        "🔒 Le bilan mensuel a déjà été établi : les données sont figées. "
                        "Pour les modifier, utilise « Reprendre les modifications » sur la page 📊 Bilan mensuel."
                    )
                elif semaine_validee:
                    st.success("✅ Semaine validée — les données ci-dessous sont prêtes pour le bilan.")

                col_img, col_data = st.columns([1, 1.4])
                with col_img:
                    st.image(fichier_i, caption=fichier_i.name, use_container_width=True)

                with col_data:
                    st.caption(
                        f"Période brute détectée sur l'image : **{donnees_brutes.get('periode_hebdo', '—')}**. "
                        f"Seuls les jours du mois confirmé ({debut_mois.strftime('%d/%m/%y')} - "
                        f"{fin_mois.strftime('%d/%m/%y')}) sont pris en compte ci-dessous."
                    )
                    st.caption(
                        "🖊️ Corrige librement ce que l'IA a mal lu : tu peux modifier une date ou un "
                        "montant, ajouter une ligne oubliée, ou supprimer une ligne en trop."
                    )

                    lecture_seule = st.session_state.bilan_etabli

                    st.write("**Recettes journalières (mois confirmé uniquement)**")
                    recettes_editees = st.data_editor(
                        donnees_initiales_semaine(i, "recettes"),
                        num_rows="dynamic",
                        column_config={
                            "date": st.column_config.TextColumn("Date (JJ/MM/AA)", required=True),
                            "montant": st.column_config.NumberColumn("Montant (FCFA)", required=True, step=500),
                        },
                        use_container_width=True,
                        disabled=lecture_seule,
                        key=f"editeur_recettes_{i}",
                    )

                    st.write("**Dépenses (mois confirmé uniquement)**")
                    depenses_editees = st.data_editor(
                        donnees_initiales_semaine(i, "depenses"),
                        num_rows="dynamic",
                        column_config={
                            "titre": st.column_config.TextColumn("Titre de la dépense", required=True),
                            "montant": st.column_config.NumberColumn("Montant (FCFA)", required=True, step=500),
                            "date": st.column_config.TextColumn("Date (JJ/MM/AA)", required=False),
                        },
                        use_container_width=True,
                        disabled=lecture_seule,
                        key=f"editeur_depenses_{i}",
                    )

                recettes_editees = normaliser_lignes_editeur(recettes_editees)
                depenses_editees = normaliser_lignes_editeur(depenses_editees)
                memoriser_donnees_semaine(i, recettes_editees, depenses_editees)

                if not recettes_editees:
                    st.info("ℹ️ Aucun jour de cette semaine n'appartient au mois confirmé.")

                total_r = calculer_total_recettes(recettes_editees)
                total_d = calculer_total_depenses(depenses_editees)
                st.divider()
                m1, m2, m3 = st.columns(3)
                m1.metric("Recettes (mois confirmé)", f"{formater_montant(total_r)} FCFA")
                m2.metric("Dépenses (mois confirmé)", f"{formater_montant(total_d)} FCFA")
                m3.metric("Solde", f"{formater_montant(total_r - total_d)} FCFA")

                if not st.session_state.bilan_etabli:
                    st.divider()
                    if semaine_validee:
                        if st.button("🖊️ Modifier à nouveau cette semaine", key=f"devalider_{i}", use_container_width=True):
                            st.session_state.semaines_validees.discard(i)
                            st.rerun()
                    else:
                        if st.button(
                            "✅ Valider cette semaine (aucune modification à apporter)",
                            type="primary",
                            use_container_width=True,
                            key=f"valider_semaine_{i}",
                        ):
                            st.session_state.semaines_validees.add(i)
                            if i < nb_rapports_total - 1:
                                st.session_state.sous_page = i + 1
                            else:
                                st.session_state.sous_page = nb_pages - 1
                            st.rerun()

        # --------------------------------------------------------
        # PAGE FINALE : bilan mensuel consolidé + enregistrement
        # --------------------------------------------------------
        else:
            st.divider()

            indices_exploitables = [
                i for i in range(nb_rapports_total)
                if "erreur" not in st.session_state.lot_donnees[i]
            ]
            non_validees = [i for i in indices_exploitables if i not in st.session_state.semaines_validees]

            # ====================================================
            # PORTE DE CONFIRMATION : tant que l'utilisateur n'a pas
            # confirmé ses modifications, aucun bilan n'est établi.
            # ====================================================
            if not st.session_state.bilan_etabli:
                st.subheader("🧮 Étape 4 — Établir le bilan mensuel")
                st.write(
                    "Le bilan n'est pas encore calculé : il le sera à partir des données que tu "
                    "auras validées, corrections comprises."
                )

                st.markdown("**État de la vérification, semaine par semaine :**")
                etat_semaines = []
                for i in range(nb_rapports_total):
                    if "erreur" in st.session_state.lot_donnees[i]:
                        statut = "⚠️ Non analysée (exclue du bilan)"
                    elif i in st.session_state.semaines_validees:
                        statut = "✅ Validée"
                    else:
                        statut = "⏳ En attente de vérification"
                    etat_semaines.append({"Semaine": f"Semaine {i + 1}", "État": statut})
                st.table(etat_semaines)

                if non_validees:
                    liste = ", ".join(f"Semaine {i + 1}" for i in non_validees)
                    st.warning(
                        f"⏳ Il reste à vérifier : {liste}. Ouvre chaque semaine pour corriger si "
                        "nécessaire, puis valide-la."
                    )
                    if st.button("⏭️ Aller à la première semaine à vérifier", use_container_width=True):
                        st.session_state.sous_page = non_validees[0]
                        st.rerun()
                    st.caption(
                        "Si tu as déjà tout relu et qu'aucune correction n'est nécessaire, tu peux "
                        "tout valider d'un coup ci-dessous."
                    )
                    if st.button("✅ Tout valider sans modification", use_container_width=True):
                        st.session_state.semaines_validees = set(indices_exploitables)
                        st.rerun()
                else:
                    st.success("✅ Toutes les semaines exploitables ont été vérifiées et validées.")

                st.divider()
                if st.button(
                    "🚀 Lancer l'analyse et établir le bilan mensuel",
                    type="primary",
                    use_container_width=True,
                    disabled=bool(non_validees),
                    help="Valide d'abord chaque semaine." if non_validees else None,
                ):
                    recettes_combinees = []
                    depenses_combinees = []
                    for i in indices_exploitables:
                        recettes_combinees.extend(obtenir_donnees_semaine(i, "recettes"))
                        depenses_combinees.extend(obtenir_donnees_semaine(i, "depenses"))

                    # Le rapport est FIGÉ : une modification ultérieure d'un
                    # éditeur ne peut plus altérer silencieusement un bilan déjà
                    # établi (et potentiellement déjà écrit dans Google Sheets).
                    st.session_state.rapport_fige = calculer_bilan_mensuel_agrege(
                        recettes_combinees, depenses_combinees, debut_mois.year, debut_mois.month
                    )
                    st.session_state.bilan_etabli = True
                    st.rerun()
                st.stop()

            # ====================================================
            # BILAN ÉTABLI : affichage + enregistrement
            # ====================================================
            rapport = st.session_state.rapport_fige
            afficher_bilan_mensuel(rapport)

            with st.expander("🖊️ Besoin de corriger encore une donnée ?"):
                st.caption(
                    "Reprendre les modifications déverrouille les tableaux hebdomadaires et annule "
                    "le bilan actuel. Il faudra le rétablir ensuite. "
                    "Les écritures déjà effectuées dans Google Sheets ne sont PAS annulées."
                )
                if st.button("↩️ Reprendre les modifications"):
                    st.session_state.bilan_etabli = False
                    st.session_state.rapport_fige = None
                    st.session_state.semaines_validees = set()
                    st.session_state.structure_depenses = None
                    st.session_state.depenses_ecrites = False
                    st.session_state.bilan_enregistre = False
                    st.session_state.sous_page = 0
                    st.rerun()

            # ====================================================
            # ÉCRITURE DES DÉPENSES DANS L'ONGLET "EXPENSES" DU CLIENT
            # ====================================================
            st.divider()
            st.markdown("### 💸 Enregistrer les dépenses dans la feuille du client")
            nom_onglet_dep = st.session_state.config.get("nom_onglet_depenses", NOM_ONGLET_DEPENSES_DEFAUT)
            nom_mois_rapport = NOMS_MOIS[rapport["mois"]]
            st.caption(
                f"Onglet ciblé : **{nom_onglet_dep}** — colonne du mois de "
                f"**{nom_mois_rapport} {rapport['annee']}** "
                f"(colonne {numero_colonne_vers_lettre(rapport['mois'] + OFFSET_COLONNE_MOIS)})."
            )

            if not rapport["depenses_par_titre"]:
                st.info("ℹ️ Aucune dépense à enregistrer pour ce mois.")

            elif st.session_state.structure_depenses is None:
                st.write(
                    "L'app va d'abord lire les postes de dépense déjà présents dans la feuille du client, "
                    "puis te proposer un rapprochement automatique que tu pourras corriger avant écriture."
                )
                if st.button("🔗 Lire l'onglet Expenses du client", type="primary"):
                    with st.spinner("Lecture de la feuille du client en cours..."):
                        try:
                            feuille_dep = get_onglet(nom_onglet_dep)
                            st.session_state.structure_depenses = lire_structure_depenses(feuille_dep)
                            st.rerun()
                        except Exception as e:
                            st.error(f"🚨 {message_erreur_sheets(e)}")

            else:
                structure = st.session_state.structure_depenses
                postes = lister_postes(structure)

                if not postes:
                    st.error(
                        f"🚨 Aucun poste de dépense n'a été trouvé dans l'onglet « {nom_onglet_dep} ». "
                        "Vérifie que les intitulés sont bien en colonne C et que les lignes de totaux "
                        "contiennent « Monthly totals: »."
                    )
                    if st.button("🔄 Relire l'onglet"):
                        st.session_state.structure_depenses = None
                        st.rerun()
                    st.stop()

                nb_rubriques = len(structure["rubriques"])
                st.success(f"✅ Feuille lue : {nb_rubriques} rubrique(s), {len(postes)} poste(s) de dépense disponibles.")

                libelle_par_ligne = {
                    p["ligne"]: f"{p['rubrique']} › {p['nom']}   (ligne {p['ligne']})"
                    for p in postes
                }
                options_lignes = [None] + [p["ligne"] for p in postes]

                st.markdown("#### Rapprochement des dépenses")
                st.caption(
                    "Vérifie la destination proposée pour chaque dépense. Les postes de rubriques "
                    "récurrentes (dépenses fixes, dépenses diverses) sont privilégiés automatiquement."
                )

                choix_par_titre = {}
                for titre, montant in sorted(rapport["depenses_par_titre"].items(), key=lambda x: -x[1]):
                    auto = trouver_appariement_auto(titre, postes)
                    index_defaut = options_lignes.index(auto["ligne"]) if auto else 0
                    marqueur = "✅" if auto else "❓"
                    choix_par_titre[titre] = st.selectbox(
                        f"{marqueur} **{titre}** — {formater_montant(montant)} FCFA",
                        options=options_lignes,
                        format_func=lambda v: "⏭️ Ne pas enregistrer" if v is None else libelle_par_ligne[v],
                        index=index_defaut,
                        key=f"dest_depense_{normaliser_libelle(titre)}",
                    )

                # Plusieurs dépenses peuvent viser le même poste : on les cumule.
                montants_par_ligne: dict = {}
                titres_par_ligne: dict = {}
                for titre, ligne_cible in choix_par_titre.items():
                    if ligne_cible is None:
                        continue
                    montants_par_ligne[ligne_cible] = montants_par_ligne.get(ligne_cible, 0.0) + rapport["depenses_par_titre"][titre]
                    titres_par_ligne.setdefault(ligne_cible, []).append(titre)

                st.markdown("#### Aperçu avant écriture")
                if not montants_par_ligne:
                    st.warning("⚠️ Aucune dépense n'est actuellement destinée à être écrite.")
                else:
                    apercu = []
                    ecrasements = 0
                    for ligne_cible, montant in sorted(montants_par_ligne.items()):
                        actuelle = valeur_actuelle_cellule(structure, ligne_cible, rapport["mois"])
                        if actuelle:
                            ecrasements += 1
                        apercu.append({
                            "Destination": libelle_par_ligne[ligne_cible],
                            "Dépense(s)": ", ".join(titres_par_ligne[ligne_cible]),
                            "Montant à écrire": f"{formater_montant(montant)} FCFA",
                            "Valeur actuelle": actuelle or "(vide)",
                        })
                    st.table(apercu)

                    total_ecrit = sum(montants_par_ligne.values())
                    non_affectees = rapport["total_depenses"] - total_ecrit
                    c1, c2 = st.columns(2)
                    c1.metric("Total qui sera écrit", f"{formater_montant(total_ecrit)} FCFA")
                    c2.metric("Non affecté", f"{formater_montant(non_affectees)} FCFA")

                    if non_affectees > 0:
                        st.warning(
                            f"⚠️ {formater_montant(non_affectees)} FCFA de dépenses ne seront pas écrites "
                            "(marquées « Ne pas enregistrer »)."
                        )
                    if ecrasements:
                        st.warning(
                            f"⚠️ {ecrasements} cellule(s) contiennent déjà une valeur pour "
                            f"{nom_mois_rapport} : elle sera remplacée."
                        )

                st.divider()
                if st.session_state.depenses_ecrites:
                    st.success(f"✅ Les dépenses ont été écrites dans l'onglet « {nom_onglet_dep} ».")
                    if st.button("🔄 Relire la feuille et recommencer le rapprochement"):
                        st.session_state.structure_depenses = None
                        st.session_state.depenses_ecrites = False
                        st.rerun()
                else:
                    col_ecrire, col_relire = st.columns(2)
                    with col_ecrire:
                        if st.button(
                            "💾 Écrire les dépenses dans la feuille du client",
                            type="primary",
                            use_container_width=True,
                            disabled=not montants_par_ligne,
                        ):
                            with st.spinner("Écriture dans la feuille du client en cours..."):
                                try:
                                    feuille_dep = get_onglet(nom_onglet_dep)
                                    ecrire_depenses_client(feuille_dep, montants_par_ligne, rapport["mois"])
                                    st.session_state.depenses_ecrites = True
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"🚨 Erreur lors de l'écriture : {message_erreur_sheets(e)}")
                    with col_relire:
                        if st.button("🔄 Relire l'onglet Expenses", use_container_width=True):
                            st.session_state.structure_depenses = None
                            st.rerun()

            st.divider()
            with st.expander("📄 Enregistrer aussi le résumé (recettes / dépenses / solde) sur le premier onglet"):
                st.caption(
                    "Écrit les 3 totaux du mois dans le premier onglet du classeur "
                    f"(lignes {LIGNE_RECETTES_MENSUEL} à {LIGNE_SOLDE_MENSUEL}). "
                    "À n'utiliser que si cet onglet est bien prévu pour ça."
                )
                if st.session_state.bilan_enregistre:
                    st.success("✅ Résumé déjà enregistré.")
                elif st.button("💾 Enregistrer le résumé mensuel"):
                    with st.spinner("Écriture en cours..."):
                        try:
                            _, feuille_principale = get_clients_config()
                            enregistrer_bilan_mensuel(feuille_principale, rapport)
                            st.session_state.bilan_enregistre = True
                            st.success("✨ Résumé mensuel enregistré.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"🚨 Une erreur est survenue lors de l'enregistrement : {message_erreur_sheets(e)}")

            st.divider()
            if st.button("📥 Traiter un nouveau lot d'images (mois suivant)"):
                reinitialiser_lot()
                st.rerun()