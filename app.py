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
import json
import os
import re
import time

import gspread
import requests
import streamlit as st
from google.oauth2.service_account import Credentials

# ============================================================
# CONFIGURATION GÉNÉRALE
# ============================================================
st.set_page_config(page_title="Taxi Dashboard - Automatisation", page_icon="🚖", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Valeur par défaut (utilisée tant que rien n'est configuré dans ⚙️ Paramètres)
SHEET_PRINCIPALE_ID_DEFAUT = "1BNlV17OasazXtFPLbp64xHwJ2RqBjPqs97_Adi4Cbuo"

GEMINI_MODEL = "gemini-2.5-flash"  # à ajuster si le nom du modèle change côté Google AI Studio

CONFIG_PATH = "config_app.json"

NOMS_MOIS = ["", "janvier", "février", "mars", "avril", "mai", "juin", "juillet",
             "août", "septembre", "octobre", "novembre", "décembre"]

MAX_IMAGES_PAR_LOT = 5

# Feuille principale : une colonne par mois (mois + décalage), 3 lignes de bilan
OFFSET_COLONNE_MOIS = 3
LIGNE_RECETTES_MENSUEL = 4
LIGNE_DEPENSES_MENSUEL = 5
LIGNE_SOLDE_MENSUEL = 6

VALUE_INPUT_OPTION = "USER_ENTERED"  # évite que Sheets force les nombres en texte (bug de l'apostrophe)


# ============================================================
# CONFIGURATION UTILISATEUR (persistée sur disque)
# ============================================================
def charger_config() -> dict:
    defaut = {
        "sheet_principale_id": SHEET_PRINCIPALE_ID_DEFAUT,
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
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    """Authentifie le compte de service Google - mis en cache car les
    identifiants ne changent jamais en cours de session."""
    if os.path.exists("credentials.json"):
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(st.secrets["google_credentials"], scopes=SCOPES)
    return gspread.authorize(creds)


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

        try:
            reponse.raise_for_status()
        except requests.HTTPError as exc:
            raise RuntimeError(masquer_cle_api(str(exc))) from exc
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


def obtenir_donnees_semaine(i: int, champ: str) -> list:
    """Renvoie les données (recettes/dépenses) éditées par l'utilisateur pour
    la semaine i si la page a déjà été ouverte (donc le data_editor a déjà
    été instancié), sinon les données filtrées par défaut (non éditées)."""
    cle_widget = f"editeur_{champ}_{i}"
    if cle_widget in st.session_state:
        return st.session_state[cle_widget]
    if st.session_state.donnees_filtrees and i in st.session_state.donnees_filtrees:
        return st.session_state.donnees_filtrees[i][champ]
    return []



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

    st.caption("📧 Compte de service Google connecté :")
    st.code(obtenir_email_service_account(), language=None)
    st.caption("Partage tes fichiers Google Sheets avec cette adresse (rôle **Éditeur**) pour que l'app puisse les modifier.")


# ============================================================
# PAGE : PARAMÈTRES
# ============================================================
if st.session_state.page == "⚙️ Paramètres":
    st.title("⚙️ Paramètres")

    st.subheader("Préférences")
    nouveau_nom = st.text_input("Ton prénom (utilisé dans la salutation)", value=st.session_state.config["nom_utilisateur"])

    st.subheader("Fichier Google Sheets")
    st.caption(
        "Colle l'URL complète du fichier Google Sheets, ou juste son identifiant "
        "(la partie entre `/d/` et `/edit` dans l'URL). C'est la feuille où sont "
        "enregistrés les bilans mensuels (recettes/dépenses/solde), une colonne par mois."
    )
    nouveau_principale = st.text_input(
        "Feuille principale (bilans mensuels)",
        value=st.session_state.config["sheet_principale_id"],
    )

    if st.button("💾 Enregistrer les paramètres", type="primary"):
        st.session_state.config = {
            "nom_utilisateur": nouveau_nom.strip() or "Pascal",
            "sheet_principale_id": extraire_id_depuis_url(nouveau_principale),
        }
        sauvegarder_config(st.session_state.config)
        st.success("✅ Paramètres enregistrés. Ils seront utilisés pour tous les prochains rapports.")

    st.divider()
    st.subheader("Test de connexion")
    if st.button("🔌 Vérifier la connexion au fichier configuré"):
        with st.spinner("Connexion en cours..."):
            try:
                _, feuille_principale = get_clients_config()
                st.success(f"✅ Connecté avec succès à la feuille principale « {feuille_principale.spreadsheet.title} ».")
            except Exception as e:
                st.error(f"🚨 Connexion impossible : {e}")

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

        total_fichiers = len(fichiers)

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
            st.image(list(fichiers), caption=[f.name for f in fichiers], width=130)

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
                        st.session_state.lot_donnees = resultats_lot
                        st.rerun()
                    except Exception as e:
                        st.error(f"🚨 Erreur lors de la connexion : {masquer_cle_api(str(e))}")
            st.stop()

        # ============================================================
        # ÉTAPE 2 : confirmation du mois + détection des trous
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

            trous = detecter_trous_hebdomadaires(st.session_state.lot_donnees)
            poursuivre_malgre_trou = True
            if trous:
                for debut_trou, fin_trou in trous:
                    st.warning(f"⚠️ La période du {debut_trou.strftime('%d/%m/%y')} au {fin_trou.strftime('%d/%m/%y')} manque.")
                st.caption(
                    "Si cette période manque vraiment, ferme cet écran, ajoute la photo correspondante "
                    "à ton prochain lot de 5, puis relance l'analyse. Sinon, tu peux poursuivre quand même."
                )
                poursuivre_malgre_trou = st.checkbox(
                    "Poursuivre quand même malgré la/les période(s) manquante(s) ci-dessus."
                )
            else:
                st.success("✅ Les 5 périodes hebdomadaires s'enchaînent sans trou détecté.")

            col_go, col_retour = st.columns(2)
            with col_go:
                if st.button(
                    "➡️ Continuer vers la vérification des rapports",
                    type="primary",
                    disabled=not poursuivre_malgre_trou,
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
                    st.rerun()
            st.stop()

        # ============================================================
        # ÉTAPE 3 : pages navigables — une par semaine (filtrée sur le
        # mois confirmé), puis le bilan mensuel consolidé
        # ============================================================
        debut_mois, fin_mois = determiner_bornes_mois(st.session_state.lot_donnees)
        noms_pages = [f"Semaine {i + 1}" for i in range(total_fichiers)] + ["📊 Bilan mensuel"]
        nb_pages = len(noms_pages)

        st.divider()
        nav1, nav2, nav3 = st.columns([1, 3, 1])
        with nav1:
            if st.button("◀ Précédent", disabled=st.session_state.sous_page == 0, use_container_width=True):
                st.session_state.sous_page -= 1
                st.rerun()
        with nav2:
            choix_page = st.selectbox(
                "Aller à la page",
                options=list(range(nb_pages)),
                format_func=lambda i: noms_pages[i],
                index=st.session_state.sous_page,
                label_visibility="collapsed",
            )
            if choix_page != st.session_state.sous_page:
                st.session_state.sous_page = choix_page
                st.rerun()
        with nav3:
            if st.button("Suivant ▶", disabled=st.session_state.sous_page == nb_pages - 1, use_container_width=True):
                st.session_state.sous_page += 1
                st.rerun()

        page_actuelle = st.session_state.sous_page

        # --------------------------------------------------------
        # PAGES 0..4 : une semaine, filtrée sur le mois confirmé
        # --------------------------------------------------------
        if page_actuelle < total_fichiers:
            i = page_actuelle
            fichier_i = fichiers[i]
            donnees_brutes = st.session_state.lot_donnees[i]

            st.divider()
            st.subheader(f"📄 Semaine {i + 1} / {total_fichiers}")

            if "erreur" in donnees_brutes:
                st.error(f"🚨 Ce rapport n'a pas pu être analysé : {donnees_brutes['erreur']}")
                st.caption("Il est exclu du bilan mensuel. Tu peux relancer l'analyse du lot si besoin (bouton à l'étape précédente).")
            else:
                col_img, col_data = st.columns([1, 1.4])
                with col_img:
                    st.image(fichier_i, caption=fichier_i.name, use_container_width=True)

                with col_data:
                    st.caption(
                        f"Période brute détectée sur l'image : **{donnees_brutes.get('periode_hebdo', '—')}**. "
                        f"Seuls les jours du mois confirmé ({debut_mois.strftime('%d/%m/%y')} - "
                        f"{fin_mois.strftime('%d/%m/%y')}) sont pris en compte ci-dessous."
                    )

                    st.write("**Recettes journalières (mois confirmé uniquement)**")
                    recettes_editees = st.data_editor(
                        obtenir_donnees_semaine(i, "recettes"),
                        num_rows="dynamic",
                        column_config={
                            "date": st.column_config.TextColumn("Date (JJ/MM/AA)", required=True),
                            "montant": st.column_config.NumberColumn("Montant (FCFA)", required=True, step=500),
                        },
                        use_container_width=True,
                        key=f"editeur_recettes_{i}",
                    )

                    st.write("**Dépenses (mois confirmé uniquement)**")
                    depenses_editees = st.data_editor(
                        obtenir_donnees_semaine(i, "depenses"),
                        num_rows="dynamic",
                        column_config={
                            "titre": st.column_config.TextColumn("Titre de la dépense", required=True),
                            "montant": st.column_config.NumberColumn("Montant (FCFA)", required=True, step=500),
                            "date": st.column_config.TextColumn("Date (JJ/MM/AA)", required=False),
                        },
                        use_container_width=True,
                        key=f"editeur_depenses_{i}",
                    )

                if not recettes_editees:
                    st.info("ℹ️ Aucun jour de cette semaine n'appartient au mois confirmé.")

                total_r = calculer_total_recettes(recettes_editees)
                total_d = calculer_total_depenses(depenses_editees)
                st.divider()
                m1, m2, m3 = st.columns(3)
                m1.metric("Recettes (mois confirmé)", f"{formater_montant(total_r)} FCFA")
                m2.metric("Dépenses (mois confirmé)", f"{formater_montant(total_d)} FCFA")
                m3.metric("Solde", f"{formater_montant(total_r - total_d)} FCFA")

        # --------------------------------------------------------
        # PAGE FINALE : bilan mensuel consolidé + enregistrement
        # --------------------------------------------------------
        else:
            st.divider()

            recettes_combinees = []
            depenses_combinees = []
            for i in range(total_fichiers):
                if "erreur" in st.session_state.lot_donnees[i]:
                    continue
                recettes_combinees.extend(obtenir_donnees_semaine(i, "recettes"))
                depenses_combinees.extend(obtenir_donnees_semaine(i, "depenses"))

            rapport = calculer_bilan_mensuel_agrege(
                recettes_combinees, depenses_combinees, debut_mois.year, debut_mois.month
            )
            afficher_bilan_mensuel(rapport)

            st.divider()
            if st.session_state.bilan_enregistre:
                st.success("✅ Ce bilan mensuel a déjà été enregistré dans Google Sheets.")
                if st.button("📥 Traiter un nouveau lot d'images (mois suivant)", type="primary"):
                    reinitialiser_lot()
                    st.rerun()
            else:
                if st.button("💾 Enregistrer le bilan mensuel dans Google Sheets", type="primary"):
                    with st.spinner("Écriture dans Google Sheets en cours..."):
                        try:
                            _, feuille_principale = get_clients_config()
                            enregistrer_bilan_mensuel(feuille_principale, rapport)
                            st.session_state.bilan_enregistre = True
                            st.success("✨ Bilan mensuel enregistré avec succès !")
                            st.rerun()
                        except Exception as e:
                            st.error(f"🚨 Une erreur est survenue lors de l'enregistrement : {e}")