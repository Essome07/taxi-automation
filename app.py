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
import html
import io
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

# ============================================================
# HABILLAGE VISUEL
# ------------------------------------------------------------
# Tout ce bloc est PUREMENT DÉCORATIF. Il n'ajoute, ne retire et
# ne modifie aucun comportement : les widgets restent des widgets
# Streamlit standards. En cas de souci d'affichage après une mise
# à jour de Streamlit, on peut neutraliser ce bloc sans casser
# l'application (elle retrouve simplement son apparence par défaut).
# ============================================================
FEUILLE_DE_STYLE = """
<style>
:root {
    --accent: #E8A33D;
    --accent-doux: rgba(232, 163, 61, 0.16);
    /* Couleurs volontairement translucides : elles se posent sur le fond du
       thème actif et restent donc lisibles en mode clair comme en mode sombre,
       sans qu'il faille dupliquer toute la feuille de style. */
    --fond-carte: rgba(128, 128, 128, 0.10);
    --bordure: rgba(128, 128, 128, 0.28);
    --texte-doux: rgba(128, 128, 128, 1);
}

/* Masque les vignettes natives du file_uploader : la revue visuelle
   des photos joue déjà ce rôle, de façon plus soignée. */
[data-testid="stFileUploaderFile"] { display: none; }

/* --- Barre latérale --- */
[data-testid="stSidebar"] {
    border-right: 1px solid var(--bordure);
}
[data-testid="stSidebar"] [data-testid="stRadio"] label {
    padding: 6px 10px;
    border-radius: 8px;
    transition: background 0.15s ease;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
    background: var(--accent-doux);
}

/* --- Boutons --- */
.stButton > button {
    border-radius: 10px;
    font-weight: 600;
    border: 1px solid var(--bordure);
    transition: transform 0.08s ease, box-shadow 0.15s ease;
}
.stButton > button:hover:not(:disabled) {
    transform: translateY(-1px);
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.35);
}
.stButton > button:disabled { opacity: 0.45; }

/* --- Cartes d'information --- */
.carte-entete {
    background: var(--fond-carte);
    border: 1px solid var(--bordure);
    border-left: 3px solid var(--accent);
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 4px;
}
.carte-entete .surtitre {
    font-size: 0.68rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--texte-doux);
}
.carte-entete .titre { font-size: 1.05rem; font-weight: 700; }

/* --- Fil des étapes --- */
.fil-etapes {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin: 4px 0 18px 0;
}
.fil-etapes .etape {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 7px 14px 7px 8px;
    border-radius: 999px;
    background: var(--fond-carte);
    border: 1px solid var(--bordure);
    font-size: 0.86rem;
    color: var(--texte-doux);
}
.fil-etapes .etape .puce {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 22px; height: 22px;
    border-radius: 50%;
    background: rgba(128, 128, 128, 0.18);
    font-size: 0.76rem;
    font-weight: 700;
}
.fil-etapes .etape.active {
    background: var(--accent-doux);
    border-color: var(--accent);
    color: #F4E7D0;
    font-weight: 600;
}
.fil-etapes .etape.active .puce { background: var(--accent); color: #14161A; }
.fil-etapes .etape.active { color: inherit; }
.fil-etapes .etape.faite { color: inherit; opacity: 0.85; }
.fil-etapes .etape.faite .puce {
    background: rgba(232, 163, 61, 0.40);
    color: inherit;
}

/* --- Cartes de semaines (aperçu du lot) --- */
.grille-semaines {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px;
    margin: 6px 0 18px 0;
}
.carte-semaine {
    background: var(--fond-carte);
    border: 1px solid var(--bordure);
    border-radius: 12px;
    overflow: hidden;
    display: flex;
    flex-direction: column;
}
.carte-semaine .vignette {
    height: 108px;
    background: rgba(128, 128, 128, 0.10);
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
}
.carte-semaine .vignette img {
    width: 100%; height: 100%;
    object-fit: cover;
}
.carte-semaine .corps { padding: 10px 12px 12px 12px; }
.carte-semaine .nom { font-weight: 700; font-size: 0.9rem; }
.carte-semaine .periode {
    font-size: 0.76rem;
    color: var(--texte-doux);
    margin: 2px 0 8px 0;
}
.carte-semaine .statut {
    font-size: 0.74rem;
    padding: 3px 8px;
    border-radius: 999px;
    display: inline-block;
    background: rgba(128, 128, 128, 0.16);
    color: var(--texte-doux);
}
.carte-semaine .statut.ok {
    background: var(--accent-doux);
    color: inherit;
    font-weight: 600;
}
.carte-semaine .statut.souci {
    background: rgba(220, 90, 90, 0.20);
    color: #C0392B;
}

/* --- Tableaux et éditeurs --- */
/* Pas de `overflow: hidden` ici : la barre d'outils du tableau (ajout et
   suppression de lignes) flotte au-dessus du coin supérieur droit et serait
   rognée, rendant la suppression native inaccessible. */
[data-testid="stDataFrame"], [data-testid="stTable"] {
    border-radius: 10px;
}
[data-testid="stElementToolbar"] {
    z-index: 10;
}

/* --- Titres --- */
h1, h2, h3 { letter-spacing: -0.01em; }
</style>
"""
st.markdown(FEUILLE_DE_STYLE, unsafe_allow_html=True)


def echapper_html(texte: str) -> str:
    """Neutralise les caractères spéciaux avant insertion dans un bloc HTML
    décoratif (un prénom contenant « & » ou « < » ne doit pas casser la page)."""
    return html.escape(str(texte), quote=True)


def vignette_base64(fichier, largeur: int = 240) -> str:
    """Produit une miniature encodée en base64, insérable directement dans le
    HTML des cartes. On redimensionne avant encodage : intégrer les photos en
    pleine résolution alourdirait la page de plusieurs mégaoctets à chaque
    rechargement. Le résultat est mis en cache par empreinte du fichier."""
    cache = st.session_state.setdefault("cache_vignettes", {})
    fichier.seek(0)
    contenu = fichier.read()
    fichier.seek(0)
    empreinte = hashlib.md5(contenu).hexdigest()
    if empreinte in cache:
        return cache[empreinte]

    try:
        from PIL import Image

        image = Image.open(io.BytesIO(contenu))
        image = image.convert("RGB")
        ratio = largeur / image.width
        image = image.resize((largeur, max(1, int(image.height * ratio))))
        tampon = io.BytesIO()
        image.save(tampon, format="JPEG", quality=70)
        encode = base64.b64encode(tampon.getvalue()).decode("utf-8")
    except Exception:
        # En cas de format inattendu, on retombe sur l'image d'origine :
        # l'aperçu reste affiché plutôt que de faire échouer la page.
        encode = base64.b64encode(contenu).decode("utf-8")

    cache[empreinte] = encode
    return encode


def afficher_cartes_semaines(fichiers: list, lot_donnees) -> None:
    """Aperçu du lot sous forme de cartes : vignette, numéro de semaine,
    période détectée et état d'avancement. Purement informatif — aucune
    interaction, donc aucun risque de casser le parcours."""
    cartes = []
    for indice, fichier in enumerate(fichiers):
        donnees = lot_donnees[indice] if lot_donnees and indice < len(lot_donnees) else None

        if donnees is None:
            periode, statut, classe = "En attente d'analyse", "Photo chargée", ""
        elif "erreur" in donnees:
            periode, statut, classe = "—", "Non lisible", "souci"
        else:
            periode = donnees.get("periode_hebdo") or "Période non détectée"
            total = calculer_total_recettes(donnees.get("recettes_journalieres", []))
            statut, classe = f"Analysée · {formater_montant(total)} FCFA", "ok"

        cartes.append(
            f'<div class="carte-semaine">'
            f'<div class="vignette"><img src="data:image/jpeg;base64,{vignette_base64(fichier)}"></div>'
            f'<div class="corps">'
            f'<div class="nom">Semaine {indice + 1}</div>'
            f'<div class="periode">{echapper_html(periode)}</div>'
            f'<div class="statut {classe}">{echapper_html(statut)}</div>'
            f'</div></div>'
        )

    st.markdown(f'<div class="grille-semaines">{"".join(cartes)}</div>', unsafe_allow_html=True)


def afficher_fil_etapes(etape_active: int) -> None:
    """Affiche le fil des 5 étapes du parcours. Purement indicatif : la
    progression est déduite de l'état réel de l'application, ce fil ne
    pilote rien et ne permet pas de naviguer."""
    etapes = ["Photos", "Analyse", "Confirmation", "Vérification", "Bilan"]
    morceaux = []
    for numero, libelle in enumerate(etapes, start=1):
        if numero < etape_active:
            classe, puce = "etape faite", "✓"
        elif numero == etape_active:
            classe, puce = "etape active", str(numero)
        else:
            classe, puce = "etape", str(numero)
        morceaux.append(
            f'<div class="{classe}"><span class="puce">{puce}</span>{libelle}</div>'
        )
    st.markdown(f'<div class="fil-etapes">{"".join(morceaux)}</div>', unsafe_allow_html=True)


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

VALUE_INPUT_OPTION = "USER_ENTERED"  # évite que Sheets force les nombres en texte (bug de l'apostrophe)


# ============================================================
# CONFIGURATION UTILISATEUR (persistée sur disque)
# ============================================================
def charger_config() -> dict:
    defaut = {
        "sheet_principale_id": SHEET_PRINCIPALE_ID_DEFAUT,
        "nom_onglet_rapport": "",
        "emails_partage_rapport": "",
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


def lister_classeurs_accessibles() -> list:
    """Liste les classeurs Google Sheets auxquels le compte de service a
    réellement accès. C'est le moyen le plus direct de vérifier qu'un partage
    a bien été pris en compte, et de récupérer le bon identifiant : si un
    fichier n'apparaît pas ici, c'est que le partage n'a pas abouti."""
    session = AuthorizedSession(charger_credentials())
    reponse = session.get(
        "https://www.googleapis.com/drive/v3/files",
        params={
            "q": "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
            "fields": "files(id,name,owners/emailAddress)",
            "pageSize": 50,
            "supportsAllDrives": "true",
            "includeItemsFromAllDrives": "true",
        },
        timeout=30,
    )
    if reponse.status_code != 200:
        raise RuntimeError(
            f"Google Drive a répondu avec le code {reponse.status_code}. "
            "Vérifie que l'API Google Drive est bien activée sur le projet du compte de service."
        )
    return reponse.json().get("files", [])


def lister_onglets_classeur(sheet_id: str) -> list:
    """Renvoie les noms des onglets d'un classeur, pour que l'utilisateur
    choisisse sa destination dans une liste plutôt que de la saisir à la main
    (une majuscule ou un accent de travers suffisait à faire échouer l'écriture)."""
    classeur = get_gspread_client().open_by_key(sheet_id)
    return [f.title for f in classeur.worksheets()]


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


def get_gemini_key() -> str:
    """Récupère la clé Gemini seule.

    Volontairement séparé de l'accès à Google Sheets : l'analyse des photos
    n'a besoin que de cette clé. Les coupler ferait échouer l'analyse dès que
    le classeur est inaccessible, alors que les deux services sont
    indépendants."""
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception as exc:
        raise RuntimeError(
            "La clé Gemini est introuvable dans les secrets de l'application. "
            "Vérifie qu'une ligne `GEMINI_API_KEY = \"...\"` existe bien dans les secrets "
            "(⋮ → Settings → Secrets sur Streamlit Cloud), puis redémarre l'application."
        ) from exc


def get_clients(sheet_principale_id: str):
    """Récupère la feuille de calcul à CHAQUE appel (volontairement PAS mise
    en cache) : si on gardait l'objet Worksheet en mémoire, un onglet
    supprimé/recréé manuellement dans Google Sheets casserait l'app avec une
    erreur 'No grid with id ...' jusqu'au redémarrage. Cet appel API est
    rapide et peu coûteux vu la fréquence d'utilisation de l'app."""
    gc = get_gspread_client()
    api_key = get_gemini_key()
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


def decrire_erreur(exc: Exception) -> str:
    """Produit un message toujours exploitable, clé API masquée.

    Certaines exceptions ne portent aucun texte : gspread lève par exemple
    SpreadsheetNotFound sans argument, ce qui affichait « Erreur : » suivi de
    rien du tout. On retombe alors sur le type de l'exception, complété d'une
    explication quand la cause est connue."""
    texte = masquer_cle_api(str(exc)).strip()
    if texte:
        return texte

    nom = type(exc).__name__
    explications = {
        "SpreadsheetNotFound": (
            "classeur Google Sheets introuvable. Vérifie l'identifiant configuré dans "
            "⚙️ Paramètres et que la feuille est bien partagée avec le compte de service."
        ),
        "WorksheetNotFound": (
            "onglet introuvable dans le classeur. Vérifie son nom dans ⚙️ Paramètres."
        ),
        "APIError": "Google a refusé la requête. Vérifie les droits d'accès au classeur.",
    }
    detail = explications.get(nom)
    return f"{nom} — {detail}" if detail else f"{nom} (aucun détail fourni par le service)."


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


class FichierMemorise(io.BytesIO):
    """Copie en mémoire d'une image téléversée.

    Streamlit efface le contenu d'un file_uploader dès que celui-ci n'est plus
    affiché à l'écran : changer d'onglet ferait donc perdre les 5 photos et
    relancerait tout le processus depuis le début. En recopiant les octets ici,
    le lot survit à la navigation. Cette classe expose les mêmes attributs
    qu'un fichier téléversé (name, size, type) pour rester interchangeable."""

    def __init__(self, nom: str, type_mime: str, contenu: bytes):
        super().__init__(contenu)
        self.name = nom
        self.type = type_mime or "image/jpeg"
        self.size = len(contenu)


def memoriser_fichiers(fichiers_uploades: list) -> list:
    return [
        FichierMemorise(f.name, getattr(f, "type", None), f.getvalue())
        for f in fichiers_uploades
    ]


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
    depenses_detaillees: list = []
    for dep in depenses_combinees:
        titre = str(dep.get("titre", "")).strip() or "Dépense"
        try:
            montant = float(dep.get("montant", 0) or 0)
        except (TypeError, ValueError):
            montant = 0.0
        depenses_par_titre[titre] = depenses_par_titre.get(titre, 0.0) + montant
        depenses_detaillees.append({
            "date": parser_date(str(dep.get("date", ""))),
            "titre": titre,
            "montant": montant,
        })

    # Tri chronologique ; les dépenses sans date exploitable sont reléguées en fin
    depenses_detaillees.sort(key=lambda d: (d["date"] is None, d["date"] or datetime.date.max))

    total_depenses = sum(depenses_par_titre.values())
    return {
        "annee": annee,
        "mois": mois,
        "jours_travailles": len(jours_recette),
        "recette_totale": recette_totale,
        "depenses_par_titre": depenses_par_titre,
        "depenses_detaillees": depenses_detaillees,
        "total_depenses": total_depenses,
        "solde_net": recette_totale - total_depenses,
    }


# ============================================================
# RAPPORT MENSUEL AUTONOME (un classeur Google Sheets par mois)
# Reproduit la maquette validée : tableau daté des dépenses,
# puis bloc de synthèse (jours travaillés, recettes, dépenses, solde).
# ============================================================
TITRE_RAPPORT_MENSUEL = "Rapport mensuel ({mois:02d}/{annee})"
EN_TETES_RAPPORT = ["DATE", "CATÉGORIE", "TOTAL (XAF)"]
COULEUR_EN_TETE = {"red": 0.60, "green": 0.11, "blue": 0.16}  # bordeaux de la maquette


def construire_lignes_rapport(rapport: dict) -> tuple[list, int]:
    """Prépare le contenu complet de la feuille : en-têtes, une ligne par
    dépense datée, puis le bloc de synthèse. Renvoie aussi le numéro de la
    première ligne de synthèse (utile pour la mise en forme)."""
    lignes = [EN_TETES_RAPPORT]

    for dep in rapport.get("depenses_detaillees", []):
        date_affichee = dep["date"].strftime("%d/%m/%y") if dep["date"] else ""
        lignes.append([date_affichee, dep["titre"], dep["montant"]])

    premiere_ligne_synthese = len(lignes) + 1
    lignes.extend([
        ["Nombre de jours travaillés", "", rapport["jours_travailles"]],
        ["Recette totale", "", rapport["recette_totale"]],
        ["Dépenses totales", "", rapport["total_depenses"]],
        ["Solde net", "", rapport["solde_net"]],
    ])
    return lignes, premiere_ligne_synthese


def mettre_en_forme_rapport(feuille, nb_lignes_total: int, premiere_ligne_synthese: int) -> None:
    """Applique la mise en forme de la maquette : en-tête bordeaux, montants
    alignés à droite avec séparateur de milliers, synthèse en gras italique."""
    sheet_id = feuille.id
    requetes = [
        # Ligne d'en-tête : fond bordeaux, texte blanc en gras
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                          "startColumnIndex": 0, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": COULEUR_EN_TETE,
                    "textFormat": {"bold": True,
                                   "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                    "verticalAlignment": "MIDDLE",
                }},
                "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)",
            }
        },
        # Ligne d'en-tête figée : les titres restent visibles au défilement
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
        # Montants : séparateur de milliers, sans décimale
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": nb_lignes_total,
                          "startColumnIndex": 2, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0"},
                    "horizontalAlignment": "RIGHT",
                }},
                "fields": "userEnteredFormat(numberFormat,horizontalAlignment)",
            }
        },
        # Bloc de synthèse : gras italique, comme sur la maquette
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id,
                          "startRowIndex": premiere_ligne_synthese - 1,
                          "endRowIndex": nb_lignes_total,
                          "startColumnIndex": 0, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "italic": True}}},
                "fields": "userEnteredFormat.textFormat",
            }
        },
        # Largeurs de colonnes lisibles
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 200},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": 1, "endIndex": 3},
                "properties": {"pixelSize": 150},
                "fields": "pixelSize",
            }
        },
    ]
    feuille.spreadsheet.batch_update({"requests": requetes})


NOM_ONGLET_RAPPORT = "Rapport {mois:02d}-{annee}"


def nom_onglet_rapport_cible(rapport: dict) -> str:
    """Nom de l'onglet où écrire le rapport mensuel : celui choisi dans
    ⚙️ Paramètres s'il est renseigné, sinon un nom automatique par mois."""
    choisi = (st.session_state.config.get("nom_onglet_rapport") or "").strip()
    return choisi or NOM_ONGLET_RAPPORT.format(mois=rapport["mois"], annee=rapport["annee"])


def creer_rapport_mensuel_onglet(rapport: dict, nom_onglet: str = "") -> dict:
    """Écrit le récapitulatif du mois dans un onglet du classeur configuré.

    Si l'onglet existe déjà, son contenu est remplacé sans demander
    confirmation : la destination est choisie juste avant de cliquer, l'écriture
    est donc toujours délibérée. Google Sheets conserve de toute façon
    l'historique des versions du classeur en cas de fausse manœuvre.

    `nom_onglet` permet de désigner explicitement la destination depuis la page
    du bilan ; à défaut, on retombe sur le réglage des paramètres ou sur le nom
    automatique du mois."""
    classeur = get_gspread_client().open_by_key(st.session_state.config["sheet_principale_id"])
    nom_onglet = (nom_onglet or "").strip() or nom_onglet_rapport_cible(rapport)
    lignes, premiere_ligne_synthese = construire_lignes_rapport(rapport)

    existant = next((f for f in classeur.worksheets() if f.title == nom_onglet), None)

    if existant is not None:
        # On vide l'onglet plutôt que de le supprimer : sa position dans le
        # classeur est conservée, ce qui compte quand c'est l'utilisateur qui
        # l'a créé et placé lui-même.
        feuille = existant
        feuille.clear()
        feuille.resize(
            rows=max(len(lignes) + 10, feuille.row_count),
            cols=max(6, feuille.col_count),
        )
    else:
        feuille = classeur.add_worksheet(title=nom_onglet, rows=max(len(lignes) + 10, 50), cols=6)

    feuille.update(f"A1:C{len(lignes)}", lignes, value_input_option=VALUE_INPUT_OPTION)
    mettre_en_forme_rapport(feuille, len(lignes), premiere_ligne_synthese)

    return {
        "titre": nom_onglet,
        "classeur": classeur.title,
        "url": f"{classeur.url}#gid={feuille.id}",
        "remplace": existant is not None,
    }


def creer_rapport_mensuel_sheets(rapport: dict, emails_partage: list) -> dict:
    """Crée un classeur Google Sheets dédié au mois, y écrit le récapitulatif
    mis en forme, puis le partage avec les adresses fournies.

    Le classeur est créé par le compte de service, qui en reste propriétaire :
    le partage est donc INDISPENSABLE pour que quelqu'un d'autre puisse
    l'ouvrir. C'est pour cette raison que la fonction refuse de créer un
    rapport que personne ne pourrait consulter."""
    if not emails_partage:
        raise RuntimeError(
            "Aucune adresse de partage n'est configurée. Le classeur serait créé au nom du "
            "compte de service et resterait invisible pour toi comme pour le client. "
            "Renseigne au moins une adresse Gmail dans ⚙️ Paramètres."
        )

    gc = get_gspread_client()
    titre = TITRE_RAPPORT_MENSUEL.format(mois=rapport["mois"], annee=rapport["annee"])
    classeur = gc.create(titre)
    feuille = classeur.get_worksheet(0)

    lignes, premiere_ligne_synthese = construire_lignes_rapport(rapport)
    feuille.update(f"A1:C{len(lignes)}", lignes, value_input_option=VALUE_INPUT_OPTION)
    mettre_en_forme_rapport(feuille, len(lignes), premiere_ligne_synthese)

    partages_reussis, partages_echoues = [], []
    for email in emails_partage:
        try:
            classeur.share(email, perm_type="user", role="writer", notify=False)
            partages_reussis.append(email)
        except Exception as e:
            partages_echoues.append(f"{email} ({e})")

    return {
        "titre": titre,
        "url": classeur.url,
        "partages_reussis": partages_reussis,
        "partages_echoues": partages_echoues,
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

    if "storageQuotaExceeded" in texte or "storage quota" in texte.lower():
        return (
            "Un compte de service ne possède pas d'espace de stockage Google Drive : il ne peut "
            "donc pas créer de nouveau fichier. Utilise la création d'un onglet dans un classeur "
            "existant, qui appartient à un vrai compte utilisateur."
        )
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
    # Filet de sécurité : certaines exceptions gspread n'ont aucun message.
    return decrire_erreur(exc)


# ============================================================
# ÉTAT DE SESSION
# ============================================================
for cle, defaut in {
    "config": None,
    "cle_uploader": 0,
    "signature_lot": None,
    "fichiers_lot": None,
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
    "semaines_validees": set(),
    "bilan_etabli": False,
    "rapport_fige": None,
    "donnees_editees": {},
    "rapport_mensuel_cree": None,
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
    st.session_state.fichiers_lot = None
    st.session_state.lot_donnees = None
    st.session_state.donnees_filtrees = None
    st.session_state.mois_confirme = False
    st.session_state.sous_page = 0
    st.session_state.bilan_enregistre = False
    st.session_state.trous_ignores = set()
    st.session_state.fichiers_combles = {}
    st.session_state.ordre_chronologique = None
    st.session_state.ordre_upload_estime = None
    st.session_state.semaines_validees = set()
    st.session_state.bilan_etabli = False
    st.session_state.rapport_fige = None
    st.session_state.donnees_editees = {}
    st.session_state.rapport_mensuel_cree = None
    for cle in list(st.session_state.keys()):
        if cle.startswith("dest_depense_"):
            del st.session_state[cle]


CHAMPS_RECETTES = ["date", "montant"]
CHAMPS_DEPENSES = ["titre", "montant", "date"]


def valeur_vide(valeur) -> bool:
    """Une cellule est considérée vide si elle n'a jamais été renseignée.
    Streamlit convertit les tableaux en DataFrame : une case laissée blanche
    peut arriver sous forme de None, de NaN ou de la chaîne « None »."""
    if valeur is None:
        return True
    if isinstance(valeur, float) and valeur != valeur:  # NaN
        return True
    return str(valeur).strip() in ("", "None", "nan", "NaT")


def ligne_vide(ligne: dict, champs: list) -> bool:
    return all(valeur_vide(ligne.get(champ)) for champ in champs)


def nettoyer_lignes_editeur(lignes: list, champs: list) -> list:
    """Ne conserve que les champs attendus et écarte les lignes entièrement
    vides — notamment la ligne blanche que Streamlit ajoute en bas du tableau
    pour permettre les ajouts, et qui ne doit jamais compter comme une donnée."""
    resultat = []
    for ligne in lignes:
        propre = {champ: ligne.get(champ) for champ in champs}
        if ligne_vide(propre, champs):
            continue
        resultat.append(propre)
    return resultat


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
    st.markdown(
        '<div style="font-size:0.68rem;letter-spacing:0.16em;text-transform:uppercase;'
        'color:rgba(232,230,227,0.55);margin:-8px 0 14px 2px;">Rapports mensuels</div>',
        unsafe_allow_html=True,
    )

    heure = datetime.datetime.now().hour
    if heure < 5:
        salutation = "Bonne nuit"
    elif heure < 12:
        salutation = "Bonjour"
    elif heure < 18:
        salutation = "Bon après-midi"
    else:
        salutation = "Bonsoir"
    st.markdown(
        f'<div class="carte-entete">'
        f'<div class="surtitre">{salutation}</div>'
        f'<div class="titre">{echapper_html(st.session_state.config["nom_utilisateur"])}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

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

    st.caption(
        "Le classeur de destination se choisit désormais directement sur la page "
        "📊 Bilan mensuel, au moment de l'écriture."
    )

    st.subheader("Apparence")
    st.caption(
        "L'application suit le thème choisi dans Streamlit. Pour passer du mode sombre au mode "
        "clair : menu **⋮** en haut à droite → **Settings** → **Theme** → *Light* ou *Dark*. "
        "Ce réglage agit sur toute l'interface, y compris les tableaux."
    )

    if st.button("💾 Enregistrer les paramètres", type="primary"):
        st.session_state.config = {
            **st.session_state.config,
            "nom_utilisateur": nouveau_nom.strip() or "Pascal",
        }
        sauvegarder_config(st.session_state.config)
        st.success("✅ Paramètres enregistrés.")

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

                st.markdown("**Classeurs auxquels le compte de service a réellement accès :**")
                try:
                    accessibles = lister_classeurs_accessibles()
                    if not accessibles:
                        st.warning(
                            "⚠️ Le compte de service n'a accès à **aucun** classeur. Le partage n'a donc "
                            "pas abouti. Rouvre le fichier dans Google Sheets, clique sur « Partager », et "
                            "vérifie que l'adresse du compte de service apparaît bien dans la liste des "
                            "personnes ayant accès (il ne suffit pas de la saisir : il faut valider l'envoi)."
                        )
                    else:
                        st.caption(
                            "Copie l'identifiant du bon classeur ci-dessous dans le champ « Classeur du "
                            "client », puis enregistre les paramètres."
                        )
                        st.table([
                            {"Nom du classeur": f.get("name", ""), "Identifiant à copier": f.get("id", "")}
                            for f in accessibles
                        ])
                except Exception as e:
                    st.error(f"🚨 Impossible de lister les classeurs : {message_erreur_sheets(e)}")

            if diagnostic["statut"] in ("editeur", "lecture_seule"):
                try:
                    classeur = get_gspread_client().open_by_key(sheet_id)
                    titres = [f.title for f in classeur.worksheets()]
                    st.write("**Onglets trouvés :** " + ", ".join(f"`{t}`" for t in titres))

                except Exception as e:
                    st.error(f"🚨 Lecture des onglets impossible : {message_erreur_sheets(e)}")

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

    # Progression déduite de l'état réel : ce fil ne pilote rien, il informe.
    if not st.session_state.fichiers_lot:
        _etape_courante = 1
    elif st.session_state.lot_donnees is None:
        _etape_courante = 2
    elif not st.session_state.mois_confirme:
        _etape_courante = 3
    elif st.session_state.sous_page >= len(st.session_state.lot_donnees):
        _etape_courante = 5
    else:
        _etape_courante = 4
    afficher_fil_etapes(_etape_courante)

    st.write(f"Uploade exactement {MAX_IMAGES_PAR_LOT} rapports hebdomadaires (un mois complet, glisser-déposer possible).")

    fichiers_uploades = st.file_uploader(
        f"Sélectionner les {MAX_IMAGES_PAR_LOT} images des rapports",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.cle_uploader}",
    )

    if fichiers_uploades:
        if len(fichiers_uploades) > MAX_IMAGES_PAR_LOT:
            st.warning(f"⚠️ Maximum {MAX_IMAGES_PAR_LOT} images à la fois. Seules les {MAX_IMAGES_PAR_LOT} premières seront prises en compte.")
            fichiers_uploades = fichiers_uploades[:MAX_IMAGES_PAR_LOT]

        if len(fichiers_uploades) < MAX_IMAGES_PAR_LOT:
            st.info(
                f"📸 L'analyse ne se lance qu'à partir d'un lot complet de {MAX_IMAGES_PAR_LOT} rapports "
                f"hebdomadaires (un mois entier). Il t'en manque {MAX_IMAGES_PAR_LOT - len(fichiers_uploades)}."
            )
            st.stop()

        signature = tuple((f.name, f.size) for f in fichiers_uploades)
        if st.session_state.signature_lot != signature:
            for cle in list(st.session_state.keys()):
                if cle.startswith("editeur_recettes_") or cle.startswith("editeur_depenses_"):
                    del st.session_state[cle]
            # Copie en mémoire : le lot doit survivre à un passage par
            # Maintenance ou Paramètres, qui vide le file_uploader.
            st.session_state.fichiers_lot = memoriser_fichiers(fichiers_uploades)
            st.session_state.signature_lot = signature
            st.session_state.lot_donnees = None
            st.session_state.donnees_filtrees = None
            st.session_state.mois_confirme = False
            st.session_state.sous_page = 0
            st.session_state.bilan_enregistre = False
            st.session_state.trous_ignores = set()
            st.session_state.fichiers_combles = {}
            st.session_state.ordre_chronologique = None
            st.session_state.ordre_upload_estime = estimer_ordre_upload(st.session_state.fichiers_lot)
            st.session_state.semaines_validees = set()
            st.session_state.bilan_etabli = False
            st.session_state.rapport_fige = None
            st.session_state.donnees_editees = {}
            st.session_state.rapport_mensuel_cree = None

    fichiers = st.session_state.fichiers_lot

    if fichiers:
        if not fichiers_uploades:
            # Retour depuis un autre onglet : le widget est vide, mais le lot
            # est conservé. On le rappelle pour éviter toute confusion.
            col_info, col_reset = st.columns([3, 1])
            with col_info:
                st.info(f"📂 Lot en cours : {len(fichiers)} photos déjà chargées.")
            with col_reset:
                if st.button("🗑️ Changer de lot", use_container_width=True):
                    reinitialiser_lot()
                    st.rerun()

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

        # Aperçu visuel du lot, tant que la revue détaillée n'a pas commencé
        # (au-delà, chaque semaine a déjà sa propre page avec sa photo).
        if not st.session_state.mois_confirme:
            afficher_cartes_semaines(fichiers, st.session_state.lot_donnees)

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
                        api_key = get_gemini_key()
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
                                resultats_lot.append({"erreur": decrire_erreur(e)})

                        ordre = calculer_ordre_chronologique(resultats_lot)
                        st.session_state.ordre_chronologique = ordre
                        st.session_state.lot_donnees = [resultats_lot[j] for j in ordre]
                        st.rerun()
                    except Exception as e:
                        st.error(f"🚨 Erreur lors de la connexion : {decrire_erreur(e)}")
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
                                        api_key = get_gemini_key()
                                        image_b64, mime_type = encoder_image(fichier_combler)
                                        res_json = appeler_gemini(api_key, image_b64, mime_type)
                                        donnees_comblees = extraire_donnees(res_json)
                                        p = parser_periode(donnees_comblees.get("periode_hebdo", ""))
                                        chevauche = p and not (p[1] < debut_trou or p[0] > fin_trou)
                                        if chevauche:
                                            st.session_state.lot_donnees.append(donnees_comblees)
                                            st.session_state.fichiers_combles[cle_trou] = FichierMemorise(
                                                fichier_combler.name,
                                                getattr(fichier_combler, "type", None),
                                                fichier_combler.getvalue(),
                                            )
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
                                        st.error(f"🚨 Erreur lors de l'analyse : {decrire_erreur(e)}")
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
                        "🖊️ Corrige librement ce que l'IA a mal lu : modifie une date ou un montant, "
                        "ou ajoute une ligne oubliée via la dernière ligne du tableau. "
                        "Pour supprimer une ligne, survole-la et coche la case qui apparaît tout à "
                        "gauche, puis clique sur l'icône 🗑️ en haut à droite du tableau (ou appuie "
                        "sur la touche Suppr)."
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
                    recettes_editees = nettoyer_lignes_editeur(
                        normaliser_lignes_editeur(recettes_editees), CHAMPS_RECETTES
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
                    depenses_editees = nettoyer_lignes_editeur(
                        normaliser_lignes_editeur(depenses_editees), CHAMPS_DEPENSES
                    )

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
                    st.session_state.bilan_enregistre = False
                    st.session_state.rapport_mensuel_cree = None
                    st.session_state.sous_page = 0
                    st.rerun()

            # ====================================================
            # RAPPORT MENSUEL RÉCAPITULATIF (onglet dédié)
            # ====================================================
            st.divider()
            st.markdown("### 📗 Créer le rapport mensuel récapitulatif")

            nom_onglet_prevu = nom_onglet_rapport_cible(rapport)
            lignes_apercu, _ = construire_lignes_rapport(rapport)

            if st.session_state.rapport_mensuel_cree:
                infos = st.session_state.rapport_mensuel_cree
                if infos.get("classeur"):
                    st.success(
                        f"✅ Onglet « {infos['titre']} » créé dans le classeur « {infos['classeur']} »."
                    )
                else:
                    st.success(f"✅ Classeur « {infos['titre']} » créé.")
                    if infos.get("partages_reussis"):
                        st.caption("Partagé avec : " + ", ".join(infos["partages_reussis"]))
                    if infos.get("partages_echoues"):
                        st.warning("⚠️ Partage impossible pour : " + " ; ".join(infos["partages_echoues"]))
                st.markdown(f"🔗 [Ouvrir le rapport dans Google Sheets]({infos['url']})")
                if st.button("🔄 Regénérer le rapport de ce mois"):
                    st.session_state.rapport_mensuel_cree = None
                    st.rerun()
            else:
                st.caption(
                    "Le rapport contient le détail daté des dépenses puis la synthèse du mois "
                    "(jours travaillés, recette totale, dépenses totales, solde net)."
                )

                # Classeur ET onglet se choisissent ICI, à l'endroit exact où
                # l'on déclenche l'écriture : plus rien à régler ailleurs.
                st.markdown("**Classeur de destination**")
                col_cls, col_btn = st.columns([3, 1])
                with col_btn:
                    if st.button("📂 Lister les classeurs", use_container_width=True, key="lister_classeurs_bilan"):
                        with st.spinner("Interrogation de Google Drive..."):
                            try:
                                st.session_state.classeurs_accessibles = lister_classeurs_accessibles()
                            except Exception as e:
                                st.session_state.classeurs_accessibles = []
                                st.error(f"🚨 {message_erreur_sheets(e)}")

                classeurs_dispo = st.session_state.get("classeurs_accessibles")
                with col_cls:
                    if classeurs_dispo:
                        noms_classeurs = {c["id"]: c.get("name", "(sans nom)") for c in classeurs_dispo}
                        ids_classeurs = list(noms_classeurs.keys())
                        id_courant = st.session_state.config["sheet_principale_id"]
                        index_cls = ids_classeurs.index(id_courant) if id_courant in ids_classeurs else 0
                        classeur_choisi = st.selectbox(
                            "Classeur de destination",
                            options=ids_classeurs,
                            format_func=lambda i: noms_classeurs[i],
                            index=index_cls,
                            label_visibility="collapsed",
                            key="classeur_cible_rapport",
                        )
                    else:
                        classeur_choisi = st.text_input(
                            "Classeur de destination",
                            value=st.session_state.config["sheet_principale_id"],
                            label_visibility="collapsed",
                            placeholder="Identifiant ou URL du classeur Google Sheets",
                            key="classeur_cible_rapport",
                        )

                classeur_choisi = extraire_id_depuis_url(classeur_choisi)
                if classeur_choisi and classeur_choisi != st.session_state.config["sheet_principale_id"]:
                    # Mémorisé pour les prochains lots, sans passer par les Paramètres.
                    st.session_state.config["sheet_principale_id"] = classeur_choisi
                    sauvegarder_config(st.session_state.config)
                    st.session_state.onglets_classeur = None
                    st.rerun()

                st.caption(
                    "Clique sur « Lister les classeurs » pour choisir parmi ceux partagés avec le "
                    "compte de service, ou colle l'identifiant / l'URL du classeur."
                )

                # Le choix de l'onglet se fait ICI, à l'endroit exact où l'on
                # déclenche l'écriture : plus besoin d'aller le régler ailleurs.
                st.markdown("**Onglet de destination**")
                col_choix, col_liste = st.columns([3, 1])
                with col_liste:
                    if st.button("🔄 Lister les onglets", use_container_width=True, key="lister_onglets_bilan"):
                        with st.spinner("Lecture des onglets..."):
                            try:
                                st.session_state.onglets_classeur = lister_onglets_classeur(
                                    st.session_state.config["sheet_principale_id"]
                                )
                            except Exception as e:
                                st.session_state.onglets_classeur = []
                                st.error(f"🚨 {message_erreur_sheets(e)}")

                onglets_connus = st.session_state.get("onglets_classeur") or []
                options_onglets = [nom_onglet_prevu] + [t for t in onglets_connus if t != nom_onglet_prevu]

                with col_choix:
                    if len(options_onglets) > 1:
                        onglet_choisi = st.selectbox(
                            "Onglet de destination",
                            options=options_onglets,
                            index=0,
                            label_visibility="collapsed",
                            key="onglet_cible_rapport",
                        )
                    else:
                        onglet_choisi = st.text_input(
                            "Onglet de destination",
                            value=nom_onglet_prevu,
                            label_visibility="collapsed",
                            placeholder="Nom exact de l'onglet",
                            key="onglet_cible_rapport",
                        )

                st.caption(
                    f"Onglet ciblé : **{onglet_choisi or nom_onglet_prevu}**. "
                    "Clique sur « Lister les onglets » pour choisir parmi ceux déjà présents dans "
                    "le classeur, ou saisis un nouveau nom pour créer un onglet. "
                    "Si l'onglet existe déjà, son contenu sera remplacé."
                )

                with st.expander(f"👁️ Aperçu du contenu ({len(lignes_apercu)} lignes)"):
                    st.table([
                        {"DATE": l[0], "CATÉGORIE": l[1], "TOTAL (XAF)": l[2]}
                        for l in lignes_apercu[1:]
                    ])

                if st.button(
                    "📗 Écrire le rapport dans cet onglet",
                    type="primary",
                    use_container_width=True,
                ):
                    with st.spinner("Écriture et mise en forme en cours..."):
                        try:
                            st.session_state.rapport_mensuel_cree = creer_rapport_mensuel_onglet(
                                rapport, nom_onglet=onglet_choisi
                            )
                            st.session_state.onglets_classeur = None
                            st.rerun()
                        except Exception as e:
                            st.error(f"🚨 Création impossible : {message_erreur_sheets(e)}")

                with st.expander("⚙️ Créer plutôt un classeur séparé (déconseillé)"):
                    st.caption(
                        "Un classeur créé par l'app appartiendrait au compte de service, qui ne dispose "
                        "d'aucun espace de stockage Google Drive : cette création échoue donc dans la "
                        "plupart des cas. À n'utiliser que si le compte de service est rattaché à un "
                        "Google Workspace avec Drive partagé."
                    )
                    emails_bruts = st.session_state.config.get("emails_partage_rapport", "")
                    emails_partage = [e.strip() for e in re.split(r"[,;\s]+", emails_bruts) if e.strip()]
                    if emails_partage:
                        st.caption("Serait partagé avec : " + ", ".join(emails_partage))
                    else:
                        st.caption("Aucune adresse de partage configurée (⚙️ Paramètres).")

                    if st.button("Tenter la création d'un classeur séparé", disabled=not emails_partage):
                        with st.spinner("Tentative de création..."):
                            try:
                                st.session_state.rapport_mensuel_cree = creer_rapport_mensuel_sheets(
                                    rapport, emails_partage
                                )
                                st.rerun()
                            except Exception as e:
                                st.error(f"🚨 Création impossible : {message_erreur_sheets(e)}")

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