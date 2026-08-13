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
.carte-semaine .vignette.sans-photo {
    font-size: 1.8rem;
    color: var(--texte-doux);
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

/* --- Panneau de l'agent de suivi --- */
.panneau-agent {
    background: var(--fond-carte);
    border: 1px solid var(--bordure);
    border-left: 3px solid var(--accent);
    border-radius: 10px;
    padding: 12px 16px;
    margin: 6px 0 16px 0;
}
.panneau-agent .titre-agent {
    font-size: 0.68rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--texte-doux);
    margin-bottom: 8px;
}
.panneau-agent .ligne-agent {
    font-size: 0.92rem;
    line-height: 1.5;
    padding: 3px 0;
}
.panneau-agent .ligne-agent.alerte { color: #D98324; }
.panneau-agent .ligne-agent.succes { color: #4C9A5A; }

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


def normaliser_texte(texte: str) -> str:
    """Minuscules, sans accents : permet de reconnaître « Février 2026 » aussi
    bien que « fevrier 2026 » quand on relit les noms d'onglets d'un classeur."""
    texte = unicodedata.normalize("NFD", str(texte or ""))
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texte.lower()).strip()


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
            if donnees.get("saisie_manuelle"):
                statut, classe = f"Saisie manuelle · {formater_montant(total)} FCFA", "ok"
            else:
                statut, classe = f"Analysée · {formater_montant(total)} FCFA", "ok"

        # Une semaine saisie à la main n'a pas de photo : on affiche un
        # symbole de substitution plutôt que de tenter une vignette.
        if fichier is None:
            vignette = '<div class="vignette sans-photo">✍️</div>'
        else:
            vignette = (
                f'<div class="vignette">'
                f'<img src="data:image/jpeg;base64,{vignette_base64(fichier)}"></div>'
            )

        cartes.append(
            f'<div class="carte-semaine">'
            f'{vignette}'
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

# Recette d'une journée pleine. Le nombre de jours travaillés en découle :
# recette totale ÷ tarif. Une journée à moitié travaillée (7 500 F) compte
# ainsi pour 0,5 jour, ce qu'un simple comptage de dates ne permettait pas.
TARIF_JOURNALIER_DEFAUT = 15000



VALUE_INPUT_OPTION = "USER_ENTERED"  # évite que Sheets force les nombres en texte (bug de l'apostrophe)


# ============================================================
# CONFIGURATION UTILISATEUR (persistée sur disque)
# ============================================================
def charger_config() -> dict:
    defaut = {
        "sheet_principale_id": SHEET_PRINCIPALE_ID_DEFAUT,
        "nom_onglet_rapport": "",
        "nom_utilisateur": "Pascal",
        "tarif_journalier": TARIF_JOURNALIER_DEFAUT,
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                defaut.update(json.load(f))
        except Exception:
            pass
    return defaut


def sauvegarder_config(config: dict) -> bool:
    """Mémorise la configuration sur disque. Renvoie False si l'écriture est
    impossible (hébergement en lecture seule) : les réglages restent alors
    valables pour la session, ce qui ne doit pas faire échouer l'application."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


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


def valeur_vide(valeur) -> bool:
    """Une cellule est considérée vide si elle n'a jamais été renseignée.
    Streamlit convertit les tableaux en DataFrame : une case laissée blanche
    peut arriver sous forme de None, de NaN ou de la chaîne « None »."""
    if valeur is None:
        return True
    if isinstance(valeur, float) and valeur != valeur:  # NaN
        return True
    return str(valeur).strip() in ("", "None", "nan", "NaT")


def parser_montant(valeur):
    """Convertit en nombre un montant qui peut arriver sous des formes variées.

    L'IA lit une écriture manuscrite : elle peut renvoyer 15000, mais aussi
    « 15000F », « 15 000 » ou « 15.000 ». Un simple float() échouait sur ces
    formes et le montant était alors compté comme zéro, sans la moindre
    alerte — le bilan devenait faux de façon invisible.

    Renvoie None si la valeur est vraiment illisible, pour que l'appelant
    puisse le signaler au lieu de l'ignorer."""
    if valeur is None:
        return None
    if isinstance(valeur, bool):
        return None
    if isinstance(valeur, (int, float)):
        if isinstance(valeur, float) and valeur != valeur:  # NaN
            return None
        return float(valeur)

    texte = re.sub(r"[^\d,.\-]", "", str(valeur).strip())
    if not texte or texte.strip("-.,") == "":
        return None

    negatif = texte.startswith("-")
    texte = texte.lstrip("-")

    # Distinguer séparateur de milliers et séparateur décimal : en FCFA les
    # centimes n'existent pas, un groupe final de 3 chiffres est donc un
    # séparateur de milliers (« 15.000 » = quinze mille, pas quinze).
    if "," in texte and "." in texte:
        if texte.rfind(",") > texte.rfind("."):
            texte = texte.replace(".", "").replace(",", ".")
        else:
            texte = texte.replace(",", "")
    elif "," in texte:
        morceaux = texte.split(",")
        texte = texte.replace(",", "" if len(morceaux[-1]) == 3 else ".")
    elif "." in texte:
        morceaux = texte.split(".")
        if len(morceaux[-1]) == 3:
            texte = texte.replace(".", "")

    try:
        montant = float(texte)
    except ValueError:
        return None
    return -montant if negatif else montant


def extraire_donnees(res_json: dict) -> dict:
    if "error" in res_json:
        raise RuntimeError(res_json["error"].get("message", "Erreur inconnue de l'API Gemini."))

    if not res_json.get("candidates"):
        raise RuntimeError("Gemini n'a renvoyé aucun résultat. L'image est peut-être illisible ou trop floue.")

    candidat = res_json["candidates"][0]
    if candidat.get("finishReason") == "SAFETY":
        raise RuntimeError("La demande a été bloquée par les filtres de sécurité de Gemini.")

    try:
        texte = candidat["content"]["parts"][0]["text"]
    except (KeyError, IndexError, TypeError):
        raison = candidat.get("finishReason", "inconnue")
        raise RuntimeError(
            f"Gemini a renvoyé une réponse vide (motif : {raison}). "
            "L'image est peut-être illisible, ou la réponse a été tronquée."
        )
    texte = re.sub(r"^```(json)?|```$", "", texte.strip(), flags=re.MULTILINE).strip()

    try:
        donnees, _ = json.JSONDecoder().raw_decode(texte)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Réponse Gemini invalide (JSON incorrect) : {exc}") from exc

    if not isinstance(donnees, dict):
        raise RuntimeError("Réponse Gemini inattendue : un objet était attendu.")

    return normaliser_donnees_extraites(donnees)


def normaliser_donnees_extraites(donnees: dict) -> dict:
    """Met les données lues par l'IA dans une forme sûre et homogène.

    Sans cette étape, un montant renvoyé sous forme de texte (« 15000F ») ou
    une liste mal typée se propagerait dans tous les calculs. On convertit ici
    une bonne fois pour toutes, et l'on garde trace des valeurs illisibles."""
    resultat = {
        "periode_hebdo": str(donnees.get("periode_hebdo") or "").strip(),
        "recettes_journalieres": [],
        "depenses": [],
    }
    non_lus = 0

    brut_recettes = donnees.get("recettes_journalieres")
    if isinstance(brut_recettes, list):
        for ligne in brut_recettes:
            if not isinstance(ligne, dict):
                continue
            montant = parser_montant(ligne.get("montant"))
            if montant is None and not valeur_vide(ligne.get("montant")):
                non_lus += 1
            resultat["recettes_journalieres"].append({
                "date": str(ligne.get("date") or "").strip(),
                "montant": montant if montant is not None else 0,
            })

    brut_depenses = donnees.get("depenses")
    if isinstance(brut_depenses, list):
        for ligne in brut_depenses:
            if not isinstance(ligne, dict):
                continue
            montant = parser_montant(ligne.get("montant"))
            if montant is None and not valeur_vide(ligne.get("montant")):
                non_lus += 1
            resultat["depenses"].append({
                "titre": str(ligne.get("titre") or "Dépense").strip() or "Dépense",
                "montant": montant if montant is not None else 0,
                "date": str(ligne.get("date") or "").strip(),
            })

    if non_lus:
        resultat["montants_non_lus"] = non_lus
    return resultat


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


def jours_couverts_par_lot(lot_donnees: list[dict]) -> set:
    """Ensemble des jours effectivement couverts par les rapports du lot.

    On se fie d'abord à la période hebdomadaire lue sur l'image ; si elle est
    illisible, on retombe sur les dates des recettes journalières."""
    jours = set()
    for donnees in lot_donnees:
        if not isinstance(donnees, dict) or "erreur" in donnees:
            continue
        periode = parser_periode(donnees.get("periode_hebdo", ""))
        if periode:
            jour, fin = periode
            while jour <= fin:
                jours.add(jour)
                jour += datetime.timedelta(days=1)
        else:
            for recette in donnees.get("recettes_journalieres", []):
                date_jour = parser_date(str(recette.get("date", "")))
                if date_jour:
                    jours.add(date_jour)
    return jours


def detecter_periodes_manquantes(lot_donnees: list[dict], debut_mois, fin_mois) -> list:
    """Périodes du mois qui ne sont couvertes par aucun rapport fourni.

    Contrairement à une simple détection de trous entre deux semaines, on
    compare ici à l'ensemble des jours du mois : un début ou une fin de mois
    absents sont donc signalés, tout comme un lot incomplet (2 semaines sur 5).
    Les jours manquants consécutifs sont regroupés en périodes."""
    if debut_mois is None or fin_mois is None:
        return []

    couverts = jours_couverts_par_lot(lot_donnees)

    manquants = []
    jour = debut_mois
    while jour <= fin_mois:
        if jour not in couverts:
            manquants.append(jour)
        jour += datetime.timedelta(days=1)

    periodes = []
    for jour in manquants:
        if periodes and jour == periodes[-1][1] + datetime.timedelta(days=1):
            periodes[-1] = (periodes[-1][0], jour)
        else:
            periodes.append((jour, jour))
    return periodes


def jours_de_periode(debut, fin) -> list:
    """Liste des jours d'une période, bornes incluses."""
    jours, jour = [], debut
    while jour <= fin:
        jours.append(jour)
        jour += datetime.timedelta(days=1)
    return jours


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
        if f is None:  # semaine saisie manuellement : aucune image à comparer
            continue
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
        montant = parser_montant(r.get("montant"))
        if montant is not None:
            total += montant
    return total


def calculer_total_depenses(depenses: list[dict]) -> float:
    total = 0.0
    for d in depenses:
        montant = parser_montant(d.get("montant"))
        if montant is not None:
            total += montant
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


def tarif_journalier() -> float:
    """Recette d'une journée pleine, servant de base au calcul des jours
    travaillés. Se protège d'une valeur nulle ou absurde, qui ferait échouer
    la division."""
    try:
        valeur = float(st.session_state.config.get("tarif_journalier") or TARIF_JOURNALIER_DEFAUT)
    except (TypeError, ValueError):
        valeur = TARIF_JOURNALIER_DEFAUT
    return valeur if valeur > 0 else TARIF_JOURNALIER_DEFAUT


def formater_jours(valeur: float) -> str:
    """Affiche un nombre de jours sans décimale inutile : 26 plutôt que 26,0,
    mais 26,5 quand la demi-journée compte."""
    valeur = round(float(valeur), 2)
    if valeur.is_integer():
        return str(int(valeur))
    return f"{valeur:.2f}".rstrip("0").rstrip(".").replace(".", ",")


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
        # Nombre de jours au sens économique : recette ÷ tarif journalier.
        # Une journée partiellement travaillée compte donc pour une fraction.
        "jours_travailles": round(recette_totale / tarif_journalier(), 2),
        # Nombre de dates distinctes, conservé pour information.
        "jours_distincts": len(jours_recette),
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
EN_TETES_RAPPORT = ["DATE", "CATÉGORIE", "TOTAL (XAF)"]
# Une couleur d'en-tête par mois : chaque feuille du classeur est ainsi
# identifiable d'un coup d'œil. L'index 0 est inutilisé (les mois vont de 1 à 12).
# La teinte suit les saisons : tons froids en hiver, chauds en été.
COULEURS_EN_TETE_PAR_MOIS = [
    None,
    {"red": 0.16, "green": 0.32, "blue": 0.56},  # janvier   — bleu nuit
    {"red": 0.31, "green": 0.29, "blue": 0.60},  # février   — indigo
    {"red": 0.20, "green": 0.45, "blue": 0.42},  # mars      — vert d'eau
    {"red": 0.22, "green": 0.52, "blue": 0.29},  # avril     — vert feuille
    {"red": 0.45, "green": 0.55, "blue": 0.20},  # mai       — vert olive
    {"red": 0.72, "green": 0.55, "blue": 0.13},  # juin      — ocre
    {"red": 0.80, "green": 0.44, "blue": 0.12},  # juillet   — orange
    {"red": 0.75, "green": 0.28, "blue": 0.15},  # août      — terracotta
    {"red": 0.60, "green": 0.11, "blue": 0.16},  # septembre — bordeaux (maquette)
    {"red": 0.52, "green": 0.20, "blue": 0.35},  # octobre   — prune
    {"red": 0.38, "green": 0.24, "blue": 0.47},  # novembre  — violet
    {"red": 0.25, "green": 0.35, "blue": 0.50},  # décembre  — bleu acier
]


def couleur_en_tete_mois(mois: int) -> dict:
    """Couleur d'en-tête associée à un mois (repli sur le bordeaux si le
    numéro de mois est inattendu)."""
    if 1 <= mois <= 12:
        return COULEURS_EN_TETE_PAR_MOIS[mois]
    return {"red": 0.60, "green": 0.11, "blue": 0.16}


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
        # La colonne du milieu rappelle le tarif : sans lui, un lecteur du
        # classeur ne comprendrait pas d'où sort un nombre de jours décimal.
        ["Nombre de jours travaillés", f"à {formater_montant(tarif_journalier())} F/jour",
         rapport["jours_travailles"]],
        ["Recette totale", "", rapport["recette_totale"]],
        ["Dépenses totales", "", rapport["total_depenses"]],
        ["Solde net", "", rapport["solde_net"]],
    ])
    return lignes, premiere_ligne_synthese


def mettre_en_forme_rapport(feuille, nb_lignes_total: int, premiere_ligne_synthese: int,
                            mois: int = 0) -> None:
    """Applique la mise en forme de la maquette : en-tête coloré selon le mois,
    montants alignés à droite avec séparateur de milliers, synthèse en gras
    italique. La couleur du mois est également appliquée à l'onglet lui-même,
    pour repérer visuellement chaque feuille du classeur."""
    sheet_id = feuille.id
    couleur = couleur_en_tete_mois(mois)
    requetes = [
        # Ligne d'en-tête : fond coloré selon le mois, texte blanc en gras
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                          "startColumnIndex": 0, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": couleur,
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
        # Jours travaillés : format décimal, sinon 26,5 s'afficherait arrondi à 27
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id,
                          "startRowIndex": premiere_ligne_synthese - 1,
                          "endRowIndex": premiere_ligne_synthese,
                          "startColumnIndex": 2, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0.##"},
                }},
                "fields": "userEnteredFormat.numberFormat",
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
        # Couleur de l'onglet, assortie à celle de l'en-tête
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "tabColor": couleur},
                "fields": "tabColor",
            }
        },
    ]
    feuille.spreadsheet.batch_update({"requests": requetes})


NOM_ONGLET_RAPPORT = "{nom_mois} {annee}"


def periode_du_titre_onglet(titre: str):
    """(année, mois) si le titre d'onglet correspond à un rapport mensuel
    (ex. « Avril 2026 »), sinon None. Sert à la fois au rangement des feuilles
    et à l'agent de suivi, qui s'en sert pour savoir quels mois sont traités."""
    reperes = {normaliser_texte(nom): numero for numero, nom in enumerate(NOMS_MOIS) if nom}
    morceaux = normaliser_texte(titre).split()
    if len(morceaux) != 2 or morceaux[0] not in reperes:
        return None
    try:
        return (int(morceaux[1]), reperes[morceaux[0]])
    except ValueError:
        return None


def positionner_onglet_chronologiquement(classeur, feuille, annee: int, mois: int) -> None:
    """Range la feuille parmi les autres rapports mensuels, du plus ancien au
    plus récent. Sans cela, chaque nouvel onglet se placerait en fin de
    classeur : en traitant les mois dans le désordre, l'ordre des feuilles
    deviendrait incohérent."""
    rapports = []
    for onglet in classeur.worksheets():
        periode = periode_du_titre_onglet(onglet.title)
        if periode:
            rapports.append((periode, onglet))

    if len(rapports) < 2:
        return  # une seule feuille de rapport : rien à ordonner

    rapports.sort(key=lambda couple: couple[0])
    position_voulue = next(
        (rang for rang, (periode, _) in enumerate(rapports) if periode == (annee, mois)),
        None,
    )
    if position_voulue is None:
        return

    # Les feuilles hors rapports (données du client, etc.) restent en tête.
    decalage = len(classeur.worksheets()) - len(rapports)
    index_cible = decalage + position_voulue
    if feuille.index != index_cible:
        classeur.reorder_worksheets(
            [f for f in classeur.worksheets() if f.id != feuille.id][:index_cible]
            + [feuille]
            + [f for f in classeur.worksheets() if f.id != feuille.id][index_cible:]
        )


# ============================================================
# AGENT DE SUIVI
# ------------------------------------------------------------
# Petit assistant qui consulte le classeur du client pour savoir
# quels mois sont déjà traités, en déduire celui qu'il attend, et
# signaler les anomalies au fil du parcours (mois déjà présent,
# mois sautés, semaines manquantes).
# ============================================================
def mois_suivant(periode: tuple) -> tuple:
    annee, mois = periode
    return (annee + 1, 1) if mois == 12 else (annee, mois + 1)


def libelle_periode(periode: tuple) -> str:
    annee, mois = periode
    return f"{NOMS_MOIS[mois]} {annee}"


def mois_manquants_entre(depuis: tuple, jusqu_a: tuple) -> list:
    """Mois absents strictement entre deux périodes (bornes exclues)."""
    manquants = []
    courant = mois_suivant(depuis)
    while courant < jusqu_a:
        manquants.append(courant)
        courant = mois_suivant(courant)
    return manquants


def lire_mois_enregistres() -> list:
    """Périodes déjà présentes dans le classeur, triées.

    C'est ce qui permet à l'agent de « voir » le travail déjà fait : il lit les
    titres des onglets plutôt que leur contenu, ce qui reste rapide même sur un
    classeur chargé."""
    classeur = get_gspread_client().open_by_key(st.session_state.config["sheet_principale_id"])
    periodes = [
        periode for periode in (periode_du_titre_onglet(f.title) for f in classeur.worksheets())
        if periode
    ]
    return sorted(periodes)


def rafraichir_suivi(silencieux: bool = True) -> None:
    """Recharge l'état du suivi depuis le classeur, sans jamais interrompre le
    parcours : si le classeur est inaccessible, l'agent se met simplement en
    retrait plutôt que d'afficher une erreur bloquante."""
    try:
        st.session_state.mois_enregistres = lire_mois_enregistres()
        st.session_state.suivi_erreur = None
    except Exception as exc:
        st.session_state.mois_enregistres = None
        st.session_state.suivi_erreur = message_erreur_sheets(exc)
        if not silencieux:
            st.error(f"🚨 Suivi indisponible : {st.session_state.suivi_erreur}")


def periode_attendue(mois_enregistres) -> tuple:
    """Mois que l'agent attend ensuite : celui qui suit le dernier enregistré."""
    if not mois_enregistres:
        return None
    return mois_suivant(mois_enregistres[-1])


def diagnostic_suivi(mois_enregistres, mois_en_cours=None) -> list:
    """Messages de l'agent sur l'état du suivi : {niveau, texte}.

    `mois_en_cours` est la période détectée dans le lot en cours d'analyse ;
    quand elle est fournie, l'agent la confronte à ce qu'il attendait."""
    messages = []

    if mois_enregistres is None:
        return [{
            "niveau": "info",
            "texte": "Je n'ai pas encore pu consulter le classeur : le suivi des mois est indisponible.",
        }]

    attendue = periode_attendue(mois_enregistres)

    if not mois_enregistres:
        messages.append({
            "niveau": "info",
            "texte": "Aucun rapport mensuel n'est encore enregistré dans ce classeur. "
                     "Le premier lot que tu analyseras ouvrira le suivi.",
        })
    else:
        messages.append({
            "niveau": "info",
            "texte": f"{len(mois_enregistres)} mois déjà enregistré(s), jusqu'à "
                     f"**{libelle_periode(mois_enregistres[-1])}**. "
                     f"J'attends maintenant les données de **{libelle_periode(attendue)}**.",
        })

        trous = []
        for precedent, suivant in zip(mois_enregistres, mois_enregistres[1:]):
            trous.extend(mois_manquants_entre(precedent, suivant))
        if trous:
            messages.append({
                "niveau": "alerte",
                "texte": "Il manque des mois dans le suivi : "
                         + ", ".join(f"**{libelle_periode(t)}**" for t in trous) + ".",
            })

    if mois_en_cours:
        if mois_en_cours in mois_enregistres:
            messages.append({
                "niveau": "alerte",
                "texte": f"**{libelle_periode(mois_en_cours)}** figure déjà dans le classeur. "
                         "Écrire ce rapport remplacera les données existantes de ce mois.",
            })
        elif attendue is None:
            messages.append({
                "niveau": "succes",
                "texte": f"Ce lot ouvre le suivi sur **{libelle_periode(mois_en_cours)}**.",
            })
        elif mois_en_cours == attendue:
            messages.append({
                "niveau": "succes",
                "texte": f"Ce lot couvre bien **{libelle_periode(mois_en_cours)}**, "
                         "le mois que j'attendais : la continuité est respectée.",
            })
        elif mois_en_cours < attendue:
            messages.append({
                "niveau": "alerte",
                "texte": f"Ce lot couvre **{libelle_periode(mois_en_cours)}**, antérieur au mois "
                         f"attendu (**{libelle_periode(attendue)}**). Vérifie qu'il s'agit bien "
                         "des photos que tu voulais traiter.",
            })
        else:
            sautes = mois_manquants_entre(mois_enregistres[-1], mois_en_cours)
            messages.append({
                "niveau": "alerte",
                "texte": f"Ce lot couvre **{libelle_periode(mois_en_cours)}**, mais "
                         + ", ".join(f"**{libelle_periode(m)}**" for m in sautes)
                         + " n'a pas encore été traité. Tu peux poursuivre, ce mois restera à faire.",
            })

    return messages


def afficher_agent(messages: list, titre: str = "Assistant de suivi") -> None:
    """Affiche les messages de l'agent dans un encadré unique."""
    if not messages:
        return

    icones = {"info": "💬", "alerte": "⚠️", "succes": "✅"}
    lignes = "".join(
        f'<div class="ligne-agent {m["niveau"]}">{icones.get(m["niveau"], "•")} '
        f'{convertir_gras(echapper_html(m["texte"]))}</div>'
        for m in messages
    )
    st.markdown(
        f'<div class="panneau-agent"><div class="titre-agent">🤖 {echapper_html(titre)}</div>{lignes}</div>',
        unsafe_allow_html=True,
    )


def convertir_gras(texte: str) -> str:
    """Convertit la syntaxe **gras** en HTML, après échappement du texte."""
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", texte)


# ============================================================
# INSTRUCTIONS EN LANGAGE NATUREL
# ------------------------------------------------------------
# Le client écrit ce qu'il veut ; l'IA se contente de TRADUIRE sa
# phrase en une action structurée parmi une liste fermée. L'action
# est ensuite prévisualisée, puis appliquée par du code ordinaire
# seulement si l'utilisateur confirme. À aucun moment l'IA ne
# modifie directement les données ni n'écrit dans Google Sheets.
# ============================================================
ACTIONS_AUTORISEES = {
    "modifier_recette", "modifier_depense",
    "supprimer_recette", "supprimer_depense",
    "ajouter_recette", "ajouter_depense",
    "renommer_depense",
}


def construire_prompt_instruction(instruction: str, apercu_donnees: str) -> str:
    return f"""Tu traduis une instruction en français en UNE action structurée, pour une
application de gestion de rapports de taxi. Tu ne fais que traduire : tu n'exécutes rien.

Données actuellement chargées :
{apercu_donnees}

Instruction de l'utilisateur :
"{instruction}"

Réponds UNIQUEMENT par un objet JSON, sans texte autour, choisi parmi :

{{"action":"modifier_recette","date":"JJ/MM/AA","montant":<nombre>}}
{{"action":"modifier_depense","titre":"<titre existant>","date":"JJ/MM/AA ou null","montant":<nombre>}}
{{"action":"supprimer_recette","date":"JJ/MM/AA"}}
{{"action":"supprimer_depense","titre":"<titre existant>","date":"JJ/MM/AA ou null"}}
{{"action":"ajouter_recette","date":"JJ/MM/AA","montant":<nombre>}}
{{"action":"ajouter_depense","titre":"<titre>","date":"JJ/MM/AA","montant":<nombre>}}
{{"action":"renommer_depense","ancien_titre":"<titre existant>","nouveau_titre":"<nouveau>"}}
{{"action":"inconnue","raison":"<pourquoi tu ne peux pas traduire cette demande>"}}

Règles :
- Les dates s'écrivent toujours JJ/MM/AA (ex : 12/04/26).
- Les montants sont des nombres sans espace ni devise.
- Si l'instruction est ambiguë, ne concerne pas ces données, ou demande autre chose
  que les actions ci-dessus, réponds avec "inconnue" en expliquant brièvement.
- N'invente jamais une date ou un titre absent des données ci-dessus, sauf pour un ajout."""


def apercu_donnees_pour_agent() -> str:
    """Résumé compact des données chargées, transmis à l'IA pour qu'elle puisse
    rattacher l'instruction aux bonnes lignes."""
    morceaux = []
    for indice in sorted(st.session_state.donnees_filtrees or {}):
        recettes = obtenir_donnees_semaine(indice, "recettes")
        depenses = obtenir_donnees_semaine(indice, "depenses")
        morceaux.append(f"Semaine {indice + 1} :")
        for r in recettes:
            morceaux.append(f"  recette {r.get('date')} = {r.get('montant')}")
        for d in depenses:
            morceaux.append(f"  dépense « {d.get('titre')} » {d.get('date') or 'sans date'} = {d.get('montant')}")
    return "\n".join(morceaux) or "(aucune donnée chargée)"


def interpreter_instruction(instruction: str) -> dict:
    """Traduit l'instruction en action structurée via Gemini."""
    api_key = get_gemini_key()
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    payload = {
        "contents": [{"parts": [{"text": construire_prompt_instruction(instruction, apercu_donnees_pour_agent())}]}],
        "generationConfig": {"temperature": 0, "response_mime_type": "application/json"},
    }
    reponse = requests.post(url, json=payload, timeout=60)
    if reponse.status_code >= 400:
        message = ""
        try:
            message = reponse.json().get("error", {}).get("message", "")
        except Exception:
            pass
        raise RuntimeError(masquer_cle_api(f"Erreur Gemini {reponse.status_code} : {message}"))

    texte = reponse.json()["candidates"][0]["content"]["parts"][0]["text"]
    texte = re.sub(r"^```(?:json)?|```$", "", texte.strip(), flags=re.MULTILINE).strip()
    action = json.loads(texte)
    if action.get("action") not in ACTIONS_AUTORISEES and action.get("action") != "inconnue":
        return {"action": "inconnue", "raison": "Action non reconnue par l'application."}
    return action


def trouver_ligne(champ: str, criteres: dict):
    """Localise une ligne dans les données chargées.
    Renvoie (indice_semaine, position, ligne) ou (None, None, None)."""
    for indice in sorted(st.session_state.donnees_filtrees or {}):
        for position, ligne in enumerate(obtenir_donnees_semaine(indice, champ)):
            if "date" in criteres and criteres["date"]:
                if parser_date(str(ligne.get("date", ""))) != parser_date(criteres["date"]):
                    continue
            if "titre" in criteres and criteres["titre"]:
                if normaliser_texte(ligne.get("titre", "")) != normaliser_texte(criteres["titre"]):
                    continue
            return indice, position, ligne
    return None, None, None


def decrire_action(action: dict) -> str:
    """Phrase lisible décrivant ce que l'action va faire."""
    a = action.get("action")
    if a == "modifier_recette":
        return f"Remplacer le montant de la recette du {action['date']} par {formater_montant(action['montant'])} FCFA."
    if a == "modifier_depense":
        return f"Remplacer le montant de la dépense « {action['titre']} » par {formater_montant(action['montant'])} FCFA."
    if a == "supprimer_recette":
        return f"Supprimer la recette du {action['date']}."
    if a == "supprimer_depense":
        return f"Supprimer la dépense « {action['titre']} »."
    if a == "ajouter_recette":
        return f"Ajouter une recette de {formater_montant(action['montant'])} FCFA au {action['date']}."
    if a == "ajouter_depense":
        return f"Ajouter la dépense « {action['titre']} » de {formater_montant(action['montant'])} FCFA au {action['date']}."
    if a == "renommer_depense":
        return f"Renommer la dépense « {action['ancien_titre']} » en « {action['nouveau_titre']} »."
    return action.get("raison", "Instruction non comprise.")


def appliquer_action(action: dict) -> tuple:
    """Exécute l'action sur les données chargées. Renvoie (succès, message)."""
    a = action.get("action")

    def enregistrer(indice, champ, lignes):
        st.session_state.donnees_filtrees[indice][champ] = lignes
        st.session_state.donnees_editees.pop(indice, None)
        st.session_state.pop(f"editeur_{champ}_{indice}", None)
        st.session_state.semaines_validees.discard(indice)

    if a in ("modifier_recette", "supprimer_recette"):
        indice, position, _ = trouver_ligne("recettes", {"date": action.get("date")})
        if indice is None:
            return False, f"Aucune recette trouvée au {action.get('date')}."
        lignes = list(obtenir_donnees_semaine(indice, "recettes"))
        if a == "modifier_recette":
            lignes[position] = {**lignes[position], "montant": action["montant"]}
        else:
            lignes.pop(position)
        enregistrer(indice, "recettes", lignes)
        return True, f"Semaine {indice + 1} mise à jour."

    if a in ("modifier_depense", "supprimer_depense", "renommer_depense"):
        criteres = ({"titre": action.get("ancien_titre")} if a == "renommer_depense"
                    else {"titre": action.get("titre"), "date": action.get("date")})
        indice, position, _ = trouver_ligne("depenses", criteres)
        if indice is None:
            cible = criteres.get("titre")
            return False, f"Aucune dépense « {cible} » trouvée."
        lignes = list(obtenir_donnees_semaine(indice, "depenses"))
        if a == "modifier_depense":
            lignes[position] = {**lignes[position], "montant": action["montant"]}
        elif a == "renommer_depense":
            lignes[position] = {**lignes[position], "titre": action["nouveau_titre"]}
        else:
            lignes.pop(position)
        enregistrer(indice, "depenses", lignes)
        return True, f"Semaine {indice + 1} mise à jour."

    if a in ("ajouter_recette", "ajouter_depense"):
        date_ajout = parser_date(action.get("date", ""))
        if date_ajout is None:
            return False, "La date de l'ajout n'a pas pu être lue."
        # On rattache l'ajout à la semaine dont la période contient cette date.
        indice_cible = None
        for indice in sorted(st.session_state.donnees_filtrees or {}):
            dates = [parser_date(str(r.get("date", ""))) for r in obtenir_donnees_semaine(indice, "recettes")]
            dates = [d for d in dates if d]
            if dates and min(dates) <= date_ajout <= max(dates):
                indice_cible = indice
                break
        if indice_cible is None:
            indice_cible = min(st.session_state.donnees_filtrees or {0: None})

        champ = "recettes" if a == "ajouter_recette" else "depenses"
        lignes = list(obtenir_donnees_semaine(indice_cible, champ))
        if a == "ajouter_recette":
            lignes.append({"date": action["date"], "montant": action["montant"]})
        else:
            lignes.append({"titre": action["titre"], "montant": action["montant"], "date": action["date"]})
        enregistrer(indice_cible, champ, lignes)
        return True, f"Ajouté à la semaine {indice_cible + 1}."

    return False, "Instruction non applicable."


def zone_instruction_agent() -> None:
    """Champ où le client formule une demande en français, avec confirmation
    avant toute modification."""
    st.markdown("#### 💬 Demander une modification à l'assistant")
    st.caption(
        "Écris ta demande en français, par exemple : « la recette du 12/04 est de 20000 », "
        "« supprime la dépense vidange » ou « renomme videnge en vidange ». "
        "L'assistant te montrera ce qu'il a compris avant d'appliquer quoi que ce soit."
    )

    instruction = st.text_input(
        "Instruction",
        key="instruction_agent",
        placeholder="Ex : la recette du 12/04 est de 20000",
        label_visibility="collapsed",
    )

    if st.button("🤖 Interpréter", disabled=not instruction.strip()):
        with st.spinner("L'assistant analyse ta demande..."):
            try:
                st.session_state.action_en_attente = interpreter_instruction(instruction)
            except Exception as exc:
                st.session_state.action_en_attente = None
                st.error(f"🚨 {decrire_erreur(exc)}")

    action = st.session_state.get("action_en_attente")
    if not action:
        return

    if action.get("action") == "inconnue":
        afficher_agent([{
            "niveau": "alerte",
            "texte": action.get("raison", "Je n'ai pas compris cette demande.")
                     + " Reformule, ou modifie directement le tableau ci-dessus.",
        }], titre="Assistant — demande non comprise")
        if st.button("Fermer", key="fermer_action_inconnue"):
            st.session_state.action_en_attente = None
            st.rerun()
        return

    afficher_agent([{
        "niveau": "info",
        "texte": "Voici ce que j'ai compris : " + decrire_action(action),
    }], titre="Assistant — confirmation requise")

    col_ok, col_non = st.columns(2)
    with col_ok:
        if st.button("✅ Appliquer", type="primary", use_container_width=True, key="appliquer_action"):
            succes, message = appliquer_action(action)
            st.session_state.action_en_attente = None
            if succes:
                st.toast(f"✅ {message}")
            else:
                st.session_state.dernier_echec_agent = message
            st.rerun()
    with col_non:
        if st.button("↩️ Annuler", use_container_width=True, key="annuler_action"):
            st.session_state.action_en_attente = None
            st.rerun()


def nom_onglet_rapport_cible(rapport: dict) -> str:
    """Nom de l'onglet où écrire le rapport mensuel : celui choisi dans
    ⚙️ Paramètres s'il est renseigné, sinon un nom automatique par mois."""
    choisi = (st.session_state.config.get("nom_onglet_rapport") or "").strip()
    return choisi or NOM_ONGLET_RAPPORT.format(
        nom_mois=NOMS_MOIS[rapport["mois"]].capitalize(), annee=rapport["annee"]
    )


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
    mettre_en_forme_rapport(feuille, len(lignes), premiere_ligne_synthese, mois=rapport["mois"])

    try:
        positionner_onglet_chronologiquement(classeur, feuille, rapport["annee"], rapport["mois"])
    except Exception:
        # Le rangement est un confort : son échec ne doit pas faire perdre
        # un rapport correctement écrit.
        pass

    return {
        "titre": nom_onglet,
        "classeur": classeur.title,
        "url": f"{classeur.url}#gid={feuille.id}",
        "remplace": existant is not None,
    }


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
    "mois_enregistres": None,
    "suivi_erreur": None,
    "suivi_charge": False,
    "action_en_attente": None,
    "dernier_echec_agent": None,
    "signature_lot": None,
    "fichiers_lot": None,
    "page": "📤 Nouveau rapport",
    "lot_donnees": None,
    "donnees_filtrees": None,
    "mois_confirme": False,
    "sous_page": 0,
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


def base_editeur(indice: int, champ: str) -> list:
    """Valeur de départ du tableau éditable.

    st.data_editor applique ses modifications par-dessus la valeur qu'on lui
    fournit : tant que le widget vit, il faut donc lui redonner la même base,
    sinon les corrections seraient appliquées deux fois.

    Mais Streamlit efface l'état d'un widget dès qu'il cesse d'être affiché —
    ce qui arrive à chaque changement de semaine. Au retour, la base d'origine
    réapparaissait et écrasait les corrections. On détecte ce cas (données
    éditées présentes alors que le widget n'a plus d'état) pour promouvoir les
    données corrigées en nouvelle base."""
    if not st.session_state.donnees_filtrees or indice not in st.session_state.donnees_filtrees:
        return []

    editees = st.session_state.donnees_editees.get(indice)
    widget_vivant = f"editeur_{champ}_{indice}" in st.session_state

    if editees is not None and not widget_vivant:
        st.session_state.donnees_filtrees[indice][champ] = editees[champ]

    return st.session_state.donnees_filtrees[indice][champ]


def obtenir_donnees_semaine(indice: int, champ: str) -> list:
    """Données de la semaine telles qu'elles serviront au bilan : la version
    corrigée si la page a été ouverte, sinon celle issue de l'analyse."""
    editees = st.session_state.donnees_editees.get(indice)
    if editees is not None:
        return editees[champ]
    if st.session_state.donnees_filtrees and indice in st.session_state.donnees_filtrees:
        return st.session_state.donnees_filtrees[indice][champ]
    return []


def afficher_bilan_mensuel(rapport: dict) -> None:
    """Affiche le détail d'un bilan mensuel (métriques + dépenses par titre)."""
    nom_mois = NOMS_MOIS[rapport["mois"]]
    st.markdown(f"### 🎉 Bilan du mois de {nom_mois} {rapport['annee']}")

    c1, c2, c3 = st.columns(3)
    c1.metric("Jours travaillés", formater_jours(rapport["jours_travailles"]))
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
        ["📤 Nouveau rapport", "⚙️ Paramètres"],
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

    nouveau_tarif = st.number_input(
        "Recette d'une journée pleine (FCFA)",
        min_value=1,
        step=500,
        value=int(st.session_state.config.get("tarif_journalier") or TARIF_JOURNALIER_DEFAUT),
        help="Sert à calculer le nombre de jours travaillés : recette totale ÷ ce tarif. "
             "Une journée à moitié travaillée compte ainsi pour 0,5 jour.",
    )

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
            "tarif_journalier": int(nouveau_tarif),
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
            # Repart d'une connexion neuve : évite les erreurs « No grid with id »
            # qui surviennent quand la structure du classeur a changé depuis.
            get_gspread_client.clear()
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

    # Premier passage : l'agent consulte le classeur pour savoir où en est le suivi.
    if not st.session_state.suivi_charge:
        with st.spinner("Consultation du classeur..."):
            rafraichir_suivi()
        st.session_state.suivi_charge = True

    # Période du lot en cours, dès qu'elle est connue, pour que l'agent la
    # confronte au mois qu'il attendait.
    _mois_du_lot = None
    if st.session_state.lot_donnees:
        _debut_lot, _ = determiner_bornes_mois(st.session_state.lot_donnees)
        if _debut_lot:
            _mois_du_lot = (_debut_lot.year, _debut_lot.month)

    afficher_agent(diagnostic_suivi(st.session_state.mois_enregistres, _mois_du_lot))

    col_maj, col_vide = st.columns([1, 3])
    with col_maj:
        if st.button("🔄 Actualiser le suivi", use_container_width=True):
            rafraichir_suivi(silencieux=False)
            st.rerun()

    st.write(f"Uploade de 1 à {MAX_IMAGES_PAR_LOT} rapports hebdomadaires (glisser-déposer possible).")

    fichiers_uploades = st.file_uploader(
        f"Sélectionner les images des rapports (jusqu'à {MAX_IMAGES_PAR_LOT})",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.cle_uploader}",
    )

    if fichiers_uploades:
        if len(fichiers_uploades) > MAX_IMAGES_PAR_LOT:
            st.warning(f"⚠️ Maximum {MAX_IMAGES_PAR_LOT} images à la fois. Seules les {MAX_IMAGES_PAR_LOT} premières seront prises en compte.")
            fichiers_uploades = fichiers_uploades[:MAX_IMAGES_PAR_LOT]

        signature = tuple((f.name, f.size) for f in fichiers_uploades)
        if st.session_state.signature_lot != signature:
            for cle in list(st.session_state.keys()):
                if cle.startswith("editeur_recettes_") or cle.startswith("editeur_depenses_"):
                    del st.session_state[cle]
            # Copie en mémoire : le lot doit survivre à un passage par la page
            # Paramètres, qui vide le file_uploader.
            st.session_state.fichiers_lot = memoriser_fichiers(fichiers_uploades)
            st.session_state.signature_lot = signature
            st.session_state.lot_donnees = None
            st.session_state.donnees_filtrees = None
            st.session_state.mois_confirme = False
            st.session_state.sous_page = 0
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
            mot_rapport = "rapport" if total_fichiers == 1 else "rapports"
            st.write(
                f"{total_fichiers} {mot_rapport} vont être lus par l'IA afin d'identifier le mois "
                "couvert, avant tout enregistrement."
                if total_fichiers > 1 else
                "Le rapport va être lu par l'IA afin d'identifier le mois couvert, avant tout enregistrement."
            )
            st.caption(
                f"Modèle utilisé : `{GEMINI_MODEL}`. Les images sont envoyées avec quelques secondes "
                "d'écart entre chacune pour rester dans les limites du palier gratuit Gemini."
            )
            if total_fichiers < MAX_IMAGES_PAR_LOT:
                afficher_agent([{
                    "niveau": "info",
                    "texte": f"Ce lot ne contient que **{total_fichiers} {mot_rapport}** sur "
                             f"{MAX_IMAGES_PAR_LOT} possibles : le bilan ne couvrira que les jours "
                             "effectivement présents. Tu pourras compléter le mois plus tard en "
                             "réécrivant le rapport avec les semaines manquantes.",
                }], titre="Assistant — lot partiel")

            if st.button(f"🔍 Analyser le lot ({total_fichiers} {mot_rapport})", type="primary"):
                with st.spinner(f"Analyse Gemini de {total_fichiers} {mot_rapport} en cours..."):
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

            # Montants que l'IA n'a pas su lire : ils valent 0 pour l'instant et
            # fausseraient le bilan s'ils passaient inaperçus.
            total_non_lus = sum(
                donnees.get("montants_non_lus", 0)
                for donnees in st.session_state.lot_donnees
                if isinstance(donnees, dict)
            )
            if total_non_lus:
                st.warning(
                    f"⚠️ {total_non_lus} montant(s) n'ont pas pu être lus et valent actuellement 0. "
                    "Corrige-les à l'étape de vérification, sinon le bilan sera sous-évalué."
                )

            # Valeurs manifestement suspectes, signe d'une lecture erronée.
            suspects = []
            for indice, donnees in enumerate(st.session_state.lot_donnees):
                if not isinstance(donnees, dict) or "erreur" in donnees:
                    continue
                for recette in donnees.get("recettes_journalieres", []):
                    montant = parser_montant(recette.get("montant"))
                    if montant is None:
                        continue
                    if montant < 0:
                        suspects.append(f"Semaine {indice + 1} : recette négative le {recette.get('date')}")
                    elif montant > tarif_journalier() * 5:
                        suspects.append(
                            f"Semaine {indice + 1} : recette de {formater_montant(montant)} F le "
                            f"{recette.get('date')}, très au-dessus d'une journée type"
                        )
            if suspects:
                st.warning("⚠️ Valeurs à vérifier :\n\n" + "\n".join(f"- {s}" for s in suspects[:8]))

            # --------------------------------------------------------
            # 2) Périodes manquantes : combler avec une photo, ou ignorer
            # --------------------------------------------------------
            st.divider()
            st.markdown("#### 🗓️ Vérification des périodes manquantes")
            periodes_manquantes = detecter_periodes_manquantes(
                st.session_state.lot_donnees, debut_mois, fin_mois
            )
            periodes_a_traiter = [
                periode for periode in periodes_manquantes
                if f"{periode[0].isoformat()}_{periode[1].isoformat()}" not in st.session_state.trous_ignores
            ]

            if periodes_a_traiter:
                jours_manquants = sum(
                    (fin - debut).days + 1 for debut, fin in periodes_a_traiter
                )
                messages_periodes = [{
                    "niveau": "alerte",
                    "texte": f"Il manque **{jours_manquants} jour(s)** sur le mois de "
                             f"**{NOMS_MOIS[debut_mois.month]} {debut_mois.year}**.",
                }]
                messages_periodes += [{
                    "niveau": "alerte",
                    "texte": f"Période manquante : du **{debut.strftime('%d/%m')}** au "
                             f"**{fin.strftime('%d/%m')}** "
                             f"({(fin - debut).days + 1} jour(s)).",
                } for debut, fin in periodes_a_traiter]
                messages_periodes.append({
                    "niveau": "info",
                    "texte": "Pour chacune : charge la photo correspondante, saisis les données à "
                             "la main, ou poursuis sans elle.",
                })
                afficher_agent(messages_periodes, titre="Assistant — périodes manquantes")

                for debut_trou, fin_trou in periodes_a_traiter:
                    cle_trou = f"{debut_trou.isoformat()}_{fin_trou.isoformat()}"
                    with st.expander(
                        f"Compléter la période du {debut_trou.strftime('%d/%m/%y')} "
                        f"au {fin_trou.strftime('%d/%m/%y')}"
                    ):
                        mode = st.radio(
                            "Comment veux-tu compléter cette période ?",
                            ["📷 Charger la photo", "✍️ Saisir les données", "⏭️ Poursuivre sans"],
                            horizontal=True,
                            key=f"mode_comblement_{cle_trou}",
                        )

                        # ---- Mode photo ----
                        if mode == "📷 Charger la photo":
                            fichier_combler = st.file_uploader(
                                "Photo du rapport de cette période (elle sera vérifiée automatiquement)",
                                type=["jpg", "jpeg", "png"],
                                key=f"upload_comblement_{cle_trou}",
                            )
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
                                            st.success("✅ Photo ajoutée : elle couvre bien la période manquante.")
                                            st.rerun()
                                        else:
                                            periode_trouvee = donnees_comblees.get("periode_hebdo") or "non détectée"
                                            st.error(
                                                f"🚨 Cette photo couvre la période « {periode_trouvee} », qui ne "
                                                f"correspond pas à la période manquante "
                                                f"({debut_trou.strftime('%d/%m/%y')} - {fin_trou.strftime('%d/%m/%y')}). "
                                                "Vérifie que c'est la bonne photo."
                                            )
                                    except Exception as e:
                                        st.error(f"🚨 Erreur lors de l'analyse : {decrire_erreur(e)}")

                        # ---- Mode saisie manuelle ----
                        elif mode == "✍️ Saisir les données":
                            st.caption(
                                "Les jours de la période sont pré-remplis. Renseigne les recettes "
                                "(laisse à 0 un jour non travaillé) et ajoute les dépenses éventuelles."
                            )
                            base_recettes = [
                                {"date": jour.strftime("%d/%m/%y"), "montant": 0}
                                for jour in jours_de_periode(debut_trou, fin_trou)
                            ]
                            recettes_saisies = st.data_editor(
                                base_recettes,
                                num_rows="dynamic",
                                column_config={
                                    "date": st.column_config.TextColumn("Date (JJ/MM/AA)", required=True),
                                    "montant": st.column_config.NumberColumn("Recette (FCFA)", required=True, step=500),
                                },
                                use_container_width=True,
                                key=f"saisie_recettes_{cle_trou}",
                            )
                            st.caption("Dépenses de la période (facultatif) :")
                            depenses_saisies = st.data_editor(
                                [],
                                num_rows="dynamic",
                                column_config={
                                    "titre": st.column_config.TextColumn("Titre de la dépense", required=True),
                                    "montant": st.column_config.NumberColumn("Montant (FCFA)", required=True, step=500),
                                    "date": st.column_config.TextColumn("Date (JJ/MM/AA)", required=False),
                                },
                                use_container_width=True,
                                key=f"saisie_depenses_{cle_trou}",
                            )

                            recettes_propres = nettoyer_lignes_editeur(
                                normaliser_lignes_editeur(recettes_saisies), CHAMPS_RECETTES
                            )
                            depenses_propres = nettoyer_lignes_editeur(
                                normaliser_lignes_editeur(depenses_saisies), CHAMPS_DEPENSES
                            )
                            total_saisi = calculer_total_recettes(recettes_propres)
                            st.caption(
                                f"Total saisi : **{formater_montant(total_saisi)} FCFA** de recettes, "
                                f"**{formater_montant(calculer_total_depenses(depenses_propres))} FCFA** de dépenses."
                            )

                            if st.button(
                                "✅ Ajouter ces données au mois",
                                key=f"valider_saisie_{cle_trou}",
                                disabled=not recettes_propres,
                                use_container_width=True,
                                type="primary",
                            ):
                                st.session_state.lot_donnees.append({
                                    "periode_hebdo": f"{debut_trou.strftime('%d/%m/%y')} - {fin_trou.strftime('%d/%m/%y')}",
                                    "recettes_journalieres": recettes_propres,
                                    "depenses": depenses_propres,
                                    "saisie_manuelle": True,
                                })
                                # Aucune photo associée : on mémorise None pour
                                # conserver l'alignement entre rapports et images.
                                st.session_state.fichiers_combles[cle_trou] = None
                                st.success("✅ Données saisies ajoutées au mois.")
                                st.rerun()

                        # ---- Mode « poursuivre sans » ----
                        else:
                            st.caption(
                                "Cette période sera absente du bilan : les jours concernés ne seront "
                                "comptés ni en recettes ni en jours travaillés."
                            )
                            if st.button(
                                "⏭️ Poursuivre sans cette période",
                                key=f"ignorer_comblement_{cle_trou}",
                                use_container_width=True,
                            ):
                                st.session_state.trous_ignores.add(cle_trou)
                                st.rerun()
            else:
                st.success("✅ Le mois est entièrement couvert (ou les périodes absentes ont été acceptées).")

            tous_les_trous_geres = len(periodes_a_traiter) == 0

            peut_continuer = tous_les_trous_geres and poursuivre_malgre_doublon

            def preparer_donnees_filtrees() -> None:
                """Prépare les tableaux hebdomadaires à partir de l'analyse."""
                for cle in list(st.session_state.keys()):
                    if cle.startswith("editeur_recettes_") or cle.startswith("editeur_depenses_"):
                        del st.session_state[cle]
                filtres = {}
                for indice, donnees in enumerate(st.session_state.lot_donnees):
                    if "erreur" in donnees:
                        continue
                    filtre = filtrer_donnees_par_mois(donnees, debut_mois, fin_mois)
                    filtres[indice] = {
                        "recettes": filtre["recettes_journalieres"],
                        "depenses": filtre["depenses"],
                    }
                st.session_state.donnees_filtrees = filtres
                # Les éditions d'un lot précédent ne doivent pas être reprises :
                # les tableaux repartent des données fraîchement analysées.
                st.session_state.donnees_editees = {}
                st.session_state.mois_confirme = True

            col_go, col_retour = st.columns(2)
            with col_go:
                if st.button(
                    "➡️ Continuer vers la vérification des rapports",
                    type="primary",
                    disabled=not peut_continuer,
                ):
                    preparer_donnees_filtrees()
                    st.session_state.sous_page = 0
                    st.rerun()
            with col_retour:
                if st.button("↩️ Recommencer l'analyse du lot"):
                    st.session_state.lot_donnees = None
                    st.session_state.trous_ignores = set()
                    st.session_state.fichiers_combles = {}
                    st.session_state.ordre_chronologique = None
                    st.rerun()

            st.caption(
                "Si tu fais confiance à la lecture de l'IA et n'as aucune correction à apporter, "
                "tu peux passer directement au bilan. Les données resteront consultables et "
                "modifiables ensuite, semaine par semaine."
            )
            if st.button(
                "⏭️ Passer la vérification et établir le bilan",
                use_container_width=True,
                disabled=not peut_continuer,
            ):
                preparer_donnees_filtrees()
                # Toutes les semaines exploitables sont validées d'office, et
                # le bilan est calculé dans la foulée à partir des données
                # issues de l'analyse, sans passer par les pages hebdomadaires.
                indices_exploitables = [
                    indice for indice, donnees in enumerate(st.session_state.lot_donnees)
                    if "erreur" not in donnees
                ]
                st.session_state.semaines_validees = set(indices_exploitables)

                recettes_combinees = []
                depenses_combinees = []
                for indice in indices_exploitables:
                    recettes_combinees.extend(obtenir_donnees_semaine(indice, "recettes"))
                    depenses_combinees.extend(obtenir_donnees_semaine(indice, "depenses"))

                st.session_state.rapport_fige = calculer_bilan_mensuel_agrege(
                    recettes_combinees, depenses_combinees, debut_mois.year, debut_mois.month
                )
                st.session_state.bilan_etabli = True
                st.session_state.sous_page = len(st.session_state.lot_donnees)
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
            fichier_i = tous_les_fichiers[i] if i < len(tous_les_fichiers) else None
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
                    if fichier_i is not None:
                        st.image(fichier_i, caption=fichier_i.name, use_container_width=True)
                    else:
                        st.info(
                            "✍️ Semaine saisie manuellement : il n'y a pas de photo à afficher. "
                            "Les données restent modifiables ci-contre."
                        )

                with col_data:
                    origine = ("Période saisie" if donnees_brutes.get("saisie_manuelle")
                               else "Période brute détectée sur l'image")
                    st.caption(
                        f"{origine} : **{donnees_brutes.get('periode_hebdo', '—')}**. "
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
                        base_editeur(i, "recettes"),
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
                        base_editeur(i, "depenses"),
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
                    if st.session_state.dernier_echec_agent:
                        afficher_agent([{
                            "niveau": "alerte",
                            "texte": st.session_state.dernier_echec_agent,
                        }], titre="Assistant — modification impossible")
                        st.session_state.dernier_echec_agent = None
                    zone_instruction_agent()

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

            # Filet de sécurité : le rapport reste récupérable même si Google
            # Sheets est indisponible (quota, panne, partage retiré).
            contenu_csv = "\n".join(
                ";".join(str(cellule) for cellule in ligne) for ligne in lignes_apercu
            )
            st.download_button(
                "⬇️ Télécharger ce rapport (CSV)",
                data=contenu_csv.encode("utf-8-sig"),
                file_name=f"rapport_{NOMS_MOIS[rapport['mois']]}_{rapport['annee']}.csv",
                mime="text/csv",
                help="Sauvegarde locale, utile si l'écriture dans Google Sheets échoue.",
            )

            if st.session_state.rapport_mensuel_cree:
                infos = st.session_state.rapport_mensuel_cree
                st.success(
                    f"✅ Rapport écrit dans l'onglet « {infos['titre']} » "
                    f"du classeur « {infos['classeur']} »."
                )
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
                            # Le classeur vient de changer : l'agent doit en
                            # tenir compte pour annoncer le prochain mois attendu.
                            rafraichir_suivi()
                            st.rerun()
                        except Exception as e:
                            st.error(f"🚨 Création impossible : {message_erreur_sheets(e)}")

            st.divider()
            if st.button("📥 Traiter un nouveau lot d'images (mois suivant)"):
                reinitialiser_lot()
                st.rerun()