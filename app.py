import streamlit as st
import requests
import base64
import json
import gspread
import datetime 
import base64
import os
import streamlit as st 
from google.oauth2.service_account import Credentials

# --- CONFIGURATION INITIALE ---
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

if os.path.exists("credentials.json"):
    # Mode LOCAL (sur ton PC)
    API_KEY = "AQ.Ab8RN6LzmBzLmmus_jP5gcPgIOBeu4hUA971zpojR_0vwAxSEQ"
    creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
else:
    # Mode CLOUD (sur Streamlit Cloud)
    API_KEY = st.secrets["GEMINI_API_KEY"]
    creds = Credentials.from_service_account_info(st.secrets["google_credentials"], scopes=scope)

client = gspread.authorize(creds)

# .get_worksheet(0) sélectionne le tout premier onglet, peu importe son nom
feuille = client.open_by_key("1aoLjyA10sSb1m1hEr0bDPj8gxB2CZ22RLnkKKkrCpMM").get_worksheet(0)
feuille_principale = client.open_by_key("1BNlV17OasazXtFPLbp64xHwJ2RqBjPqs97_Adi4Cbuo").get_worksheet(0)
# --- INTERFACE STREAMLIT ---
st.title("🚖 Taxi Dashboard - Automatisation")
st.write("Uploade le rapport hebdomadaire pour l'envoyer dans Google Sheets.") 

file = st.file_uploader("Sélectionner l'image du rapport", type=["jpg", "jpeg", "png"])

if file is not None:
    st.image(file, caption="Rapport sélectionné", use_container_width=True)
    
    if st.button("🚀 Analyser et Enregistrer"):
        with st.spinner("Analyse Gemini et écriture en cours..."):
            try:
                # 1. Encodage de l'image en Base64
                image_bytes = file.read()
                image_b64 = base64.b64encode(image_bytes).decode("utf-8")
                
                # 2. Préparation de la requête Gemini
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={API_KEY}"
                
                payload = {
                   "contents": [
                       {
                           "parts": [
                              {
                                "text": "Analyse cette image de rapport hebdomadaire et extrait les informations financières sous forme d'un objet JSON strict avec les clés : 'periode_hebdo' (au format exact 'JJ/MM/AA - JJ/MM/AA' basé sur la date de début et de fin écrite sur le cahier), 'total_recettes', 'total_depenses', 'solde_net'. Rends uniquement le JSON sans texte autour."
                              },
                              {
                                "inline_data": {
                                   "mime_type": "image/jpeg",
                                   "data": image_b64   
                                }
                              }
                            ]
                       }
                    ]   
                }
                # 3. Envoi de la requête à l'API Gemini
                response = requests.post(url, json=payload)
                res_json = response.json()
                
                # 4. Vérifications de sécurité sur la réponse de l'API
                if "error" in res_json:
                    st.error(f"❌ Erreur API Gemini : {res_json['error']['message']}")
                    st.stop()
                    
                if "candidates" not in res_json:
                    st.error("❌ Gemini n'a pas renvoyé de texte. Le document est peut-être illisible.")
                    st.stop()
                
                # 5. Extraction du texte JSON et nettoyage des backticks
                texte_json = res_json['candidates'][0]['content']['parts'][0]['text']
                texte_json = texte_json.replace("```json", "").replace("```", "").strip()
                
                # 6. Conversion en dictionnaire Python
                donnees = json.loads(texte_json)
                
                # 7. Affichage des résultats sur l'interface
                st.success("Analyse réussie !")
                st.json(donnees)
                
                # 8. Écriture automatique dans Google Sheets
                # --- CALCUL AUTOMATIQUE DE LA SEMAINE EN COURS ---
                aujourdhui = datetime.date.today()
                # Trouve le lundi de la semaine actuelle
                debut_semaine = aujourdhui - datetime.timedelta(days=aujourdhui.weekday())
                # Trouve le dimanche de la semaine actuelle
                fin_semaine = debut_semaine + datetime.timedelta(days=6)

                # Formate la période exactement comme sur ton cahier (ex: "15/06/26 - 21/06/26")
                periode_hebdo = f"{debut_semaine.strftime('%d/%m/%y')} - {fin_semaine.strftime('%d/%m/%y')}"


                # --- ÉCRITURE DANS GOOGLE SHEETS (Alignement parfait A, B, C, D) ---
                feuille.append_row([
                donnees.get("periode_hebdo"),    # Colonne A : Périodes Hebdomadaires
                donnees.get("total_recettes"),   # Colonne B : Recettes Totales
                donnees.get("total_depenses"),   # Colonne C : Dépenses Totales
                donnees.get("solde_net")         # Colonne D : Soldes finals
                ])
                
                st.balloons()
                st.success("✨ Données ajoutées à Google Sheets avec succès !")
                # --- CONSOLIDATION MENSUELLE AUTOMATIQUE ---
                toutes_les_lignes = feuille.get_all_values()

                # Si la feuille contient au moins 5 lignes (1 entête + 4 semaines)
                if len(toutes_les_lignes) >= 5:
                    total_recettes_mois = 0

                    # Somme des recettes de la colonne B
                    for ligne in toutes_les_lignes[1:]:
                        try:
                            total_recettes_mois += float(ligne[1])
                        except (ValueError, IndexError):
                            pass

                    # Correspondance auto avec les colonnes basée sur la date réelle des rapports
                    try:
                        # toutes_les_lignes[1][0] contient la première période (ex: "15/06/26 - 21/06/26")
                        periode_texte = toutes_les_lignes[1][0]
                        date_debut = periode_texte.split(" - ")[0]  # Extrait "15/06/26"
                        mois_actuel = int(date_debut.split("/")[1])  # Extrait le mois (ex: 6 pour Juin)
                    except Exception:
                        # Sécurité : si la lecture de la date échoue, on se replie sur le mois en cours
                        mois_actuel = datetime.date.today().month

                    colonne_mois = mois_actuel + 3

                    # Mise à jour de la ligne 4 ("15k journalier") en mode CUMULATIF
                    valeur_actuelle = feuille_principale.cell(4, colonne_mois).value

                    if valeur_actuelle:
                        valeur_nettoyee = str(valeur_actuelle).replace("CFA", "").replace(",", "").replace(" ", "").replace("F", "").strip()
                        recettes_precedentes = float(valeur_nettoyee) if valeur_nettoyee else 0.0
                    else:
                        recettes_precedentes = 0.0

                    nouveau_total_mois = recettes_precedentes + total_recettes_mois

                    # On enregistre le montant cumulé final
                    feuille_principale.update_cell(4, colonne_mois, nouveau_total_mois)

                
                    st.success("📊 Bilan mensuel automatiquement consolidé et envoyé sur le fichier principal !")
                    
                    # Réinitialisation du fichier secondaire pour le mois suivant
                    feuille.delete_rows(2, len(toutes_les_lignes))

            except Exception as e:
            # Fermeture propre du bloc try principal
              st.error(f"🚨 Une erreur est survenue lors du traitement : {e}")               