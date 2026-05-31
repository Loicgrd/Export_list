import streamlit as st
import pandas as pd
import io
import requests
from datetime import datetime # NOUVEAU : Pour gérer la date du jour
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

# ==========================================
# CONNEXION AU GOOGLE SHEET
# ==========================================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit?usp=sharing"

conn = st.connection("gsheets", type=GSheetsConnection)

@st.cache_data(ttl=10)
def load_bailleurs():
    try:
        # On lit toutes les colonnes disponibles
        df = conn.read(spreadsheet=SPREADSHEET_URL)
        
        # S'il n'y a pas encore de 3ème colonne dans le GSheet, on la crée virtuellement
        if len(df.columns) < 3:
            df["Date d'ajout"] = ""
            
        # On nettoie les lignes vides (basé sur le nom)
        df = df.dropna(subset=[df.columns[0]])
        
        # On sécurise en ne gardant que les 3 premières colonnes (Nom, SIREN, Date)
        df = df.iloc[:, :3]
        
        # Le dictionnaire pour le traitement { "NOM": "SIREN" }
        dict_siren = dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str)))
        
        return dict_siren, df
    except Exception as e:
        st.error(f"Erreur de lecture du Google Sheet : {e}")
        return {}, pd.DataFrame()

dict_siren, df_bailleurs_gsheet = load_bailleurs()

st.title("Générateur d'Exports CEE")

tab_generateur, tab_reglages = st.tabs(["📊 Générateur d'Exports", "⚙️ Base de données Bailleurs"])

# ==========================================
# ONGLET 2 : GESTION DANS GOOGLE SHEETS
# ==========================================

with tab_reglages:
    st.header("Base de données 'Confort'")

    # --- INITIALISATION DE LA MÉMOIRE (SESSION STATE) ---
    if "recherche_en_cours" not in st.session_state:
        st.session_state.recherche_en_cours = False
        st.session_state.bailleurs_trouves = []

    # --- SECTION MULTI-AJOUT ---
    st.subheader("➕ Ajouter plusieurs bailleurs (par SIREN)")
    liste_sirens_brut = st.text_area("Collez vos SIREN ici (séparés par des virgules ou retours à la ligne) :")
    
    # ÉTAPE 1 : LA RECHERCHE
    if st.button("🔍 Rechercher les SIREN"):
        sirens = [s.strip() for s in liste_sirens_brut.replace('\n', ',').split(',') if s.strip()]
        
        trouves = []
        with st.spinner("Recherche dans la base du Gouvernement..."):
            for s in sirens:
                resp = requests.get(f"https://recherche-entreprises.api.gouv.fr/search?q={s}")
                if resp.status_code == 200 and resp.json().get("results"):
                    res = resp.json()["results"][0]
                    trouves.append({
                        "Nom": res.get('sigle') or res.get('nom_raison_sociale'),
                        "SIREN": res.get('siren')
                    })
        
        # On stocke les résultats et on affiche l'étape 2
        st.session_state.bailleurs_trouves = trouves
        st.session_state.recherche_en_cours = True

    # ÉTAPE 2 : LA CONFIRMATION
    if st.session_state.recherche_en_cours:
        if st.session_state.bailleurs_trouves:
            st.success(f"✅ {len(st.session_state.bailleurs_trouves)} bailleur(s) trouvé(s) ! Veuillez vérifier avant d'ajouter :")
            
            # Affichage de prévisualisation
            st.dataframe(pd.DataFrame(st.session_state.bailleurs_trouves), hide_index=True)
            
            col_btn1, col_btn2 = st.columns([0.2, 0.8])
            with col_btn1:
                if st.button("✅ Confirmer l'ajout", type="primary"):
                    nom_col = df_bailleurs_gsheet.columns[0]
                    siren_col = df_bailleurs_gsheet.columns[1]
                    date_col = df_bailleurs_gsheet.columns[2]
                    date_jour = datetime.now().strftime("%d/%m/%Y") # Date d'aujourd'hui
                    
                    nouveaux_bailleurs = []
                    for b in st.session_state.bailleurs_trouves:
                        nouveaux_bailleurs.append({
                            nom_col: b['Nom'],
                            siren_col: b['SIREN'],
                            date_col: date_jour
                        })
                    
                    # Mise à jour GSheet
                    df_new = pd.DataFrame(nouveaux_bailleurs)
                    df_updated = pd.concat([df_bailleurs_gsheet, df_new], ignore_index=True)
                    conn.update(spreadsheet=SPREADSHEET_URL, worksheet="Confort", data=df_updated)
                    
                    # Reset et rafraîchissement
                    st.session_state.recherche_en_cours = False
                    st.cache_data.clear()
                    st.rerun()
                    
            with col_btn2:
                if st.button("❌ Annuler"):
                    st.session_state.recherche_en_cours = False
                    st.rerun()
        else:
            st.warning("❌ Aucun bailleur trouvé pour ces SIREN.")
            if st.button("Nouvelle recherche"):
                st.session_state.recherche_en_cours = False
                st.rerun()

    st.divider()

    # --- LISTE ACTUELLE ---
    st.subheader("📋 Liste actuelle")
    
    if not df_bailleurs_gsheet.empty:
        # On prépare l'affichage pour mettre la date en premier
        df_display = df_bailleurs_gsheet.copy()
        col_nom = df_display.columns[0]
        col_siren = df_display.columns[1]
        col_date = df_display.columns[2]
        
        # On réorganise l'ordre des colonnes : Date -> Nom -> SIREN
        df_display = df_display[[col_date, col_nom, col_siren]]
        
        # NOUVEAU : hide_index=True supprime les numéros 0, 1, 2... à gauche !
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    else:
        st.info("La liste est actuellement vide.")

    # --- SECTION SUPPRESSION ---
    st.subheader("🗑️ Supprimer des bailleurs")
    if not df_bailleurs_gsheet.empty:
        bailleurs_a_supprimer = st.multiselect("Sélectionner les bailleurs à supprimer :", options=df_bailleurs_gsheet.iloc[:, 0].tolist())
        
        if st.button("Supprimer la sélection", type="primary"):
            df_updated = df_bailleurs_gsheet[~df_bailleurs_gsheet.iloc[:, 0].isin(bailleurs_a_supprimer)]
            conn.update(spreadsheet=SPREADSHEET_URL, worksheet="Confort", data=df_updated)
            st.cache_data.clear()
            st.rerun()

    st.divider()


# ==========================================
# ONGLET 1 : GÉNÉRATEUR D'EXCEL
# ==========================================
with tab_generateur:
    uploaded_file = st.file_uploader("Importer le fichier Excel (Liste globale)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df_source = pd.read_excel(uploaded_file)
            st.success(f"Fichier chargé ! ({len(df_source)} lignes)")

            for col in ['Date réception', "Date d'engagement", 'Date prévisionnelle de réalisation', 'Date réelle de réalisation']:
                if col in df_source.columns:
                    df_source[col] = pd.to_datetime(df_source[col], errors='coerce').dt.date

            bailleurs_cibles = list(dict_siren.keys())
            df_confort = pd.DataFrame()
            
            if 'Bénéficiaire' in df_source.columns:
                df_confort = df_source[df_source['Bénéficiaire'].isin(bailleurs_cibles)].copy()
                if not df_confort.empty:
                    df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_siren))
                    df_confort.insert(1, 'BS Confort', df_confort['Bénéficiaire'])
                    
                    cols_attendues = ['SIREN', 'BS Confort', 'Date réception', 'Numéro dossier', 'Bénéficiaire', 'Stade']
                    df_confort = df_confort[[c for c in cols_attendues if c in df_confort.columns]]

            df_export = df_source.copy()
            if 'Contrôle' in df_export.columns:
                df_export = df_export[df_export['Contrôle'] != 'Non concerné']
            
            if not df_confort.empty and 'Numéro dossier' in df_export.columns:
                dossiers_confort = df_confort['Numéro dossier'].dropna().unique()
                df_export = df_export[~df_export['Numéro dossier'].isin(dossiers_confort)]

            st.divider()
            c1, c2 = st.columns(2)
            
            with c1:
                st.subheader("📊 Liste à exporter")
                st.text(f"{len(df_export)} lignes.")
                buffer_export = io.BytesIO()
                with pd.ExcelWriter(buffer_export, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                    df_export.to_excel(w, index=False, sheet_name='Liste à exporter')
                st.download_button("📥 Télécharger", buffer_export.getvalue(), "Liste_a_exporter.xlsx", use_container_width=True)
                
            with c2:
                st.subheader("🏢 Fichier Confort")
                st.text(f"{len(df_confort)} lignes.")
                if not df_confort.empty:
                    buffer_confort = io.BytesIO()
                    with pd.ExcelWriter(buffer_confort, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                        df_confort.to_excel(w, index=False, sheet_name='Confort')
                    st.download_button("📥 Télécharger", buffer_confort.getvalue(), "Fichier_Confort.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"Erreur : {e}")
