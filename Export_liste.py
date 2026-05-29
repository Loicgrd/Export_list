import streamlit as st
import pandas as pd
import io
import requests
from streamlit_gsheets import GSheetsConnection

# Configuration
st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

# URL corrigée (sans ?usp=sharing)
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit"
conn = st.connection("gsheets", type=GSheetsConnection)

# Fonction de lecture sécurisée
def load_bailleurs():
    try:
        # worksheet=0 cible le premier onglet, ttl=0 pour forcer la lecture fraîche
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=0, ttl=0)
        if df.empty or len(df.columns) < 2:
            return {}, pd.DataFrame(columns=["Nom du Bailleur", "SIREN"])
        df = df.dropna(subset=[df.columns[0]])
        # Retourne le dictionnaire et le dataframe nettoyé
        return dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str))), df
    except Exception as e:
        st.error(f"Erreur de lecture Google Sheet : {e}")
        return {}, pd.DataFrame(columns=["Nom du Bailleur", "SIREN"])

dict_siren, df_bailleurs_gsheet = load_bailleurs()

# ==========================================
# INTERFACE
# ==========================================
st.title("Générateur d'Exports CEE")
tab_generateur, tab_reglages = st.tabs(["📊 Générateur d'Exports", "⚙️ Gestion des Bailleurs"])

# --- ONGLET 1 : RÉGLAGES (Avec affichage "Poubelle" par ligne) ---
with tab_reglages:
    st.subheader("📋 Bailleurs enregistrés")
    if not df_bailleurs_gsheet.empty:
        for index, row in df_bailleurs_gsheet.iterrows():
            c1, c2, c3 = st.columns([3, 2, 1])
            c1.write(f"🏢 {row.iloc[0]}")
            c2.write(f"🆔 {row.iloc[1]}")
            if c3.button("❌", key=f"del_{index}"):
                df_updated = df_bailleurs_gsheet.drop(index)
                conn.update(spreadsheet=SPREADSHEET_URL, worksheet=0, data=df_updated)
                st.rerun()
    
    st.divider()
    st.subheader("➕ Ajouter des bailleurs (par SIREN)")
    liste_sirens_brut = st.text_area("Collez les SIREN (séparés par des virgules) :")
    if st.button("Rechercher et Ajouter"):
        sirens = [s.strip() for s in liste_sirens_brut.replace('\n', ',').split(',') if s.strip()]
        nouveaux = []
        for s in sirens:
            resp = requests.get(f"https://recherche-entreprises.api.gouv.fr/search?q={s}")
            if resp.status_code == 200 and resp.json().get("results"):
                res = resp.json()["results"][0]
                nouveaux.append({df_bailleurs_gsheet.columns[0]: res.get('sigle') or res.get('nom_raison_sociale'),
                                 df_bailleurs_gsheet.columns[1]: res.get('siren')})
        if nouveaux:
            df_updated = pd.concat([df_bailleurs_gsheet, pd.DataFrame(nouveaux)], ignore_index=True)
            conn.update(spreadsheet=SPREADSHEET_URL, worksheet=0, data=df_updated)
            st.rerun()

# --- ONGLET 2 : GÉNÉRATEUR ---
with tab_generateur:
    uploaded_file = st.file_uploader("Importer le fichier Excel source (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df_source = pd.read_excel(uploaded_file)
        
        # Nettoyage dates
        for col in ['Date réception', "Date d'engagement", 'Date prévisionnelle de réalisation', 'Date réelle de réalisation']:
            if col in df_source.columns:
                df_source[col] = pd.to_datetime(df_source[col], errors='coerce').dt.date

        # Traitement Confort & Export
        bailleurs_cibles = list(dict_siren.keys())
        df_confort = df_source[df_source['Bénéficiaire'].isin(bailleurs_cibles)].copy()
        
        if not df_confort.empty:
            df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_siren))
            df_confort.insert(1, 'BS Confort', df_confort['Bénéficiaire'])
            
        df_export = df_source[df_source['Contrôle'] != 'Non concerné'].copy()
        if not df_confort.empty:
            df_export = df_export[~df_export['Numéro dossier'].isin(df_confort['Numéro dossier'].dropna().unique())]

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("📊 Liste à exporter")
            buffer_exp = io.BytesIO()
            with pd.ExcelWriter(buffer_exp, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                df_export.to_excel(w, index=False)
            st.download_button("📥 Télécharger Liste", buffer_exp.getvalue(), "Liste_a_exporter.xlsx")
            
        with c2:
            st.subheader("🏢 Fichier Confort")
            buffer_con = io.BytesIO()
            with pd.ExcelWriter(buffer_con, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as w:
                df_confort.to_excel(w, index=False)
            st.download_button("📥 Télécharger Confort", buffer_con.getvalue(), "Fichier_Confort.xlsx")
