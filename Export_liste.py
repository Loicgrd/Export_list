import streamlit as st
import pandas as pd
import io
from datetime import datetime
import pytz
import re
from streamlit_gsheets import GSheetsConnection

from fonctions import (
    afficher_gestion_base, 
    sauvegarder_admin, 
    afficher_tableau_synthese, 
    ajouter_feuille_formatee, 
    ajouter_feuille_dcr,       
    afficher_gestion_liste_bs
)

st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

TZ_FRANCE = pytz.timezone('Europe/Paris')
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1pkj6frncXmzUUVAClAWp63HY_UpQUahOCLz19w-UseI/edit?usp=sharing"
conn = st.connection("gsheets", type=GSheetsConnection)

@st.cache_data(ttl=10)
def load_worksheet(sheet_name):
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name)
        if df.empty or len(df.columns) < 3: df = pd.DataFrame(columns=["Nom", "SIREN", "Date d'ajout"])
        if len(df.columns) < 3: df["Date d'ajout"] = ""
        df = df.dropna(subset=[df.columns[0]])
        df = df.iloc[:, :3]
        return dict(zip(df.iloc[:, 0], df.iloc[:, 1].astype(str))), df
    except Exception as e:
        return {}, pd.DataFrame(columns=["Nom", "SIREN", "Date d'ajout"])

@st.cache_data(ttl=10)
def load_admin():
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="ADMIN")
        if df.empty: return [], []
        mots_docs = [str(x).strip() for x in df.get("Nom du document", pd.Series()).dropna() if str(x).strip() and str(x).strip().lower() != 'nan']
        mots_coms = [str(x).strip() for x in df.get("Commentaire", pd.Series()).dropna() if str(x).strip() and str(x).strip().lower() != 'nan']
        return mots_docs, mots_coms
    except Exception as e:
        return [], []

@st.cache_data(ttl=30)
def load_liste_bs():
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="Liste_BS")
        if df.empty: return {}, pd.DataFrame(columns=["SIREN", "Nom BS"])
        df_clean = df.copy()
        df_clean['SIREN'] = df_clean['SIREN'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True)
        df_clean['Nom BS'] = df_clean['Nom BS'].astype(str).str.strip()
        return dict(zip(df_clean['SIREN'], df_clean['Nom BS'])), df
    except Exception as e:
        return {}, pd.DataFrame(columns=["SIREN", "Nom BS"])

# --- Initialisation ---
dict_confort, df_confort_gsheet = load_worksheet("Confort")
dict_national, df_national_gsheet = load_worksheet("CDC") 
mots_docs_admin, mots_coms_admin = load_admin()
dict_bs_global, df_liste_bs_gsheet = load_liste_bs()

st.title("Générateur d'Exports CEE")

tab_generateur, tab_confort, tab_national, tab_admin, tab_liste_bs = st.tabs([
    "📊 Générateur", "⚙️ Base Confort", "⚙️ Base National", "🛡️ Filtres ADMIN", "🏢 Base BS ODICEE"
])

with tab_confort: afficher_gestion_base("Confort", df_confort_gsheet, conn, SPREADSHEET_URL, dict_bs_global)
with tab_national: afficher_gestion_base("CDC", df_national_gsheet, conn, SPREADSHEET_URL, dict_bs_global)
with tab_liste_bs: afficher_gestion_liste_bs(df_liste_bs_gsheet, conn, SPREADSHEET_URL)

with tab_admin:
    st.header("🛠️ Mots-clés pour le tri automatique ADMIN")
    col_d, col_c = st.columns(2)
    with col_d:
        st.subheader("📄 Colonne 'Nom du document'")
        nouveau_doc = st.text_input("Ajouter un mot-clé (ex: Visa) :")
        if st.button("➕ Ajouter", key="add_doc") and nouveau_doc:
            if nouveau_doc not in mots_docs_admin: sauvegarder_admin(mots_docs_admin + [nouveau_doc], mots_coms_admin, conn, SPREADSHEET_URL)
        if mots_docs_admin:
            a_suppr_doc = st.multiselect("Supprimer :", mots_docs_admin, key="suppr_doc")
            if st.button("🗑️ Enlever", key="btn_suppr_doc") and a_suppr_doc:
                sauvegarder_admin([m for m in mots_docs_admin if m not in a_suppr_doc], mots_coms_admin, conn, SPREADSHEET_URL)
            st.dataframe(pd.DataFrame(mots_docs_admin, columns=["Mots-clés (Documents)"]), hide_index=True, use_container_width=True)
    with col_c:
        st.subheader("💬 Colonne 'Commentaire'")
        nouveau_com = st.text_input("Ajouter un mot-clé (ex: Abandon) :")
        if st.button("➕ Ajouter", key="add_com") and nouveau_com:
            if nouveau_com not in mots_coms_admin: sauvegarder_admin(mots_docs_admin, mots_coms_admin + [nouveau_com], conn, SPREADSHEET_URL)
        if mots_coms_admin:
            a_suppr_com = st.multiselect("Supprimer :", mots_coms_admin, key="suppr_com")
            if st.button("🗑️ Enlever", key="btn_suppr_com") and a_suppr_com:
                sauvegarder_admin(mots_docs_admin, [m for m in mots_coms_admin if m not in a_suppr_com], conn, SPREADSHEET_URL)
            st.dataframe(pd.DataFrame(mots_coms_admin, columns=["Mots-clés (Commentaires)"]), hide_index=True, use_container_width=True)

# ==========================================
# ONGLET PRINCIPAL : GÉNÉRATEUR
# ==========================================
with tab_generateur:
    uploaded_file = st.file_uploader("Importer le fichier Excel (Liste globale)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df_source = pd.read_excel(uploaded_file)
            st.success(f"Fichier chargé ! ({len(df_source)} lignes)")

            for col in df_source.columns:
                if 'date' in str(col).lower(): df_source[col] = pd.to_datetime(df_source[col], errors='coerce')

            st.markdown("### 📊 Synthèse Globale de l'import (Avant filtres)")
            afficher_tableau_synthese(df_source, "Données brutes")
            st.divider()

            st.subheader("📅 Analyse et définition de la priorité (DCR)")
            date_prio = st.date_input("Dossiers reçus JUSQU'À cette date (incluse) = Prioritaires par défaut :", value=None, format="DD/MM/YYYY")

            df_confort, df_national = pd.DataFrame(), pd.DataFrame()
            if 'Bénéficiaire' in df_source.columns:
                df_confort = df_source[df_source['Bénéficiaire'].isin(dict_confort.keys())].copy()
                df_national = df_source[df_source['Bénéficiaire'].isin(dict_national.keys())].copy()

            df_export = df_source.copy()
            if 'Numéro dossier' in df_export.columns:
                if not df_confort.empty: df_export = df_export[~df_export['Numéro dossier'].isin(df_confort['Numéro dossier'])]
                if not df_national.empty: df_export = df_export[~df_export['Numéro dossier'].isin(df_national['Numéro dossier'])]

            df_admin = pd.DataFrame()
            df_admin_national = pd.DataFrame()
            df_admin_dcr = pd.DataFrame()
            mask_admin = pd.Series(False, index=df_export.index)
            
            if 'Nom du document' in df_export.columns and mots_docs_admin:
                for mot in mots_docs_admin: mask_admin = mask_admin | df_export['Nom du document'].astype(str).str.contains(mot, case=False, na=False, regex=False)
            if 'Commentaire' in df_export.columns and mots_coms_admin:
                for mot in mots_coms_admin: mask_admin = mask_admin | df_export['Commentaire'].astype(str).str.contains(mot, case=False, na=False, regex=False)
            
            if mask_admin.any():
                df_admin = df_export[mask_admin].copy()
                df_export = df_export[~mask_admin] 
                
                if 'Bénéficiaire' in df_admin.columns:
                    mask_admin_nat = df_admin['Bénéficiaire'].isin(dict_national.keys())
                    df_admin_national = df_admin[mask_admin_nat].copy()
                    df_admin_dcr = df_admin[~mask_admin_nat].copy()
                else:
                    df_admin_dcr = df_admin.copy()

            def trier_df(df):
                cols_tri = []
                if 'Date réception' in df.columns: cols_tri.append('Date réception')
                if 'Numéro dossier' in df.columns: cols_tri.append('Numéro dossier')
                if cols_tri: return df.sort_values(by=cols_tri, na_position='last')
                return df
                
            df_export = trier_df(df_export)
            df_confort = trier_df(df_confort)
            df_national = trier_df(df_national)
            df_admin_national = trier_df(df_admin_national)
            df_admin_dcr = trier_df(df_admin_dcr)

            # ---------------------------------------------------------
            # GESTION INTERACTIVE DE LA LISTE PRIORITAIRE DCR
            # ---------------------------------------------------------
            mask_prio = pd.Series(False, index=df_export.index)
            if 'Date réception' in df_export.columns and date_prio is not None:
                dates_reception = pd.to_datetime(df_export['Date réception']).dt.date
                mask_prio = dates_reception <= date_prio
                
            # Ajout de la colonne Prioritaire tout au début
            df_export.insert(0, 'Prioritaire', mask_prio)

            # Configuration des colonnes
            config_colonnes = {
                "Prioritaire": st.column_config.CheckboxColumn(
                    "⭐ Prioritaire",
                    help="Cochez pour placer ce dossier dans la liste prioritaire",
                    default=False
                ),
                "Bénéficiaire": st.column_config.TextColumn(
                    "Bénéficiaire",
                    width="small"
                )
            }
            
            # Formatage dynamique : on applique les largeurs réduites sur Commentaire et certaines Dates
            for col in df_export.columns:
                col_name_lower = str(col).lower()
                
                # Réduction du "commentaire"
                if 'commentaire' in col_name_lower:
                    config_colonnes[col] = st.column_config.TextColumn(
                        col,
                        width="small"
                    )
                
                # Formatage JJ/MM/AAAA pour toutes les "dates"
                elif 'date' in col_name_lower:
                    # On réduit la largeur si c'est une date "prévisionnelle" ou "réelle"
                    if 'prévisionnelle' in col_name_lower or 'réelle' in col_name_lower:
                        config_colonnes[col] = st.column_config.DateColumn(
                            col,
                            format="DD/MM/YYYY",
                            width="small"
                        )
                    else:
                        config_colonnes[col] = st.column_config.DateColumn(
                            col,
                            format="DD/MM/YYYY"
                        )

            with st.expander("🛠️ Afficher les dossiers de la file DCR et ajuster la liste prioritaire", expanded=True):
                st.info("💡 **Astuce :** Vous pouvez cocher/décocher manuellement n'importe quelle ligne.\n\n📌 **Pour figer la colonne 'Prioritaire' pendant que vous allez vers la droite :** Passez votre souris sur l'en-tête '⭐ Prioritaire', cliquez sur les **3 petits points (⋮)** qui apparaissent, puis choisissez **'Pin column'**.")
                
                # Le tableau de base propre
                df_export_modifie = st.data_editor(
                    df_export,
                    column_config=config_colonnes,
                    disabled=[col for col in df_export.columns if col != "Prioritaire"],
                    hide_index=True,
                    use_container_width=True,
                    height=450,
                    key=f"editor_dcr_{date_prio}" 
                )

            df_prio = df_export_modifie[df_export_modifie['Prioritaire']].drop(columns=['Prioritaire']).copy()
            df_classique = df_export_modifie[~df_export_modifie['Prioritaire']].drop(columns=['Prioritaire']).copy()
            # ---------------------------------------------------------

            st.divider()
            
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                ajouter_feuille_dcr(writer, df_prio, df_classique, 'DCR')
                if not df_national.empty: ajouter_feuille_formatee(writer, df_national.drop(columns=['SIREN'], errors='ignore'), 'National')
                if not df_admin_national.empty: ajouter_feuille_formatee(writer, df_admin_national.drop(columns=['SIREN'], errors='ignore'), 'ADMIN National')
                if not df_admin_dcr.empty: ajouter_feuille_formatee(writer, df_admin_dcr.drop(columns=['SIREN'], errors='ignore'), 'ADMIN DCR')
                if not df_confort.empty: ajouter_feuille_formatee(writer, df_confort.drop(columns=['SIREN'], errors='ignore'), 'Confort')

            date_export = datetime.now(TZ_FRANCE).strftime("%d-%m-%Y")
            st.download_button(
                label="📥 TÉLÉCHARGER L'EXPORT GLOBAL ODICEE (.xlsx)", 
                data=excel_buffer.getvalue(), 
                file_name=f"ODICEE-{date_export}-Global.xlsx", 
                use_container_width=True, 
                type="primary"
            )
            
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("📈 DCR", f"{len(df_prio) + len(df_classique)} lignes")
            c2.metric("🇫🇷 National", f"{len(df_national)} lignes")
            c3.metric("🛡️ ADMIN National", f"{len(df_admin_national)} lignes")
            c4.metric("🛡️ ADMIN DCR", f"{len(df_admin_dcr)} lignes")
            c5.metric("🏢 Confort", f"{len(df_confort)} lignes")

            st.divider()
            
            st.markdown("### 📊 Tableaux de synthèse par onglet")
            
            tog1, tog2, tog3, tog4, tog5 = st.columns(5)
            show_dcr = tog1.toggle("📈 Synthèse DCR", value=False)
            show_nat = tog2.toggle("🇫🇷 Synthèse National", value=False)
            show_admin_nat = tog3.toggle("🛡️ Synthèse ADM Nat", value=False)
            show_admin_dcr = tog4.toggle("🛡️ Synthèse ADM DCR", value=False)
            show_confort = tog5.toggle("🏢 Synthèse CONFORT", value=False)
            
            if show_dcr:
                df_dcr_complet = pd.concat([df_prio, df_classique]) if not df_prio.empty or not df_classique.empty else pd.DataFrame()
                afficher_tableau_synthese(df_dcr_complet, "📈 Synthèse DCR")
            if show_nat: afficher_tableau_synthese(df_national, "🇫🇷 Synthèse National")
            if show_admin_nat: afficher_tableau_synthese(df_admin_national, "🛡️ Synthèse ADMIN National")
            if show_admin_dcr: afficher_tableau_synthese(df_admin_dcr, "🛡️ Synthèse ADMIN DCR")
            if show_confort: afficher_tableau_synthese(df_confort, "🏢 Synthèse CONFORT")

        except Exception as e:
            st.error(f"Erreur lors de la génération de l'Excel : {e}")