import streamlit as st
import pandas as pd
import io

# Configuration de la page
st.set_page_config(page_title="Générateur d'Exports CEE", layout="wide")

st.title("Générateur d'Exports : Liste à Exporter & Confort")
st.markdown("Chargez votre fichier contenant la **Liste globale** pour générer les fichiers d'export.")

# 1. Upload du fichier
uploaded_file = st.file_uploader("Importer le fichier Excel source (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Lecture de la feuille principale
        df_source = pd.read_excel(uploaded_file)
        st.success(f"Fichier chargé avec succès ! ({len(df_source)} lignes trouvées)")

        # ==========================================
        # NETTOYAGE ET FORMATAGE DES DATES
        # ==========================================
        colonnes_dates = ['Date réception', "Date d'engagement", 'Date prévisionnelle de réalisation', 'Date réelle de réalisation']
        for col in colonnes_dates:
            if col in df_source.columns:
                # Conversion en datetime puis extraction de la date pure (supprime les heures)
                df_source[col] = pd.to_datetime(df_source[col], errors='coerce').dt.date

        # ==========================================
        # TRAITEMENT 1 : CONFORT
        # ==========================================
        bailleurs_cibles = [
            "INOLYA", 
            "IMMOBILIERE RHONE ALPES SA D'HLM", 
            "OPH TROYES AUBE HABITAT", 
            "SEINE-SAINT-DENIS HABITAT"
        ]
        
        dict_siren = {
            "INOLYA": "780705703",
            "IMMOBILIERE RHONE ALPES SA D'HLM": "661750067",
            "OPH TROYES AUBE HABITAT": "902718998",
            "SEINE-SAINT-DENIS HABITAT": "279300198"
        }

        df_confort = pd.DataFrame()
        if 'Bénéficiaire' in df_source.columns:
            # Filtrer pour ne garder que les bailleurs cibles
            df_confort = df_source[df_source['Bénéficiaire'].isin(bailleurs_cibles)].copy()
            
            # Création des colonnes spécifiques
            df_confort.insert(0, 'SIREN', df_confort['Bénéficiaire'].map(dict_siren))
            df_confort.insert(1, 'BS Confort', df_confort['Bénéficiaire'])
            
            # Sélection et ordre des colonnes
            colonnes_attendues = ['SIREN', 'BS Confort', 'Date réception', 'Numéro dossier', 'Bénéficiaire', 'Stade']
            colonnes_dispos = [c for c in colonnes_attendues if c in df_confort.columns]
            df_confort = df_confort[colonnes_dispos]

        # ==========================================
        # TRAITEMENT 2 : LISTE À EXPORTER (SANS DOUBLONS)
        # ==========================================
        df_export = df_source.copy()
        
        # Filtre sur le statut de contrôle si nécessaire
        if 'Contrôle' in df_export.columns:
            df_export = df_export[df_export['Contrôle'] != 'Non concerné']
        
        # EXCLUSION DES LIGNES CONFORT : On supprime de l'export les dossiers présents dans Confort
        if not df_confort.empty and 'Numéro dossier' in df_export.columns and 'Numéro dossier' in df_confort.columns:
            liste_dossiers_confort = df_confort['Numéro dossier'].dropna().unique()
            df_export = df_export[~df_export['Numéro dossier'].isin(liste_dossiers_confort)]

        # ==========================================
        # INTERFACE DE TÉLÉCHARGEMENT
        # ==========================================
        st.divider()
        col1, col2 = st.columns(2)
        
        # --- Bloc Gauche : Liste à exporter ---
        with col1:
            st.subheader("📊 Liste à exporter")
            st.text(f"{len(df_export)} lignes (Excluant les dossiers envoyés dans Confort).")
            st.dataframe(df_export.head(3))
            
            # Génération du fichier Excel avec forçage du format date natif sans heure
            buffer_export = io.BytesIO()
            with pd.ExcelWriter(buffer_export, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Liste à exporter')
            
            st.download_button(
                label="📥 Télécharger la Liste à exporter",
                data=buffer_export.getvalue(),
                file_name="Liste_a_exporter.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        # --- Bloc Droit : Confort ---
        with col2:
            st.subheader("🏢 Confort (Bailleurs ciblés)")
            st.text(f"{len(df_confort)} lignes identifiées.")
            st.dataframe(df_confort.head(3))
            
            buffer_confort = io.BytesIO()
            with pd.ExcelWriter(buffer_confort, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
                df_confort.to_excel(writer, index=False, sheet_name='Confort')
                
            st.download_button(
                label="📥 Télécharger le fichier Confort",
                data=buffer_confort.getvalue(),
                file_name="Fichier_Confort.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Une erreur est survenue lors de l'analyse du fichier : {e}")
