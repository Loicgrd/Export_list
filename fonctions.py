import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
import pytz
from xlsxwriter.utility import xl_col_to_name
import re
from jours_feries_france import JoursFeries

TZ_FRANCE = pytz.timezone('Europe/Paris')


def get_n_derniers_jours_ouvres(date_reference, n=5):
    """Retourne les n derniers jours ouvrés (hors week-ends et jours fériés FR)
    se terminant à date_reference (incluse si elle est ouvrée)."""
    jours_ouvres = []
    jour_courant = date_reference

    # Jours fériés sur l'année en cours + précédente (au cas où on est début janvier)
    annees = {date_reference.year, date_reference.year - 1}
    feries = set()
    for annee in annees:
        feries.update(JoursFeries.for_year(annee).values())

    while len(jours_ouvres) < n:
        # lundi=0 ... dimanche=6 -> on exclut samedi(5)/dimanche(6) et jours fériés
        if jour_courant.weekday() < 5 and jour_courant not in feries:
            jours_ouvres.append(jour_courant)
        jour_courant -= timedelta(days=1)

    return jours_ouvres  # liste de dates, du plus récent au plus ancien

# ==========================================
# FONCTIONS DE GESTION DES BASES (Google Sheets)
# ==========================================
def afficher_gestion_base(sheet_name, df_gsheet, conn, spreadsheet_url, dict_bs):
    st.header(f"Base de données '{sheet_name}'")
    state_recherche = f"recherche_{sheet_name}"
    state_trouves = f"trouves_{sheet_name}"
    
    if state_recherche not in st.session_state:
        st.session_state[state_recherche] = False
        st.session_state[state_trouves] = []
        
    def nettoyer_siren(val):
        val_str = str(val).strip()
        val_str = re.sub(r'\.0$', '', val_str)
        return re.sub(r'\D', '', val_str)

    sirens_existants = []
    if not df_gsheet.empty and len(df_gsheet.columns) > 1:
        sirens_existants = [nettoyer_siren(x) for x in df_gsheet.iloc[:, 1].tolist()]

    st.subheader("➕ Ajouter plusieurs bailleurs (par SIREN)")
    liste_sirens_brut = st.text_area(f"Collez vos SIREN ici :", key=f"input_siren_{sheet_name}")
    
    if st.button("🔍 Rechercher", key=f"btn_search_{sheet_name}"):
        sirens_bruts_split = liste_sirens_brut.replace('\n', ',').split(',')
        sirens_propres = [nettoyer_siren(s) for s in sirens_bruts_split if nettoyer_siren(s)]
        sirens = list(dict.fromkeys(sirens_propres)) 
        
        trouves = []
        non_trouves = []
        deja_presents = []
        
        with st.spinner("Recherche et vérification des doublons..."):
            for s in sirens:
                if s in sirens_existants: deja_presents.append(s)
                elif s in dict_bs: trouves.append({"Nom": dict_bs[s], "SIREN": s})
                else: non_trouves.append(s)
                    
        st.session_state[state_trouves] = trouves
        st.session_state[state_recherche] = True
        
        if deja_presents: st.info(f"ℹ️ Ces SIREN sont **déjà présents** dans la base '{sheet_name}' et ont été ignorés : {', '.join(deja_presents)}")
        if non_trouves: st.warning(f"⚠️ Ces SIREN n'ont **pas été trouvés** dans la base Liste_BS : {', '.join(non_trouves)}")

    if st.session_state[state_recherche]:
        if st.session_state[state_trouves]:
            st.success(f"✅ {len(st.session_state[state_trouves])} nouveaux Bailleurs Sociaux prêts à être ajoutés !")
            st.dataframe(pd.DataFrame(st.session_state[state_trouves]), hide_index=True)
            
            if st.button("✅ Confirmer l'ajout", key=f"btn_conf_{sheet_name}", type="primary"):
                date_jour = datetime.now(TZ_FRANCE).strftime("%d/%m/%Y")
                nom_col, siren_col, date_col = df_gsheet.columns[0], df_gsheet.columns[1], df_gsheet.columns[2]
                nouveaux = [{nom_col: b['Nom'], siren_col: b['SIREN'], date_col: date_jour} for b in st.session_state[state_trouves]]
                df_updated = pd.concat([df_gsheet, pd.DataFrame(nouveaux)], ignore_index=True)
                conn.update(spreadsheet=spreadsheet_url, worksheet=sheet_name, data=df_updated)
                
                st.session_state[state_trouves] = []
                st.session_state[state_recherche] = False
                st.cache_data.clear() 
                st.rerun()
                
        elif not st.session_state[state_trouves] and liste_sirens_brut.strip():
            st.error("Aucun nouveau bailleur à ajouter à la base.")
    
    st.divider()
    st.subheader("📋 Liste actuelle")
    if not df_gsheet.empty:
        df_display = df_gsheet[[df_gsheet.columns[2], df_gsheet.columns[0], df_gsheet.columns[1]]]
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        st.subheader("🗑️ Supprimer")
        a_supprimer = st.multiselect("Sélectionner les noms à supprimer :", options=df_gsheet.iloc[:, 0].tolist(), key=f"del_select_{sheet_name}")
        
        if st.button("🗑️ Supprimer la sélection", type="primary", key=f"btn_del_{sheet_name}"):
            if a_supprimer:
                df_updated = df_gsheet[~df_gsheet.iloc[:, 0].isin(a_supprimer)]
                conn.update(spreadsheet=spreadsheet_url, worksheet=sheet_name, data=df_updated)
                st.cache_data.clear()
                st.rerun()

def sauvegarder_admin(l_docs, l_coms, conn, spreadsheet_url, worksheet="ADMIN"):
    max_len = max(len(l_docs), len(l_coms))
    df_new = pd.DataFrame({
        "Nom du document": l_docs + [""] * (max_len - len(l_docs)),
        "Commentaire": l_coms + [""] * (max_len - len(l_coms))
    })
    conn.update(spreadsheet=spreadsheet_url, worksheet=worksheet, data=df_new)
    st.cache_data.clear()
    st.rerun()

def afficher_gestion_liste_bs(df_gsheet, conn, spreadsheet_url):
    st.markdown("""
    Cette interface vous permet d'interagir directement avec la base de données globale.
    - **➕ Ajouter** : Faites défiler tout en bas du tableau et remplissez la ligne vide.
    - **✏️ Modifier** : Double-cliquez sur n'importe quel texte.
    - **🗑️ Supprimer** : Cochez la case tout à gauche d'une ligne, puis appuyez sur `Suppr`.
    """)
    df_display = df_gsheet.copy()
    if not df_display.empty and 'SIREN' in df_display.columns:
        df_display['SIREN'] = df_display['SIREN'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace('nan', '')

    df_modifie = st.data_editor(df_display, num_rows="dynamic", use_container_width=True, hide_index=True, key="editor_liste_bs")
    st.divider()
    if st.button("💾 Enregistrer la Liste BS", type="primary", use_container_width=True):
        with st.spinner("Sauvegarde en cours..."):
            conn.update(spreadsheet=spreadsheet_url, worksheet="Liste_BS", data=df_modifie)
            st.cache_data.clear()
            st.success("✅ La base de données BS a bien été mise à jour !")
            st.rerun()

def afficher_gestion_initiales(df_gsheet, conn, spreadsheet_url):
    st.markdown("""
    Définissez ici les initiales de votre équipe et choisissez leur couleur en cliquant sur le carré visuel.
    Ces couleurs s'appliqueront automatiquement dans vos exports Excel.
    """)
    
    df_display = df_gsheet.copy()
    if 'Initiales' not in df_display.columns: df_display['Initiales'] = ""
    if 'Couleur' not in df_display.columns: df_display['Couleur'] = ""

    st.subheader("🎨 Ajouter ou modifier une initiale")
    c1, c2, c3 = st.columns([2, 1, 2])
    
    with c1:
        nouvelle_init = st.text_input("Initiales (ex: AB)")
    with c2:
        nouvelle_couleur = st.color_picker("Couleur", "#FFC300")
    with c3:
        st.write("") 
        st.write("")
        if st.button("➕ Enregistrer", use_container_width=True, type="primary"):
            if nouvelle_init.strip():
                df_temp = df_display.copy()
                init_propre = nouvelle_init.strip()
                
                if init_propre in df_temp['Initiales'].values:
                    df_temp.loc[df_temp['Initiales'] == init_propre, 'Couleur'] = nouvelle_couleur
                else:
                    nouvelle_ligne = pd.DataFrame([{"Initiales": init_propre, "Couleur": nouvelle_couleur}])
                    df_temp = pd.concat([df_temp, nouvelle_ligne], ignore_index=True)
                    
                with st.spinner("Sauvegarde en cours..."):
                    conn.update(spreadsheet=spreadsheet_url, worksheet="Liste_initiales", data=df_temp)
                    st.cache_data.clear()
                    st.rerun()
            else:
                st.warning("Veuillez saisir des initiales.")

    st.divider()
    st.subheader("📋 Liste actuelle")
    
    if not df_display.empty and not df_display.dropna(subset=['Initiales']).empty:
        df_clean = df_display.dropna(subset=['Initiales']).copy()
        
        def coloriser_fond(valeur):
            try:
                if str(valeur).startswith('#'):
                    return f'background-color: {valeur}; color: {valeur}; border-radius: 5px;'
            except: pass
            return ''
            
        st.dataframe(
            df_clean.style.map(coloriser_fond, subset=['Couleur']) if hasattr(df_clean.style, 'map') else df_clean.style.applymap(coloriser_fond, subset=['Couleur']),
            use_container_width=True, 
            hide_index=True
        )
        
        st.subheader("🗑️ Supprimer")
        a_supprimer = st.multiselect("Sélectionner les initiales à supprimer :", options=df_clean['Initiales'].tolist(), key="del_init")
        
        if st.button("🗑️ Supprimer la sélection", type="primary", key="btn_del_init"):
            if a_supprimer:
                df_updated = df_display[~df_display['Initiales'].isin(a_supprimer)]
                with st.spinner("Suppression..."):
                    conn.update(spreadsheet=spreadsheet_url, worksheet="Liste_initiales", data=df_updated)
                    st.cache_data.clear()
                    st.rerun()
    else:
        st.info("Aucune initiale n'est paramétrée.")

def afficher_tableau_synthese(df, titre, dict_confort=None, dict_national=None):
    if not df.empty and 'Date réception' in df.columns and 'DCR' in df.columns:
        df_synth = df.dropna(subset=['Date réception']).copy()
        if df_synth.empty: return
        st.markdown(f"**{titre}**")
        df_synth['Date réception'] = pd.to_datetime(df_synth['Date réception'], dayfirst=True).dt.date
        df_synth['DCR'] = df_synth['DCR'].fillna('Non renseigné')
        
        tableau = pd.crosstab(index=df_synth['Date réception'], columns=df_synth['DCR'], margins=True, margins_name='Total')
        tableau_final = pd.concat([tableau.drop('Total').sort_index(ascending=True), tableau.loc[['Total']]])

        # --- Plage "dans les délais" = du 5ème jour OUVRÉ en arrière jusqu'à AUJOURD'HUI ---
        # (bornes incluses, week-ends et jours fériés compris dans la plage : un dossier reçu
        # un samedi/dimanche reste "dans les délais" s'il tombe dans cette fenêtre)
        date_du_jour = datetime.now(TZ_FRANCE).date()
        jours_ouvres = get_n_derniers_jours_ouvres(date_du_jour, n=5)
        borne_min = min(jours_ouvres)  # le 5ème jour ouvré, le plus ancien de la fenêtre
        top_5_dates = {
            (borne_min + timedelta(days=i)).strftime('%d/%m/%Y')
            for i in range((date_du_jour - borne_min).days + 1)
        }

        tableau_final.index = [idx if isinstance(idx, str) else idx.strftime('%d/%m/%Y') for idx in tableau_final.index]

        def coloriser_delais(row):
            styles = []
            for col_name in row.index:
                if row.name == 'Total' or col_name == 'Total': styles.append('text-align: center; background-color: #e6e6e6; font-weight: bold; color: black')
                elif row.name in top_5_dates: styles.append('text-align: center; background-color: #d4edda; color: #155724')
                else: styles.append('text-align: center; background-color: #f8d7da; color: #721c24')
            return styles

        # Le détail cliquable des pièces n'a de sens que lorsqu'on peut déterminer la base
        # d'appartenance (Confort / National / DCR) de chaque bénéficiaire, donc uniquement
        # quand les deux dictionnaires sont fournis (typiquement : la synthèse globale, avant filtres).
        interactif = dict_confort is not None and dict_national is not None

        if not interactif:
            st.dataframe(tableau_final.style.apply(coloriser_delais, axis=1), use_container_width=True)
            return

        event = st.dataframe(
            tableau_final.style.apply(coloriser_delais, axis=1),
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-cell",
            key=f"synthese_{titre}"
        )

        cellules = event.selection.get("cells", []) if event and event.selection else []
        if not cellules:
            st.caption("💡 Cliquez sur une case du tableau (hors ligne/colonne Total) pour afficher le détail des pièces correspondantes.")
            return

        row_idx, col_idx = cellules[0]
        date_selectionnee = tableau_final.index[row_idx]
        dcr_selectionnee = tableau_final.columns[col_idx]

        if date_selectionnee == 'Total' or dcr_selectionnee == 'Total':
            st.info("La ligne/colonne 'Total' n'a pas de détail associé — sélectionnez une case individuelle.")
            return

        mask_detail = (
            df_synth['Date réception'].apply(lambda d: d.strftime('%d/%m/%Y')) == date_selectionnee
        ) & (df_synth['DCR'] == dcr_selectionnee)
        df_detail = df_synth[mask_detail].copy()

        def determiner_base(beneficiaire):
            if beneficiaire in dict_confort: return "Confort"
            if beneficiaire in dict_national: return "National"
            return dcr_selectionnee

        if 'Bénéficiaire' in df_detail.columns:
            df_detail.insert(0, 'Base', df_detail['Bénéficiaire'].map(determiner_base))
        else:
            df_detail.insert(0, 'Base', dcr_selectionnee)

        st.markdown(f"#### 📄 Pièces — {date_selectionnee} / {dcr_selectionnee} ({len(df_detail)} ligne(s))")
        config_detail = {"Base": st.column_config.TextColumn("Base", width="small")}
        for col in df_detail.columns:
            col_name_lower = str(col).lower()
            if 'commentaire' in col_name_lower:
                config_detail[col] = st.column_config.TextColumn(col, width="small")
            elif 'date' in col_name_lower:
                config_detail[col] = st.column_config.DateColumn(col, format="DD/MM/YYYY", width="small")
        st.dataframe(df_detail, column_config=config_detail, hide_index=True, use_container_width=True)
    else:
        st.info(f"**{titre}** : Base vide ou colonnes manquantes.")

# ==========================================
# FONCTIONS EXCEL (HYPER-OPTIMISÉES)
# ==========================================
def ajouter_feuille_formatee(writer, df, nom_feuille, dict_initiales=None):
    if dict_initiales is None: dict_initiales = {}
    if df.empty: return
    if 'Initiales' not in df.columns: df.insert(0, 'Initiales', '')

    df.to_excel(writer, index=False, sheet_name=nom_feuille)
    worksheet = writer.sheets[nom_feuille]
    (max_row, max_col) = df.shape

    if max_col > 0:
        worksheet.autofilter(0, 0, max_row, max_col - 1)
        worksheet.freeze_panes(1, 0)
        
        total_lignes = max_row + 1
        dossier_col_letter = xl_col_to_name(df.columns.get_loc('Numéro dossier')) if 'Numéro dossier' in df.columns else None
        
        # Cadrage strict des plages de recherche (de la ligne 2 à la dernière ligne existante)
        plage_dossier = f"${dossier_col_letter}$2:${dossier_col_letter}${total_lignes}" if dossier_col_letter else ""
        plage_init = f"$A$2:$A${total_lignes}"
        
        # Création des formats une seule fois avant la boucle
        formats_init = {}
        for init, color in dict_initiales.items():
            color_clean = str(color).strip()
            if color_clean and color_clean.lower() != 'nan':
                if not color_clean.startswith('#') and len(color_clean) == 6: color_clean = f"#{color_clean}"
                formats_init[init] = writer.book.add_format({'bg_color': color_clean, 'font_color': '#000000'})

        for init, fmt_init in formats_init.items():
            if dossier_col_letter:
                formula_init = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}2, {plage_init}, "{init}")>0'
            else:
                formula_init = f'=$A2="{init}"'
            worksheet.conditional_format(1, 0, max_row, 2, {'type': 'formula', 'criteria': formula_init, 'format': fmt_init})
        
        fmt_en_cours = writer.book.add_format({'bg_color': '#FFE699', 'font_color': '#595959'})
        if dossier_col_letter:
            formula_default = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}2, {plage_init}, "<>")>0'
        else:
            formula_default = '=$A2<>""'
        worksheet.conditional_format(1, 0, max_row, 2, {'type': 'formula', 'criteria': formula_default, 'format': fmt_en_cours})
        
        for i in range(max_col): worksheet.set_column(i, i, 16) 

def ajouter_feuille_dcr(writer, df_prio, df_classique, nom_feuille, dict_initiales=None):
    if dict_initiales is None: dict_initiales = {}
    if df_prio.empty and df_classique.empty: return
    if not df_prio.empty and 'Initiales' not in df_prio.columns: df_prio.insert(0, 'Initiales', '')
    if not df_classique.empty and 'Initiales' not in df_classique.columns: df_classique.insert(0, 'Initiales', '')

    dossier_col_letter = None
    if not df_classique.empty and 'Numéro dossier' in df_classique.columns: dossier_col_letter = xl_col_to_name(df_classique.columns.get_loc('Numéro dossier'))
    elif not df_prio.empty and 'Numéro dossier' in df_prio.columns: dossier_col_letter = xl_col_to_name(df_prio.columns.get_loc('Numéro dossier'))

    workbook = writer.book
    worksheet = workbook.add_worksheet(nom_feuille)
    
    worksheet.freeze_panes(2, 0)
    
    fmt_rouge = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'red', 'align': 'center', 'valign': 'vcenter'})
    fmt_bleu = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#3a75c4', 'align': 'center', 'valign': 'vcenter'})
    fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})
    fmt_en_cours = workbook.add_format({'bg_color': '#FFE699', 'font_color': '#595959'})
    
    current_row = 0
    max_col = len(df_classique.columns) if not df_classique.empty else len(df_prio.columns)
    
    # Cadrage strict des plages de recherche pour l'onglet DCR (ignore les lignes de séparation)
    first_data_row_excel = 3
    last_data_row_excel = len(df_prio) + len(df_classique) + 5
    
    plage_dossier = f"${dossier_col_letter}${first_data_row_excel}:${dossier_col_letter}${last_data_row_excel}" if dossier_col_letter else ""
    plage_init = f"$A${first_data_row_excel}:$A${last_data_row_excel}"

    # Création des formats une seule fois pour toute la fonction DCR
    formats_init = {}
    for init, color in dict_initiales.items():
        color_clean = str(color).strip()
        if color_clean and color_clean.lower() != 'nan':
            if not color_clean.startswith('#') and len(color_clean) == 6: color_clean = f"#{color_clean}"
            formats_init[init] = workbook.add_format({'bg_color': color_clean, 'font_color': '#000000'})
    
    if not df_prio.empty:
        worksheet.merge_range(current_row, 0, current_row, max_col - 1, "↓ /!\\ Liste prioritaire /!\\ ↓", fmt_rouge)
        current_row += 1
        for i, val in enumerate(df_prio.columns.values): worksheet.write(current_row, i, val, fmt_header)
        current_row += 1
        start_prio = current_row
        df_prio.to_excel(writer, sheet_name=nom_feuille, startrow=current_row, header=False, index=False)
        end_prio = current_row + len(df_prio) - 1
        
        for init, fmt_init in formats_init.items():
            if dossier_col_letter:
                f_prio_init = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_prio + 1}, {plage_init}, "{init}")>0'
            else:
                f_prio_init = f'=$A{start_prio + 1}="{init}"'
            worksheet.conditional_format(start_prio, 0, end_prio, 2, {'type': 'formula', 'criteria': f_prio_init, 'format': fmt_init})

        if dossier_col_letter:
            formule_prio = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_prio + 1}, {plage_init}, "<>")>0'
        else:
            formule_prio = f'=$A{start_prio + 1}<>""'
        worksheet.conditional_format(start_prio, 0, end_prio, 2, {'type': 'formula', 'criteria': formule_prio, 'format': fmt_en_cours})
        
        current_row += len(df_prio)
        worksheet.merge_range(current_row, 0, current_row, max_col - 1, "↑ /!\\ Liste prioritaire /!\\ ↑", fmt_rouge)
        current_row += 1
        
    if not df_classique.empty or df_prio.empty:
        worksheet.merge_range(current_row, 0, current_row, max_col - 1, "Puis basculer sur DCR", fmt_bleu)
        current_row += 1
        ligne_entete = current_row
        for i, val in enumerate(df_classique.columns.values): worksheet.write(current_row, i, val, fmt_header)
        current_row += 1
        if not df_classique.empty:
            start_classique = current_row
            df_classique.to_excel(writer, sheet_name=nom_feuille, startrow=current_row, header=False, index=False)
            end_classique = current_row + len(df_classique) - 1
            
            for init, fmt_init in formats_init.items():
                if dossier_col_letter:
                    f_classique_init = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_classique + 1}, {plage_init}, "{init}")>0'
                else:
                    f_classique_init = f'=$A{start_classique + 1}="{init}"'
                worksheet.conditional_format(start_classique, 0, end_classique, 2, {'type': 'formula', 'criteria': f_classique_init, 'format': fmt_init})

            if dossier_col_letter:
                formule_classique = f'=COUNTIFS({plage_dossier}, ${dossier_col_letter}{start_classique + 1}, {plage_init}, "<>")>0'
            else:
                formule_classique = f'=$A{start_classique + 1}<>""'
            worksheet.conditional_format(start_classique, 0, end_classique, 2, {'type': 'formula', 'criteria': formule_classique, 'format': fmt_en_cours})
            worksheet.autofilter(ligne_entete, 0, end_classique, max_col - 1)
            
    for i in range(max_col): worksheet.set_column(i, i, 16)
