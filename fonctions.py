# fonctions.py
import datetime

def parse_numero_complet(texte_complet):
    """Sépare '14.5/17-2273_v6' en ('14.5/17-2273', 'V6') pour la base de données."""
    if not texte_complet:
        return "", "V1"
    if "_" in texte_complet:
        parties = texte_complet.rsplit("_", 1)
        return parties[0].strip(), parties[1].strip().upper()
    return texte_complet.strip(), "V1"

def filtrer_et_trier_avis(donnees_atec, filtre_marque, filtre_texte, filtre_date):
    """Filtre les avis techniques selon les critères et les trie par pertinence/révision."""
    # 1. Trouver les familles (numéros de base) d'avis correspondant aux filtres texte/marque
    bases_retenues = set()
    for doc in donnees_atec:
        num_atec = str(doc.get('numero_atec', ''))
        base_num = num_atec.split('_')[0].strip()
        distributeur = str(doc.get('distributeur', '')).lower()
        titulaire = str(doc.get('titulaire', '')).lower()
        
        match_marque = True
        if filtre_marque != "Toutes":
            match_marque = (filtre_marque.lower() in distributeur) or (filtre_marque.lower() in titulaire)
        
        match_texte = True
        if filtre_texte:
            txt = filtre_texte.lower()
            match_model = any(txt in str(m.get('nom_modele', '')).lower() for m in doc.get('modeles', []) if isinstance(m, dict))
            match_texte = (txt in num_atec.lower()) or (txt in distributeur) or match_model
        
        if match_marque and match_texte:
            bases_retenues.add(base_num)
    
    # 2. Re-sélectionner TOUTES les versions pour ces familles (évite la perte d'historique)
    resultats_filtres = []
    for doc in donnees_atec:
        base_num = str(doc.get('numero_atec', '')).split('_')[0].strip()
        if base_num in bases_retenues:
            if filtre_date:
                try:
                    deb = datetime.datetime.strptime(doc.get('debut_validite', ''), "%Y-%m-%d").date() if doc.get('debut_validite') else None
                    fin = datetime.datetime.strptime(doc.get('fin_validite', ''), "%Y-%m-%d").date() if doc.get('fin_validite') else None
                    if deb and fin and not (deb <= filtre_date <= fin): continue
                    elif deb and filtre_date < deb: continue
                    elif fin and filtre_date > fin: continue
                except:
                    pass
            resultats_filtres.append(doc)

    if not resultats_filtres:
        return []

    # 3. Tri ordonné : par numéro de base, puis par indice de révision décroissant
    def get_sort_key(d):
        base = str(d.get('numero_atec', '')).split('_')[0].strip()
        rev_str = str(d.get('indice_revision', 'V1')).upper().replace('V', '').replace('MODIFICATIF', '').strip()
        try: 
            rev_num = int(rev_str)
        except ValueError: 
            rev_num = 0
        return (base, -rev_num)

    resultats_finaux = sorted(resultats_filtres, key=get_sort_key)

    # 4. Identification dynamique de la version la plus haute à l'écran
    bases_vues = set()
    for d in resultats_finaux:
        base = str(d.get('numero_atec', '')).split('_')[0].strip()
        if base not in bases_vues:
            d['_est_version_recente'] = True
            bases_vues.add(base)
        else:
            d['_est_version_recente'] = False

    return resultats_finaux

def nettoyer_modeles_ia(donnees_extraites):
    """Filtre post-IA pour supprimer les faux modèles et nettoyer les données extraites."""
    if "modeles" in donnees_extraites:
        vrais_modeles = []
        mots_interdits = [
            "débits", "décroissants", "config", "bouches", "pmin", 
            "multipiquage", "courbe", "caractéristique",
            "b100", "b200", "fan_", "-fan", "t.flow", "thermodynamique"
        ]
        
        for m in donnees_extraites["modeles"]:
            nom = str(m.get("nom_modele", "")).lower()
            if len(nom) < 45 and not any(mot in nom for mot in mots_interdits):
                vrais_modeles.append(m)
                
        donnees_extraites["modeles"] = vrais_modeles
    return donnees_extraites
