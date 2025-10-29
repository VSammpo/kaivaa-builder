"""
Page de gestion des transformations réutilisables
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import json
from datetime import datetime

from backend.services.transformation_service import (
    list_transformations,
    get_transformation,
    delete_transformation,
    validate_transformation
)
from backend.services.gabarit_registry import list_gabarits

st.set_page_config(
    page_title="Transformations",
    page_icon="🔀",
    layout="wide"
)

st.title("🔀 Transformations")
st.markdown("Gérez vos transformations de données réutilisables")

# Initialiser la session
if "selected_transformation" not in st.session_state:
    st.session_state.selected_transformation = None

# Tabs
tab_list, tab_create = st.tabs(["📋 Liste", "➕ Créer"])

# ==================== TAB LISTE ====================
with tab_list:
    transformations = list_transformations()
    
    if not transformations:
        st.info("Aucune transformation créée. Utilisez l'onglet 'Créer' pour commencer.")
    else:
        # Barre de recherche
        search = st.text_input("🔍 Rechercher", placeholder="Nom ou description...")
        
        # Filtrer
        if search:
            transformations = [
                t for t in transformations
                if search.lower() in t.get("name", "").lower()
                or search.lower() in t.get("description", "").lower()
            ]
        
        # Affichage en grille
        cols = st.columns(3)
        for idx, transfo in enumerate(transformations):
            with cols[idx % 3]:
                with st.container(border=True):
                    # En-tête
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.markdown(f"### {transfo['name']}")
                    with col2:
                        st.caption(transfo.get("version", "v1"))
                    
                    # Description
                    desc = transfo.get("description", "")
                    if desc:
                        st.caption(desc[:100] + ("..." if len(desc) > 100 else ""))
                    
                    # Gabarit de base
                    gabarit_base = transfo.get("gabarit_base", {})
                    st.markdown(f"**Source :** {gabarit_base.get('name', 'N/A')}")
                    
                    # Stats
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        n_cols = len(transfo.get("columns_enabled", []))
                        st.metric("Colonnes", n_cols)
                    with col2:
                        n_enrich = len(transfo.get("enrichments", []))
                        st.metric("Enrichis.", n_enrich)
                    with col3:
                        n_methods = len(transfo.get("methods", []))
                        st.metric("Méthodes", n_methods)
                    
                    # Script Python ?
                    if transfo.get("overlay_python"):
                        st.success("🐍 Script Python")
                    
                    # Actions
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        if st.button("👁️ Voir", key=f"view_{idx}_{transfo['name']}_{transfo['version']}"):
                            st.session_state.selected_transformation = (transfo['name'], transfo['version'])
                            st.switch_page("pages/_2a_🔀_Detail_Transformation.py")
                    
                    with col2:
                        if st.button("✏️ Éditer", key=f"edit_{idx}_{transfo['name']}_{transfo['version']}"):
                            st.session_state.selected_transformation = (transfo['name'], transfo['version'])
                            st.session_state.edit_mode = True
                            st.switch_page("pages/_2b_🔀_Builder_Transformation.py")
                    
                    with col3:
                        if st.button("🗑️", key=f"del_{idx}_{transfo['name']}_{transfo['version']}", 
                                    help="Archiver"):
                            delete_transformation(transfo['name'], transfo['version'])
                            st.rerun()
                    
                    # Date de mise à jour
                    updated = transfo.get("updated_at", "")
                    if updated:
                        try:
                            dt = datetime.fromisoformat(updated)
                            st.caption(f"Mis à jour : {dt.strftime('%d/%m/%Y %H:%M')}")
                        except:
                            pass

# ==================== TAB CRÉER ====================
with tab_create:
    st.markdown("### Nouvelle transformation")
    
    with st.form("create_transformation"):
        col1, col2 = st.columns(2)
        
        with col1:
            name = st.text_input("Nom *", placeholder="Ex: Synthese_Ventes")
            version = st.text_input("Version", value="v1")
        
        with col2:
            # Sélection du gabarit de base
            gabarits = list_gabarits()
            gabarit_options = [f"{g.name} ({g.version})" for g in gabarits]  # PAS de v ajouté !
            
            if gabarit_options:
                selected_gab = st.selectbox("Gabarit source *", gabarit_options)
                # Parser le nom et version
                if " (" in selected_gab:
                    gab_name = selected_gab.split(" (")[0]
                    # La version est entre parenthèses, on la prend telle quelle
                    gab_version = selected_gab.split(" (")[1].rstrip(")")
                else:
                    gab_name = selected_gab
                    gab_version = "v1"
            else:
                st.error("Aucun gabarit disponible")
                gab_name = ""
                gab_version = "v1"
        
        description = st.text_area("Description", 
                                  placeholder="Décrivez l'objectif de cette transformation...")
        
        if st.form_submit_button("Créer et configurer", type="primary"):
            if not name:
                st.error("Le nom est obligatoire")
            elif not gab_name:
                st.error("Sélectionnez un gabarit source")
            else:
                # Créer la transformation
                from backend.services.transformation_service import create_transformation
                
                try:
                    transfo = create_transformation(
                        name=name,
                        version=version,
                        description=description,
                        gabarit_base={
                            "name": gab_name,
                            "version": gab_version
                        }
                    )
                    
                    # Rediriger vers le builder
                    st.session_state.selected_transformation = (name, version)
                    st.session_state.edit_mode = True
                    st.success("Transformation créée ! Redirection...")
                    st.switch_page("pages/_2b_🔀_Builder_Transformation.py")
                    
                except ValueError as e:
                    st.error(str(e))