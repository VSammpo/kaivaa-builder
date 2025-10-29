# frontend/pages/3_🧱_Gabarits.py
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.gabarit_registry import (
    list_gabarits, 
    get_role, 
    get_unique_roles,
    get_gabarits_by_role,
    get_gabarits_related_to
)

st.set_page_config(page_title="Gabarits", page_icon="🧱", layout="wide")

# Header
st.title("🧱 Gabarits de données")
st.caption("Structures de tables réutilisables pour vos projets")

# Action principale
col1, col2 = st.columns([3, 1])
with col2:
    if st.button("➕ Nouveau gabarit", type="primary", use_container_width=True):
        st.session_state.selected_gabarit = None
        st.switch_page("pages/_1b_🧱_Structure_Gabarit.py")

st.divider()

# ============= ZONE DE FILTRES =============
st.subheader("🔍 Filtres")

# Ligne 1 de filtres : Recherche textuelle + Filtre par rôle
col_search, col_role = st.columns([2, 1])

with col_search:
    search = st.text_input(
        "🔎 Rechercher par nom", 
        placeholder="Ex: SELL-IN, Dim_clients...",
        label_visibility="visible"
    )

with col_role:
    # Récupérer les rôles disponibles
    available_roles = get_unique_roles()
    role_options = ["Tous"] + [r.capitalize() for r in available_roles if r]
    
    selected_role = st.selectbox(
        "📊 Type de table",
        options=role_options,
        index=0,
        help="Filtrer par type de gabarit (Fact, Dimension, Mixed)"
    )

# Ligne 2 de filtres : Tables liées
col_related, col_spacer = st.columns([2, 1])

with col_related:
    # Liste de tous les gabarits pour le sélecteur
    all_gabarits = list_gabarits()
    gabarit_options = ["Aucun filtre"] + [f"{g.name} [{g.version}]" for g in all_gabarits]
    
    selected_related = st.selectbox(
        "🔗 Tables liées à",
        options=gabarit_options,
        index=0,
        help="Afficher uniquement les tables qui ont une relation d'enrichissement avec la table sélectionnée"
    )

st.divider()

# ============= APPLICATION DES FILTRES =============

# Initialiser avec tous les gabarits
gabarits = list_gabarits()

# Filtre 1: Par rôle
if selected_role != "Tous":
    role_lower = selected_role.lower()
    gabarits = get_gabarits_by_role(role_lower)

# Filtre 2: Tables liées
if selected_related != "Aucun filtre":
    # Extraire le nom et la version du gabarit sélectionné
    try:
        # Format: "NOM [VERSION]"
        parts = selected_related.rsplit(" [", 1)
        ref_name = parts[0]
        ref_version = parts[1].rstrip("]") if len(parts) > 1 else "v1"
        
        # Récupérer les gabarits liés
        related_gabarits = get_gabarits_related_to(ref_name, ref_version)
        
        # Ajouter aussi le gabarit de référence lui-même
        ref_gab = next((g for g in gabarits if g.name == ref_name and g.version == ref_version), None)
        if ref_gab:
            related_gabarits.append(ref_gab)
        
        # Filtrer la liste
        gabarits = [g for g in gabarits if any(
            rg.name == g.name and rg.version == g.version 
            for rg in related_gabarits
        )]
    except Exception as e:
        st.error(f"Erreur lors du filtrage par tables liées: {e}")

# Filtre 3: Recherche textuelle
if search:
    gabarits = [g for g in gabarits if search.lower() in g.name.lower()]

# ============= AFFICHAGE DES RÉSULTATS =============

if not gabarits:
    st.info("Aucun gabarit trouvé avec ces critères. Essayez de modifier les filtres.")
else:
    # Afficher le compteur avec les filtres actifs
    filter_tags = []
    if selected_role != "Tous":
        filter_tags.append(f"Type: {selected_role}")
    if selected_related != "Aucun filtre":
        filter_tags.append(f"Liées à: {selected_related}")
    if search:
        filter_tags.append(f"Recherche: '{search}'")
    
    filter_text = " • ".join(filter_tags) if filter_tags else "Aucun filtre actif"
    
    st.markdown(f"**{len(gabarits)} gabarit(s) trouvé(s)**")
    st.caption(f"🏷️ {filter_text}")
    st.markdown("")
    
    # Grid 3 colonnes
    cols_per_row = 3
    
    for i in range(0, len(gabarits), cols_per_row):
        cols = st.columns(cols_per_row)
        
        for j, col in enumerate(cols):
            idx = i + j
            if idx < len(gabarits):
                gab = gabarits[idx]
                
                with col:
                    with st.container(border=True):
                        # Badge de rôle avec couleur
                        role = (get_role(gab.name, gab.version) or "mixed").lower()
                        color = {
                            "fact": "#f63b44", 
                            "dimension": "#0b9ea3", 
                            "mixed": "#6b7280"
                        }.get(role, "#6b7280")
                        
                        st.markdown(
                            f'<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:6px"></span>'
                            f'<span style="color:{color};font-weight:600;text-transform:uppercase">{role}</span>',
                            unsafe_allow_html=True,
                        )

                        # Nom + version
                        st.markdown(f"### {gab.name}")
                        st.caption(f"Version {gab.version}")
                        
                        # Stats
                        st.markdown("")
                        col_stat1, col_stat2 = st.columns(2)
                        with col_stat1:
                            st.metric("Colonnes", len(gab.columns))
                        with col_stat2:
                            n_keys = sum(1 for c in gab.columns if c.is_key)
                            st.metric("Clés", n_keys)
                        
                        # Description
                        st.markdown("")
                        desc = gab.description or "Aucune description"
                        if len(desc) > 80:
                            desc = desc[:77] + "..."
                        st.caption(desc)
                        
                        st.markdown("")
                        
                        # Bouton d'accès
                        if st.button(
                            "📊 Ouvrir", 
                            key=f"open_{gab.name}_{gab.version}", 
                            use_container_width=True
                        ):
                            st.session_state.selected_gabarit = (gab.name, gab.version)
                            st.switch_page("pages/_1a_🧱_Detail_Gabarit.py")

# ============= AIDE / LÉGENDE =============
st.divider()
with st.expander("ℹ️ À propos des filtres"):
    st.markdown("""
    ### Filtres disponibles
    
    **🔎 Recherche par nom**
    - Filtrez les gabarits en tapant une partie de leur nom
    - Non sensible à la casse
    
    **📊 Type de table**
    - **Fact** : Tables de faits (événements, transactions, métriques)
    - **Dimension** : Tables de référence (clients, produits, temps)
    - **Mixed** : Tables hybrides ou non classifiées
    
    **🔗 Tables liées**
    - Affiche uniquement les tables qui ont une relation d'enrichissement avec la table sélectionnée
    - Inclut les tables qui enrichissent ET les tables enrichies par la table de référence
    - Utile pour visualiser rapidement votre modèle de données
    
    ### Combiner les filtres
    Tous les filtres peuvent être combinés pour affiner votre recherche.
    """)