# frontend/pages/3_🧱_Gabarits.py
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from backend.services.gabarit_registry import get_role

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.gabarit_registry import list_gabarits

st.set_page_config(page_title="Gabarits", page_icon="🧱", layout="wide")

# Header
st.title("🧱 Gabarits de données")
st.caption("Structures de tables réutilisables pour vos projets")

# Action principale
col1, col2 = st.columns([3, 1])
with col2:
    if st.button("➕ Nouveau gabarit", type="primary", use_container_width=True):
        st.session_state.selected_gabarit = None
        st.switch_page("pages/_3b1_🧱_Structure_Gabarit.py")


st.divider()

# Filtres
col1, col2 = st.columns([3, 1])
with col1:
    search = st.text_input("🔍 Rechercher", placeholder="Nom du gabarit...")

# Charger les gabarits
gabarits = list_gabarits()

# Filtrer
if search:
    gabarits = [g for g in gabarits if search.lower() in g.name.lower()]

# Affichage
if not gabarits:
    st.info("Aucun gabarit trouvé. Créez-en un pour démarrer.")
else:
    st.markdown(f"**{len(gabarits)} gabarit(s) trouvé(s)**")
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
                        role = (get_role(gab.name, gab.version) or "mixed").lower()
                        color = {"fact":"#f63b44", "dimension":"#0b9ea3", "mixed":"#6b7280"}.get(role, "#6b7280")
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
                        if st.button("📊 Ouvrir", key=f"open_{gab.name}_{gab.version}", use_container_width=True):
                            st.session_state.selected_gabarit = (gab.name, gab.version)
                            st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")