# frontend/pages/3a_🧱_Detail_Gabarit.py
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from backend.services.gabarit_registry import get_default_source, get_default_preview

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.gabarit_registry import get_gabarit
from backend.services.gabarit_registry import get_relations, get_role
from backend.services.gabarit_registry import count_links, soft_delete_gabarit

st.set_page_config(page_title="Détail Gabarit", page_icon="🧱", layout="wide")

# Vérifier sélection
if 'selected_gabarit' not in st.session_state or not st.session_state.selected_gabarit:
    st.error("Aucun gabarit sélectionné")
    if st.button("Retour aux gabarits"):
        st.switch_page("pages/3_🧱_Gabarits.py")
    st.stop()

gab_name, gab_version = st.session_state.selected_gabarit
gabarit = get_gabarit(gab_name, gab_version)

if not gabarit:
    st.error(f"Gabarit {gab_name} v{gab_version} introuvable")
    st.stop()

# Header cliquable
if st.button(f"🧱 {gabarit.name} (v{gabarit.version})", use_container_width=True):
    del st.session_state.selected_gabarit
    st.switch_page("pages/3_🧱_Gabarits.py")

st.caption("Cliquez sur le titre pour retourner à la liste")

if gabarit.description:
    st.info(gabarit.description)

st.divider()

# Layout 2 colonnes fixes
col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("Actions")
    st.subheader("Rôle")
    role = get_role(gabarit.name, gabarit.version) or "mixed"
    st.markdown(f"**{role.upper()}**")

    st.markdown("")
    st.subheader("Relations (catalogue)")
    rels = get_relations(gabarit.name, gabarit.version)
    if not rels:
        st.caption("Aucune relation déclarée.")
    else:
        with st.container(height=200):
            for r in rels:
                st.markdown(
                    f"- `{r['left_key']}` = `{r['right_key']}` → "
                    f"**{r['to_gabarit']}[{r.get('to_version','v1')}]**"
                )

    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("✏️ Éditer", use_container_width=True, type="primary"):
            st.switch_page("pages/_3b_➕_Form_Gabarit.py")
    
    with col2:
        if st.button("🗑️ Supprimer", use_container_width=True):
            st.session_state.show_delete_modal_gabarit = True
            st.rerun()
    
    st.markdown("")
    st.subheader("Métadonnées")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Colonnes totales", len(gabarit.columns))
    with col2:
        n_keys = sum(1 for c in gabarit.columns if c.is_key)
        st.metric("Colonnes clés", n_keys)
    
    st.markdown("")
    st.subheader("Colonnes")
    
    # Liste colonnes dans container scrollable
    with st.container(height=300):
        cols_data = []
        for c in gabarit.columns:
            cols_data.append({
                "Nom": c.name,
                "Type": c.type,
                "Clé": "✓" if c.is_key else ""
            })
        
        if cols_data:
            st.dataframe(
                pd.DataFrame(cols_data),
                use_container_width=True,
                hide_index=True
            )

with col_right:
    st.subheader("Méthodes disponibles")
    
    # Placeholder pour méthodes (à implémenter selon registre)
    with st.container(height=400):
        st.info("Section méthodes à venir")
        st.caption("Les méthodes permettront d'appliquer des transformations standard sur ce gabarit")
    
    st.markdown("")
    st.subheader("Utilisation")
    
    st.metric("Templates utilisant ce gabarit", 0)
    st.caption("Fonctionnalité à venir : liste des templates rattachés")

# Modal suppression
if st.session_state.get('show_delete_modal_gabarit'):

    @st.dialog("Confirmer la suppression")
    def delete_confirmation():
        links = count_links(gabarit.name, gabarit.version)
        st.warning(
            f"Attention, ce gabarit est lié à **{links}** autre(s) gabarit(s). "
            "En le supprimant, **tous ces liens seront supprimés**.",
            icon="⚠️",
        )
        st.markdown("Tapez le nom exact pour confirmer :")
        confirmation = st.text_input("Nom du gabarit", key="delete_confirm_gabarit")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Annuler", use_container_width=True):
                st.session_state.show_delete_modal_gabarit = False
                st.rerun()
        with col2:
            if st.button("Supprimer", type="primary", use_container_width=True):
                if confirmation != gabarit.name:
                    st.error("Le nom ne correspond pas")
                    return
                out = soft_delete_gabarit(gabarit.name, gabarit.version)
                st.success(
                    f"Gabarit supprimé. Archivé sous **{out['new_name']}**. "
                    f"Relations supprimées : **{out['removed_relations']}**."
                )
                st.session_state.show_delete_modal_gabarit = False
                if 'selected_gabarit' in st.session_state:
                    del st.session_state.selected_gabarit
                st.switch_page("pages/3_🧱_Gabarits.py")

    delete_confirmation()

st.divider()
st.subheader("Donnée par défaut")

src = get_default_source(gabarit.name, gabarit.version)
preview = get_default_preview(gabarit.name, gabarit.version)

if not src:
    st.caption("Aucune donnée par défaut mémorisée.")
else:
    # Affiche un rappel de la source
    t = src.get("type", "?")
    path_info = src.get("path", "")
    st.markdown(f"**Source :** `{t}` · `{path_info}`")

    # Mini tableau (persistant)
    if preview and preview.get("rows"):
        rows = preview.get("rows") or []
        cols = preview.get("columns") or []
        df_preview = pd.DataFrame(rows)
        # Respecte l'ordre des colonnes sauvegardé
        if cols:
            df_preview = df_preview[[c for c in cols if c in df_preview.columns]]
        st.caption("Aperçu persistant (20 lignes max)")
        st.dataframe(df_preview, use_container_width=True, height=260)
    else:
        st.caption("Aucun aperçu persistant enregistré pour cette source.")
