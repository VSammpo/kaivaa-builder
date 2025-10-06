# frontend/pages/_3a_🧱_Detail_Gabarit.py
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
from backend.services.gabarit_registry import list_methods_for_gabarit

st.set_page_config(page_title="Détail Gabarit", page_icon="🧱", layout="wide")

# CSS amélioré
st.markdown("""
<style>
/* Cartes avec ombre */
.metric-card {
    background: white;
    padding: 1rem;
    border-radius: 8px;
    border-left: 4px solid #4CAF50;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

/* Badges de rôle */
.role-badge {
    display: inline-block;
    padding: 4px 12px;
    border-radius: 12px;
    font-weight: 600;
    font-size: 0.9rem;
}
.role-fact { background: #e3f2fd; color: #1976d2; }
.role-dimension { background: #f3e5f5; color: #7b1fa2; }
.role-mixed { background: #fff3e0; color: #f57c00; }

/* En-tête gabarit */
.gabarit-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 2rem;
    border-radius: 8px;
    margin-bottom: 1.5rem;
}
</style>
""", unsafe_allow_html=True)

# Vérifier sélection
if 'selected_gabarit' not in st.session_state or not st.session_state.selected_gabarit:
    st.error("Aucun gabarit sélectionné")
    if st.button("Retour aux gabarits", use_container_width=True):
        st.switch_page("pages/3_🧱_Gabarits.py")
    st.stop()

gab_name, gab_version = st.session_state.selected_gabarit
gabarit = get_gabarit(gab_name, gab_version)

if not gabarit:
    st.error(f"Gabarit {gab_name} v{gab_version} introuvable")
    st.stop()

# ============= EN-TÊTE =============
col_back, col_title, col_actions = st.columns([1, 4, 2])

with col_back:
    if st.button("← Retour", use_container_width=True):
        del st.session_state.selected_gabarit
        st.switch_page("pages/3_🧱_Gabarits.py")

with col_title:
    st.title(f"🧱 {gabarit.name}")
    role = get_role(gabarit.name, gabarit.version) or "mixed"
    role_class = f"role-{role}"
    st.markdown(f'<span class="role-badge {role_class}">{role.upper()}</span> · Version {gabarit.version}', 
                unsafe_allow_html=True)

with col_actions:
    st.write("")  # Spacing
    col_edit, col_del = st.columns(2)
    with col_del:
        if st.button("🗑️ Supprimer", use_container_width=True):
            st.session_state.show_delete_modal_gabarit = True
            st.rerun()

st.divider()

# ============= MÉTRIQUES CLÉS =============
col_m1, col_m2, col_m3, col_m4 = st.columns(4)

with col_m1:
    st.metric("Colonnes totales", len(gabarit.columns))

with col_m2:
    n_keys = sum(1 for c in gabarit.columns if c.is_key)
    st.metric("Colonnes clés", n_keys)

with col_m3:
    rels = get_relations(gabarit.name, gabarit.version)
    st.metric("Enrichissements", len(rels or []))

with col_m4:
    # TODO: compter les templates utilisant ce gabarit
    st.metric("Templates", 0)

if gabarit.description:
    st.info(gabarit.description)

st.divider()

# ============= CONTENU PRINCIPAL EN ONGLETS =============
tab_structure, tab_enrichments, tab_methods, tab_data = st.tabs([
    "📊 Structure",
    "🔗 Enrichissements",
    "⚙️ Méthodes",
    "📁 Données par défaut"
])

# ----------- ONGLET COLONNES -----------
with tab_structure:
    st.subheader("Structure du gabarit")
    edit_col1, edit_col2 = st.columns([1,5])
    with edit_col1:
        if st.button("✏️ Éditer la structure", use_container_width=True, key="edit_structure_btn"):
            st.switch_page("pages/_3b1_🧱_Structure_Gabarit.py")

    
    cols_data = []
    for c in gabarit.columns:
        cols_data.append({
            "Nom": c.name, 
            "Type": c.type, 
            "Clé": "✓" if c.is_key else ""
        })
    
    if cols_data:
        df_cols = pd.DataFrame(cols_data)
        st.dataframe(
            df_cols, 
            use_container_width=True, 
            hide_index=True,
            height=400
        )
    else:
        st.info("Aucune colonne définie")
    
    st.caption(f"💡 {len(gabarit.columns)} colonne(s) · {n_keys} clé(s)")

# ----------- ONGLET MÉTHODES -----------
with tab_methods:
    st.subheader("Méthodes de calcul")
    
    # Charger les méthodes
    methods = list_methods_for_gabarit(gabarit.name, gabarit.version)
    
    col_btn, col_count = st.columns([2, 1])
    with col_btn:
        if st.button("⚙️ Gérer les méthodes", use_container_width=True, type="primary"):
            st.session_state.selected_gabarit = (gabarit.name, gabarit.version)
            st.switch_page("pages/_3c_⚙️_Methodes_Gabarit.py")
    with col_count:
        st.metric("Méthodes", len(methods or []))
    
    st.divider()
    
    if not methods:
        st.info("💡 Aucune méthode configurée. Cliquez sur 'Gérer les méthodes' pour en créer.")
    else:
        for idx, m in enumerate(methods):
            with st.container(border=True):
                col_name, col_output = st.columns([2, 1])
                
                with col_name:
                    st.markdown(f"**{idx+1}. {m.get('name', 'Sans nom')}**")
                    if m.get("description"):
                        st.caption(m["description"])
                
                with col_output:
                    st.markdown(f"**→** `{m.get('output_column', '?')}`")
                
                # Badges informatifs en bas
                info_parts = []
                if m.get("required_columns"):
                    cols_str = ", ".join([f"`{c}`" for c in m["required_columns"]])
                    info_parts.append(f"📥 Entrées: {cols_str}")
                if m.get("param_schema"):
                    info_parts.append(f"⚙️ {len(m['param_schema'])} paramètre(s)")
                
                if info_parts:
                    st.caption(" • ".join(info_parts))

# ----------- ONGLET ENRICHISSEMENTS -----------
with tab_enrichments:
    st.subheader("Enrichissements déclarés")
    if st.button("🔗 Gérer les enrichissements", use_container_width=True, key="edit_enrich_btn", type="primary"):
        st.switch_page("pages/_3b2_🔗_Enrichissements_Gabarit.py")

    
    rels = get_relations(gabarit.name, gabarit.version)
    
    if not rels:
        st.info("Aucun enrichissement configuré. Utilisez l'édition pour en ajouter.")
    else:
        st.caption(f"{len(rels)} enrichissement(s) configuré(s)")
        
        for idx, r in enumerate(rels):
            with st.container(border=True):
                col_info, col_badge = st.columns([4, 1])
                
                with col_info:
                    st.markdown(f"**{idx+1}. Enrichissement depuis** `{r['to_gabarit']}` [{r.get('to_version','v1')}]")
                    st.caption(f"Jointure : `{r['left_key']}` = `{r['right_key']}`")
                
                with col_badge:
                    # Récupérer le rôle de la table d'enrichissement
                    enrich_role = get_role(r['to_gabarit'], r.get('to_version', 'v1')) or "mixed"
                    role_class = f"role-{enrich_role}"
                    st.markdown(f'<span class="role-badge {role_class}">{enrich_role}</span>', 
                              unsafe_allow_html=True)
    
    st.caption("💡 Modifiez les enrichissements via le bouton 'Éditer' en haut de page")

# ----------- ONGLET DONNÉES PAR DÉFAUT -----------
with tab_data:
    st.subheader("Donnée par défaut")
    if st.button("📁 Configurer la donnée par défaut", use_container_width=True, key="edit_default_btn"):
        st.switch_page("pages/_3b3_📁_Donnee_Par_Defaut.py")

    src = get_default_source(gabarit.name, gabarit.version)
    preview = get_default_preview(gabarit.name, gabarit.version)
    
    if not src:
        st.info("Aucune donnée par défaut configurée. Utilisez l'édition pour en définir une.")
    else:
        # Informations sur la source
        col_src1, col_src2 = st.columns(2)
        with col_src1:
            st.metric("Format", src.get("type", "?").upper())
        with col_src2:
            st.metric("Python transformé", "Oui" if src.get("python") else "Non")
        
        st.code(src.get("path", ""), language=None)
        
        if src.get("python"):
            with st.expander("Code Python appliqué"):
                st.code(src.get("python"), language="python")
        
        st.divider()
        
        # Aperçu persistant
        if preview and preview.get("rows"):
            st.subheader("Aperçu (20 lignes)")
            rows = preview.get("rows") or []
            cols = preview.get("columns") or []
            df_preview = pd.DataFrame(rows)
            if cols:
                df_preview = df_preview[[c for c in cols if c in df_preview.columns]]
            
            st.dataframe(df_preview, use_container_width=True, height=400)
            st.caption(f"📊 {len(rows)} lignes · {len(cols)} colonnes")
        else:
            st.warning("Aucun aperçu disponible. Rechargez la source depuis l'édition.")
    
    st.caption("💡 Configurez ou modifiez la donnée par défaut via le bouton 'Éditer'")

# ============= MODAL SUPPRESSION =============
if st.session_state.get('show_delete_modal_gabarit'):

    @st.dialog("Confirmer la suppression")
    def delete_confirmation():
        links = count_links(gabarit.name, gabarit.version)
        
        if links > 0:
            st.error(
                f"⚠️ Ce gabarit est lié à **{links}** autre(s) gabarit(s). "
                "En le supprimant, **tous ces liens seront supprimés**."
            )
        else:
            st.warning("Cette action est irréversible.")
        
        st.divider()
        st.markdown("**Tapez le nom exact du gabarit pour confirmer :**")
        confirmation = st.text_input("Nom du gabarit", key="delete_confirm_gabarit", 
                                    placeholder=gabarit.name)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Annuler", use_container_width=True):
                st.session_state.show_delete_modal_gabarit = False
                st.rerun()
        with col2:
            if st.button("Supprimer définitivement", type="primary", use_container_width=True):
                if confirmation != gabarit.name:
                    st.error("Le nom ne correspond pas")
                    return
                
                out = soft_delete_gabarit(gabarit.name, gabarit.version)
                st.success(
                    f"✅ Gabarit supprimé et archivé sous **{out['new_name']}**"
                )
                if out['removed_relations'] > 0:
                    st.info(f"🔗 {out['removed_relations']} relation(s) supprimée(s)")
                
                st.session_state.show_delete_modal_gabarit = False
                if 'selected_gabarit' in st.session_state:
                    del st.session_state.selected_gabarit
                st.switch_page("pages/3_🧱_Gabarits.py")

    delete_confirmation()