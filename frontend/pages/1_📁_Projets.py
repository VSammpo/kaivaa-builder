# frontend/pages/1_📁_Projets.py
import streamlit as st
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService

st.set_page_config(page_title="Projets", page_icon="📁", layout="wide")

st.title("📁 Projets")
st.caption("Orchestrez plusieurs livrables avec des pipelines de données configurés")

st.divider()

# Actions principales
col_title, col_action = st.columns([3, 1])

with col_title:
    st.markdown("### Mes projets")

with col_action:
    if st.button("➕ Nouveau projet", type="primary", use_container_width=True):
        st.session_state.show_create_project_modal = True
        st.rerun()

# Liste des projets
DatabaseService.initialize()
with DatabaseService.get_session() as db:
    ps = ProjectService(db)
    projects = ps.list_projects()

if not projects:
    st.info("Aucun projet pour l'instant. Créez-en un pour démarrer.")
else:
    # Affichage en grille (2 colonnes)
    cols_per_row = 2
    
    for i in range(0, len(projects), cols_per_row):
        cols = st.columns(cols_per_row)
        
        for j, col in enumerate(cols):
            idx = i + j
            if idx < len(projects):
                proj = projects[idx]
                
                with col:
                    with st.container(border=True):
                        # En-tête
                        st.markdown(f"### {proj.get('name', '(sans nom)')}")
                        st.caption(f"Client : {proj.get('client_name', '—')}")
                        
                        # Stats
                        st.markdown("")
                        col_stat1, col_stat2 = st.columns(2)
                        
                        with col_stat1:
                            nb_deliverables = len(proj.get("deliverables", []))
                            st.metric("Livrables", nb_deliverables)
                        
                        with col_stat2:
                            status = proj.get("status", "active")
                            status_emoji = "🟢" if status == "active" else "🟡"
                            st.markdown(f"{status_emoji} **{status.capitalize()}**")
                        
                        # Description
                        st.markdown("")
                        desc = proj.get("description", "")
                        if desc:
                            if len(desc) > 100:
                                desc = desc[:97] + "..."
                            st.caption(desc)
                        
                        st.markdown("")
                        
                        # Actions
                        col_btn1, col_btn2 = st.columns(2)
                        
                        with col_btn1:
                            if st.button("🗂️ Ouvrir", key=f"open_{proj['project_id']}", use_container_width=True):
                                st.session_state.selected_project_id = proj["project_id"]
                                st.switch_page("pages/_1a_🗂️_Hub_Projet.py")
                        
                        with col_btn2:
                            if st.button("🗑️ Archiver", key=f"del_{proj['project_id']}", 
                                       use_container_width=True, type="secondary"):
                                st.session_state.delete_project_id = proj["project_id"]
                                st.session_state.show_delete_modal = True
                                st.rerun()

# ===== MODALE CRÉATION PROJET =====
if st.session_state.get('show_create_project_modal'):
    @st.dialog("➕ Nouveau projet", width="large")
    def create_project_modal():
        st.markdown("### Informations du projet")
        
        name = st.text_input("Nom du projet *", placeholder="Ex: Carrefour Q4 2024")
        client = st.text_input("Client", placeholder="Ex: Carrefour")
        description = st.text_area("Description", height=100, 
                                  placeholder="Description du projet...")
        
        st.divider()
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("❌ Annuler", use_container_width=True):
                st.session_state.show_create_project_modal = False
                st.rerun()
        
        with col2:
            if st.button("✅ Créer", type="primary", use_container_width=True):
                if not name.strip():
                    st.error("❌ Le nom du projet est obligatoire")
                else:
                    try:
                        with DatabaseService.get_session() as db:
                            ps = ProjectService(db)
                            proj = ps.create_project(
                                name=name.strip(),
                                description=description.strip()
                            )
                            
                            # Ajouter client_name dans le JSON
                            proj["client_name"] = client.strip()
                            ps.save_project(proj)
                        
                        st.success(f"✅ Projet '{name}' créé !")
                        st.balloons()
                        
                        st.session_state.show_create_project_modal = False
                        st.session_state.selected_project_id = proj["project_id"]
                        
                        import time
                        time.sleep(1)
                        st.switch_page("pages/_1a_🗂️_Hub_Projet.py")
                    
                    except Exception as e:
                        st.error(f"❌ Erreur : {e}")
    
    create_project_modal()

# ===== MODALE SUPPRESSION =====
if st.session_state.get('show_delete_modal'):
    @st.dialog("⚠️ Confirmer l'archivage", width="large")
    def delete_modal():
        project_id = st.session_state.get('delete_project_id')
        proj = next((p for p in projects if p['project_id'] == project_id), None)
        
        if not proj:
            st.error("Projet introuvable")
            return
        
        project_name = proj.get('name', '(sans nom)')
        
        st.warning(f"**Vous êtes sur le point d'archiver le projet :**")
        st.markdown(f"### {project_name}")
        
        if proj.get('client_name'):
            st.caption(f"Client : {proj['client_name']}")
        
        # Avertissements
        with st.container(border=True):
            st.markdown("**📦 Ce qui sera archivé :**")
            
            nb_deliverables = len(proj.get("deliverables", []))
            nb_data_sources = len(proj.get("data_sources", []))
            
            st.markdown(f"• {nb_deliverables} livrable(s)")
            st.markdown(f"• {nb_data_sources} source(s) de données")
            st.markdown(f"• Tous les masters dupliqués")
            st.markdown(f"• Tout l'historique de génération")
        
        st.divider()
        
        st.markdown("**Cette action est réversible.** Le projet sera déplacé dans `configuration/projets/_trash/`")
        st.markdown("**Tapez le nom exact du projet pour confirmer :**")
        
        confirmation = st.text_input(
            "Nom du projet",
            key="delete_confirm_input",
            placeholder=project_name
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("❌ Annuler", use_container_width=True):
                st.session_state.show_delete_modal = False
                if 'delete_project_id' in st.session_state:
                    del st.session_state.delete_project_id
                st.rerun()
        
        with col2:
            if st.button("🗑️ Archiver définitivement", type="primary", use_container_width=True):
                if confirmation != project_name:
                    st.error("❌ Le nom ne correspond pas")
                else:
                    try:
                        with DatabaseService.get_session() as db:
                            ps = ProjectService(db)
                            archived_name = ps.soft_delete(project_id)
                        
                        st.success(f"✅ Projet archivé sous : **{archived_name}**")
                        
                        st.session_state.show_delete_modal = False
                        if 'delete_project_id' in st.session_state:
                            del st.session_state.delete_project_id
                        
                        import time
                        time.sleep(2)
                        st.rerun()
                    
                    except Exception as e:
                        st.error(f"❌ Erreur : {e}")
                        import traceback
                        with st.expander("🔍 Détails techniques"):
                            st.code(traceback.format_exc())
    
    delete_modal()