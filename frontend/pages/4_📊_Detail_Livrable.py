"""
Page de détail d'un livrable - Version optimisée
"""

import streamlit as st
from pathlib import Path
import sys
import subprocess
import platform
from datetime import datetime

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService

st.set_page_config(page_title="Détail Livrable", page_icon="📊", layout="wide")

# Vérifier qu'un template est sélectionné
if 'selected_template_detail' not in st.session_state:
    st.error("Aucun template sélectionné")
    if st.button("Retour à la bibliothèque"):
        st.switch_page("pages/1_📚_Bibliotheque.py")
    st.stop()

template_id = st.session_state.selected_template_detail

# Charger les données
with DatabaseService.get_session() as db:
    service = TemplateService(db)
    template = service.get_template(template_id)
    
    if not template:
        st.error(f"Template {template_id} introuvable")
        st.stop()
    
    config = service.load_template_config(template_id)
    stats = service.get_template_stats(template_id)
    
    # Extraire données
    template_name = template.name
    template_version = template.version
    template_description = template.description
    ppt_path = template.ppt_template_path
    excel_path = template.excel_template_path

# En-tête cliquable
if st.button(f"📊 {template_name} (v{template_version})", key="header_deselect", use_container_width=True):
    del st.session_state.selected_template_detail
    st.switch_page("pages/1_📚_Bibliotheque.py")

st.caption("Cliquez sur le titre pour retourner à la bibliothèque")

if template_description:
    st.info(template_description)

st.divider()

# Layout principal
col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("Actions")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("▶️ Générer", use_container_width=True, type="primary"):
            st.session_state.selected_template_for_generation = template_id
            st.switch_page("pages/3_▶️_Generer_Rapport.py")
    
    with col2:
        if st.button("✏️ Éditer", use_container_width=True):
            st.session_state.selected_template = template_id
            st.switch_page("pages/2_➕_Nouveau_Template.py")
    
    with col3:
        if st.button("🗑️ Supprimer", use_container_width=True):
            st.session_state.show_delete_modal = True
            st.rerun()
    
    st.markdown("")
    
    st.subheader("Éditer les fichiers master")
    
    def open_file(filepath):
        try:
            filepath_abs = str(Path(filepath).resolve())
            if platform.system() == 'Windows':
                subprocess.run(['cmd', '/c', 'start', '', filepath_abs], check=True)
            elif platform.system() == 'Darwin':
                subprocess.run(['open', filepath_abs], check=True)
            else:
                subprocess.run(['xdg-open', filepath_abs], check=True)
            return True
        except Exception as e:
            st.error(f"Erreur : {e}")
            return False
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Excel", use_container_width=True, disabled=not (excel_path and Path(excel_path).exists())):
            if open_file(excel_path):
                st.toast("Excel ouvert")
    
    with col2:
        if st.button("📄 PPT", use_container_width=True, disabled=not (ppt_path and Path(ppt_path).exists())):
            if open_file(ppt_path):
                st.toast("PowerPoint ouvert")
    
    st.markdown("")
    
    st.subheader("Importer / Éditer les tables")
    st.button("⚙️ Configurer", disabled=True, use_container_width=True)
    st.caption("Fonctionnalité à venir")
    
    st.markdown("")
    
    st.subheader("Statistiques d'utilisation")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("Total", stats['total_executions'])
        st.metric("Durée moy.", f"{stats['avg_execution_time_seconds']}s")
    
    with col2:
        st.metric("Succès", f"{stats['success_rate']}%")
        st.metric("Échecs", stats['failed_executions'])

with col_right:
    st.subheader("Historique des téléchargements")
    
    from backend.database.models import ExecutionJob
    
    # Forcer refresh des données
    with DatabaseService.get_session() as db:
        recent_jobs = db.query(ExecutionJob).filter_by(
            template_id=template_id
        ).order_by(ExecutionJob.created_at.desc()).limit(15).all()
        
        jobs_data = []
        for job in recent_jobs:
            jobs_data.append({
                'id': job.id,
                'date': job.created_at,
                'status': job.status,
                'duration': job.execution_time_seconds,
                'excel_path': job.output_excel_path,
                'ppt_path': job.output_ppt_path,
                'parameters': job.parameters,
                'error': job.error_message
            })
    
    if jobs_data:
        with st.container(height=600):
            for job in jobs_data:
                
                # Ligne principale avec boutons
                col_status, col_date, col_excel, col_ppt, col_actions = st.columns([1, 3, 2, 2, 1])

                with col_status:
                    if job['status'] == 'running':
                        st.markdown("🔄")
                    elif job['status'] == 'completed':
                        st.markdown("✅")
                    else:
                        st.markdown("❌")

                with col_date:
                    # Sécurise l'affichage même si la date est naive
                    dt = job['date']
                    if hasattr(dt, "tzinfo") and dt.tzinfo is not None:
                        date_str = dt.astimezone().strftime('%d/%m/%Y %H:%M')
                    else:
                        date_str = dt.strftime('%d/%m/%Y %H:%M')
                    duration_str = f"- {job['duration']:.1f}s" if job['duration'] else ""
                    st.markdown(f"**{date_str}** {duration_str}")

                with col_excel:
                    excel_exists = job['excel_path'] and Path(job['excel_path']).exists()
                    if st.button("📊 Excel", key=f"excel_{job['id']}",
                                 disabled=not excel_exists or job['status'] != 'completed',
                                 use_container_width=True):
                        if open_file(job['excel_path']):
                            st.toast("Excel ouvert")

                with col_ppt:
                    ppt_exists = job['ppt_path'] and Path(job['ppt_path']).exists()
                    if st.button("📄 PPT", key=f"ppt_{job['id']}",
                                 disabled=not ppt_exists or job['status'] != 'completed',
                                 use_container_width=True):
                        if open_file(job['ppt_path']):
                            st.toast("PowerPoint ouvert")

                with col_actions:
                    if st.button("🗑️", key=f"del_{job['id']}", use_container_width=True,
                                 help="Supprimer cette exécution (fichiers + KPI)"):
                        with DatabaseService.get_session() as db_del:
                           ok = DatabaseService.delete_job_and_files(db_del, job['id'])

                        if ok:
                            st.success("Exécution supprimée")
                        else:
                            st.warning("Exécution introuvable")
                        st.rerun()

                
                # Expander pour détails
                with st.expander("📋 Détails", expanded=False):
                    st.caption(f"**Heure de génération :** {job['date'].strftime('%H:%M:%S')}")
                    
                    if job['duration']:
                        st.caption(f"**Durée :** {job['duration']:.2f} secondes")
                    
                    if job['parameters']:
                        st.caption("**Paramètres :**")
                        st.json(job['parameters'])
                    
                    if job['error']:
                        st.error(f"**Erreur :** {job['error']}")
                    
                    if job['excel_path']:
                        st.caption(f"**Excel :** `{job['excel_path']}`")
                    if job['ppt_path']:
                        st.caption(f"**PPT :** `{job['ppt_path']}`")
                
                st.divider()
    else:
        st.info("Aucune génération pour ce template")

# Modal suppression
if st.session_state.get('show_delete_modal'):
    
    @st.dialog("Confirmer la suppression")
    def delete_confirmation():
        st.warning(f"Suppression du template **{template_name}**")
        st.markdown("Tapez le nom exact pour confirmer :")
        
        confirmation = st.text_input("Nom du template", key="delete_confirm_input")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Annuler", use_container_width=True):
                st.session_state.show_delete_modal = False
                st.rerun()
        
        with col2:
            if st.button("Supprimer", type="primary", use_container_width=True):
                if confirmation == template_name:
                    with DatabaseService.get_session() as db:
                        service = TemplateService(db)
                        service.delete_template(template_id, hard_delete=False)
                    
                    st.success(f"Template '{template_name}' désactivé")
                    st.session_state.show_delete_modal = False
                    del st.session_state.selected_template_detail
                    st.switch_page("pages/1_📚_Bibliotheque.py")
                else:
                    st.error("Le nom ne correspond pas")
    
    delete_confirmation()