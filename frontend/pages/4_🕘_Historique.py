# frontend/pages/4_🕘_Historique.py
# (Ancien: 6_🕘_Historique.py)
# CHANGEMENTS: Numérotation 4, breadcrumbs mis à jour, design optimisé

import streamlit as st
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.database.models import ExecutionJob

st.set_page_config(page_title="Historique", page_icon="🕘", layout="wide")

st.title("🕘 Historique des exécutions")
st.caption("Retrouvez rapidement tous vos livrables générés")

st.divider()

PARIS = ZoneInfo("Europe/Paris")

def fmt_paris(ts):
    if ts is None:
        return "—"
    if isinstance(ts, str):
        try:
            ts = datetime.fromisoformat(ts)
        except:
            return ts
    if isinstance(ts, datetime):
        if ts.tzinfo is not None:
            ts = ts.astimezone(PARIS)
        return ts.strftime("%d/%m/%Y %H:%M")
    return str(ts)

DatabaseService.initialize()
with DatabaseService.get_session() as db:
    # Récupérer les 50 dernières exécutions
    jobs = db.query(ExecutionJob).order_by(
        ExecutionJob.created_at.desc()
    ).limit(50).all()

    if not jobs:
        st.info("Aucune exécution enregistrée pour le moment.")
        st.caption("Les rapports générés depuis vos projets apparaîtront ici.")
    else:
        st.markdown(f"**{len(jobs)} exécution(s) récente(s)**")
        st.markdown("")
        
        # Conteneur scrollable
        with st.container(height=700):
            for job in jobs:
                with st.container(border=True):
                    # Ligne statut + date
                    col_status, col_info, col_actions = st.columns([1, 6, 2])
                    
                    with col_status:
                        status_icons = {
                            'completed': '✅',
                            'failed': '❌',
                            'running': '⏳'
                        }
                        st.markdown(f"## {status_icons.get(job.status, '❓')}")
                    
                    with col_info:
                        # Infos principales
                        from backend.services.template_service import TemplateService
                        ts = TemplateService(db)
                        tpl = ts.get_template(job.template_id)
                        template_name = tpl.name if tpl else f"Template #{job.template_id}"
                        
                        st.markdown(f"**{template_name}**")
                        st.caption(f"{fmt_paris(job.created_at)} · Durée : {job.execution_time_seconds or '—'} s")
                        
                        if job.error_message:
                            st.error(f"Erreur : {job.error_message}")
                    
                    with col_actions:
                        # Boutons d'action
                        excel_exists = job.output_excel_path and Path(job.output_excel_path).exists()
                        ppt_exists = job.output_ppt_path and Path(job.output_ppt_path).exists()
                        
                        if st.button("📊 Excel", key=f"excel_{job.id}", 
                                   disabled=not excel_exists,
                                   use_container_width=True):
                            import subprocess, platform
                            filepath = str(Path(job.output_excel_path).resolve())
                            if platform.system() == 'Windows':
                                subprocess.run(['cmd', '/c', 'start', '', filepath])
                            st.toast("Excel ouvert")
                        
                        if st.button("📄 PPT", key=f"ppt_{job.id}",
                                   disabled=not ppt_exists,
                                   use_container_width=True):
                            import subprocess, platform
                            filepath = str(Path(job.output_ppt_path).resolve())
                            if platform.system() == 'Windows':
                                subprocess.run(['cmd', '/c', 'start', '', filepath])
                            st.toast("PowerPoint ouvert")