import streamlit as st
from datetime import datetime
from zoneinfo import ZoneInfo
from frontend.utils.ui_helpers import page_header

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService

st.set_page_config(page_title="Historique", page_icon="🕘", layout="wide")
page_header("Historique des exécutions", icon=None, crumbs=[("Catalogue", "1_📚_Bibliotheque"), ("Historique", "")])

PARIS = ZoneInfo("Europe/Paris")
DatabaseService.initialize()
with DatabaseService.get_session() as db:
    ts = TemplateService(db)
    jobs = ts.list_executions(limit=100) if hasattr(ts, "list_executions") else []

    if not jobs:
        st.info("Aucune exécution enregistrée.")
    else:
        for j in jobs:
            with st.container(border=True):
                tpl = ts.get_template(j.template_id)
                st.markdown(
                    f"**{tpl.name if tpl else 'Template ?'}**  ·  ID exécution #{j.id}  \n"
                    f"Statut : **{j.status}**  ·  Durée : {j.execution_time_seconds or '—'} s  \n"
                    f"Début : {j.started_at}  ·  Fin : {j.finished_at}  \n"
                    f"Projet : `{(j.parameters or {}).get('project_id', '—')}`"
                )
                c1, c2, c3 = st.columns(3)
                c1.button("Ouvrir Excel", disabled=not bool(j.excel_output_path), use_container_width=True)
                c2.button("Ouvrir PPT", disabled=not bool(j.ppt_output_path), use_container_width=True)
                if c3.button("Supprimer", use_container_width=True):
                    # tu as déjà un delete_job_and_files côté backend; expose-le si besoin
                    st.warning("Suppression à brancher (backend déjà prêt).")
