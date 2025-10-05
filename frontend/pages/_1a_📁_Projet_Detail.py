# frontend/pages/1a_📁_Projet_Detail.py
# CHANGEMENTS: Navigation mise à jour ligne 10-11

import streamlit as st
import pandas as pd
from loguru import logger

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService
from backend.services.template_service import TemplateService
from backend.services.report_service import ReportService
from frontend.utils.ui_helpers import page_header
from backend.database.models import Template, ExecutionJob
from datetime import datetime, timezone

st.set_page_config(page_title="Détail Projet", page_icon="🗂️", layout="wide")
st.session_state.setdefault("selected_project_id", None)

def _goto(page_path: str):
    try:
        st.switch_page(page_path)
    except Exception:
        st.info("Utilise le menu pour naviguer.")

pid = st.session_state.get("selected_project_id")
if not pid:
    page_header("Détail du projet", icon=None, crumbs=[("Projets", "1_📁_Projets"), ("Détail du projet", "")])
    st.warning("Aucun projet sélectionné. Ouvre d'abord « Projets ».")
    st.stop()

page_header("Détail du projet", icon=None, crumbs=[("Projets", "1_📁_Projets"), ("Détail du projet", "")])

DatabaseService.initialize()
with DatabaseService.get_session() as db:
    ps = ProjectService(db)
    ts = TemplateService(db)
    proj = ps.load_project(pid)

    st.caption(f"Projet : `{proj['project_id']}` — **{proj.get('name','(sans nom)')}**")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["Templates attachés", "Union colonnes", "Pipelines", "Validation", "Exécuter"]
    )

    with tab1:
        st.subheader("Templates attachés")
        all_tpl = ts.list_templates()
        choices = {f"{t.name} (#{t.id})": t.id for t in all_tpl}
        attached_ids = set(proj.get("template_ids", []))

        colA, colB = st.columns([2, 1])
        add_sel = colA.selectbox("Ajouter un template", ["—"] + list(choices.keys()), index=0)
        if colB.button("Attacher", use_container_width=True, disabled=(add_sel == "—")):
            ps.attach_template(pid, choices[add_sel])
            st.success("Template attaché.")
            st.rerun()

        st.divider()
        if not attached_ids:
            st.info("Aucun template attaché pour l'instant.")
        else:
            for t in all_tpl:
                if t.id not in attached_ids:
                    continue
                with st.container(border=True):
                    st.markdown(f"**{t.name}**  ·  ID #{t.id}")
                    c1, c2 = st.columns([1, 3])
                    if c1.button("Détacher", key=f"detach_{t.id}", use_container_width=True):
                        ps.detach_template(pid, t.id)
                        st.rerun()
                    c2.caption("Ouvre le détail template depuis la Bibliothèque si besoin.")

    with tab2:
        st.subheader("Union colonnes (par gabarit)")
        if st.button("Recalculer l'union", type="primary"):
            ps.compute_union(pid)
            st.success("Union recalculée.")
            proj = ps.load_project(pid)

        union = proj.get("gabarit_union")
        if not union:
            st.info("Aucune union à afficher (attache au moins un template).")
        else:
            st.dataframe(pd.DataFrame(union), use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("Pipelines par gabarit")
        union = proj.get("gabarit_union") or []
        if not union:
            st.info("Rien à configurer : recalcule d'abord l'union.")
        else:
            for u in union:
                key = f"{u['gabarit_name']}__{u['gabarit_version']}"
                with st.container(border=True):
                    st.markdown(f"**{u['gabarit_name']}** · version `{u['gabarit_version']}`")
                    st.caption(f"{len(u.get('columns_required', []))} colonnes requises")
                    if st.button("Configurer le pipeline", key=f"cfg_{key}", use_container_width=True):
                        st.session_state["selected_pipeline_gab"] = (u["gabarit_name"], u["gabarit_version"])
                        _goto("pages/_1b_🔧_Pipeline_Gabarit.py")

    with tab4:
        st.subheader("Validation (non bloquante)")
        pipes = proj.get("gabarit_pipelines") or []
        if not pipes:
            st.info("Aucun pipeline défini.")
        else:
            for p in pipes:
                with st.container(border=True):
                    gname, gver = p["gabarit_name"], p["gabarit_version"]
                    st.markdown(f"**{gname}** · version `{gver}`")
                    cols = st.columns([1, 1])
                    if cols[0].button("Valider", key=f"val_{gname}_{gver}", use_container_width=True):
                        res = ps.validate(pid, gname, gver)
                        st.session_state[f"valres_{gname}_{gver}"] = res

                    last = p.get("last_validation_result")
                    if last:
                        cols[1].caption(f"Dernier résultat : {last}")

                    live = st.session_state.get(f"valres_{gname}_{gver}")
                    if live:
                        st.json(live)


    with tab5:
        st.subheader("Exécuter les livrables de ce projet")
        attached_ids = proj.get("template_ids", [])
        if not attached_ids:
            st.info("Aucun template attaché.")
        else:
            for tid in attached_ids:
                t = ts.get_template(tid)
                with st.container(border=True):
                    st.markdown(f"**{t.name}**  ·  ID #{t.id}")
                    if st.button("▶️ Lancer", key=f"run_{tid}", use_container_width=True):
                        try:
                            tpl_config = ts.load_template_config(t.id)

                            # 1) Créer un job
                            with DatabaseService.get_session() as db2:
                                job = ExecutionJob(
                                    template_id=t.id,
                                    parameters={},
                                    status='running'
                                )
                                db2.add(job)
                                db2.commit()
                                job_id = job.id

                            # 2) Exécuter
                            rs = ReportService(tpl_config)
                            result = rs.generate_report(parameters={}, project_id=pid)

                            # 3) Maj job + stats template
                            with DatabaseService.get_session() as db3:
                                job = db3.query(ExecutionJob).filter_by(id=job_id).first()
                                template_row = db3.query(Template).filter_by(id=t.id).first()

                                if result.get("success"):
                                    job.status = 'completed'
                                    job.output_ppt_path = result.get('pptx_path')
                                    job.output_excel_path = result.get('excel_path')
                                    job.execution_time_seconds = result.get('execution_time_seconds')
                                    job.completed_at = datetime.now(timezone.utc)

                                    if template_row:
                                        template_row.execution_count = (template_row.execution_count or 0) + 1
                                        template_row.last_executed = datetime.now(timezone.utc)

                                    db3.commit()

                                    st.success("OK")
                                    st.code(result['excel_path'])
                                    st.code(result['pptx_path'])
                                else:
                                    job.status = 'failed'
                                    job.error_message = result.get("error","Erreur inconnue")
                                    job.execution_time_seconds = result.get('execution_time_seconds')
                                    job.completed_at = datetime.now(timezone.utc)
                                    db3.commit()
                                    st.error(result.get("error","Erreur inconnue"))

                        except Exception as e:
                            st.error(f"Erreur run : {e}")