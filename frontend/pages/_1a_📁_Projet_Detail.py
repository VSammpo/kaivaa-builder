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
from backend.services.gabarit_registry import get_default_source

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

                # ➜ Option rapide : utiliser la donnée par défaut du gabarit (si elle existe)
                default_src = get_default_source(u['gabarit_name'], u['gabarit_version'])
                if default_src:
                    if st.button("Utiliser la donnée par défaut", key=f"default_{key}", use_container_width=True):
                        ps.set_pipeline(
                            pid,
                            u['gabarit_name'],
                            u['gabarit_version'],
                            source=default_src,  # type, path, sep, encoding
                        )
                        st.success("Pipeline défini avec la donnée par défaut du gabarit.")
                        st.rerun()
                else:
                    st.caption("Aucune donnée par défaut définie pour ce gabarit.")


                st.divider()
        st.divider()
st.subheader("Jointures (activées au niveau projet)")

# Liste des joins déjà activés
joins = ps.list_joins(pid)
if joins:
    for i, j in enumerate(joins):
        with st.container(border=True):
            st.markdown(
                f"**{j['from_gabarit']}** ⟶ **{j['to_gabarit']}** "
                f"· `{j['left_key']} = {j['right_key']}` · type: `{j.get('join_type','left')}`"
            )
            colj1, colj2, colj3 = st.columns([1,1,2])
            enabled = colj1.toggle("Actif", value=bool(j.get('enabled', True)), key=f"join_on_{i}")
            if enabled != j.get("enabled", True):
                ps.toggle_join(pid, i, enabled)
                st.rerun()
            if colj2.button("Supprimer", key=f"join_del_{i}"):
                ps.remove_join_by_relation(pid, template_id=int(j["template_id"]), relation_id=j["relation_id"])
                st.rerun()
else:
    st.info("Aucune jointure activée pour ce projet.")

with st.expander("➕ Activer une relation de gabarit"):
    # 1) choisir le gabarit 'fact' (source) présent dans l'union
    union = proj.get("gabarit_union") or []
    gabs = [f"{u['gabarit_name']}|{u['gabarit_version']}" for u in union]
    src = st.selectbox("Table de FAIT (source)", gabs, index=0) if gabs else None

    # 2) choisir un template (catalogue) dans lequel les relations sont déclarées
    #    → selon ton modèle, c’est souvent le template où vivent ces gabarits
    #    on propose les templates attachés au projet
    attached_template_ids = proj.get("template_ids") or []
    tmpl_label = {tid: f"Template #{tid}" for tid in attached_template_ids}  # tu peux améliorer l'intitulé
    selected_tid = st.selectbox("Catalogue (template) contenant la relation", attached_template_ids) if attached_template_ids else None

    rel_choice = None
    if src and selected_tid:
        s_name, s_ver = src.split("|")
        # 3) on récupère les relations depuis le CATALOGUE pour ce 'from_gabarit'
        rels = ts.list_relations(int(selected_tid), from_gabarit=s_name, from_version=s_ver)
        if not rels:
            st.warning("Aucune relation déclarée dans ce template pour ce gabarit.")
        else:
            # affichage lisible
            rel_lbl = [
                f"{r['from_gabarit']}[{r.get('from_version','v1')}] "
                f"— {r['left_key']} = {r['right_key']} —> "
                f"{r['to_gabarit']}[{r.get('to_version','v1')}]"
            for r in rels]
            idx = st.selectbox("Relation de gabarit", list(range(len(rels))), format_func=lambda i: rel_lbl[i])

            join_type = st.selectbox("Type de jointure", ["left","inner"], index=0)

            if st.button("Activer la relation", type="primary", use_container_width=True):
                ps.add_join_by_relation(
                    pid,
                    template_id=int(selected_tid),
                    relation_id=rels[idx]["relation_id"],
                    join_type=join_type,
                    enabled=True,
                )
                st.success("Relation activée pour ce projet.")
                st.rerun()


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