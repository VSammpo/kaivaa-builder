# frontend/pages/_2a_📊_Detail_Livrable.py
# Page HUB (lecture seule) d'un template : nav + (col gauche: infos) + (col droite: historique téléchargements)

import streamlit as st
from pathlib import Path
import sys
import subprocess
import platform
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
from backend.services.report_service import ReportService

# ===== Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService

# ===== Utils
def fmt_paris(ts) -> str:
    if ts is None:
        return "—"
    if isinstance(ts, str):
        try:
            ts = datetime.fromisoformat(ts)
        except Exception:
            return ts
    if isinstance(ts, datetime):
        if ts.tzinfo is not None:
            ts = ts.astimezone(ZoneInfo("Europe/Paris"))
        return ts.strftime("%d/%m/%Y %H:%M")
    return str(ts)

def open_file(filepath: str) -> bool:
    try:
        if not filepath:
            st.error("Chemin vide.")
            return False
        abspath = str(Path(filepath).resolve())
        if platform.system() == "Windows":
            subprocess.run(["cmd", "/c", "start", "", abspath], check=True)
        elif platform.system() == "Darwin":
            subprocess.run(["open", abspath], check=True)
        else:
            subprocess.run(["xdg-open", abspath], check=True)
        return True
    except Exception as e:
        st.error(f"Erreur d'ouverture : {e}")
        return False

st.set_page_config(page_title="Détail du template", page_icon="🗂️", layout="wide")

# ===== Guard : un template doit être sélectionné
if "selected_template_detail" not in st.session_state or not st.session_state.selected_template_detail:
    st.error("Aucun template sélectionné.")
    if st.button("← Retour bibliothèque", use_container_width=True):
        st.switch_page("pages/2_📚_Bibliotheque.py")
    st.stop()

template_id = st.session_state.selected_template_detail

# ===== Charger PRIMITIFS du template (évite Detached)
with DatabaseService.get_session() as db:
    ts = TemplateService(db)
    tpl = ts.get_template(template_id)
    if not tpl:
        st.error(f"Template #{template_id} introuvable.")
        st.stop()
    cfg = ts.get_config(template_id)  # dict
    stats = ts.get_template_stats(template_id)  # KPIs agrégés si dispo

    tpl_name = tpl.name
    tpl_version = tpl.version
    tpl_desc = tpl.description
    ppt_path = tpl.ppt_template_path
    excel_path = tpl.excel_template_path

# ===== Navbar (4 boutons unifiés)
def render_template_subnav(active: str, template_id: int):
    cols = st.columns([1,1,1,1,1])
    with cols[0]:
        if st.button("← Retour bibliothèque", use_container_width=True):
            if "selected_template" in st.session_state:
                del st.session_state.selected_template
            st.switch_page("pages/2_📚_Bibliotheque.py")
    with cols[1]:
        st.button("🗂️ Détail du template", type="primary" if active=="detail" else "secondary", use_container_width=True)
    with cols[2]:
        if st.button("⚙️ Paramètres généraux", type=("primary" if active=="general" else "secondary"), use_container_width=True):
            st.session_state.selected_template = template_id
            st.switch_page("pages/_2b_➕_Form_Template.py")
    with cols[3]:
        if st.button("📑 Injection des données", type=("primary" if active=="inject" else "secondary"), use_container_width=True):
            st.session_state.selected_template = template_id
            st.switch_page("pages/_2b3_📑_Tables_Template.py")
    with cols[4]:
        if st.button("🧾 Ajustement de la table", type=("primary" if active=="adjust" else "secondary"), use_container_width=True):
            st.session_state.selected_template = template_id
            st.switch_page("pages/_2b4_🧾_Ajustement_Table.py")
    st.divider()


render_template_subnav("detail", template_id)

# ===== Titre + description + actions
col_title, col_actions = st.columns([4, 1])

with col_title:
    st.title(f"🗂️ Détail du template — {tpl_name} (v{tpl_version})")
    if tpl_desc:
        st.info(tpl_desc)

with col_actions:
    st.write("")
    st.write("")

    if st.button("▶️ Créer livrable (données par défaut)", use_container_width=True, type="primary", key="btn_build_from_defaults"):
        from backend.services.database_service import DatabaseService
        with DatabaseService.get_session() as db:
            try:
                ts_local = TemplateService(db)
                cfg_obj = ts_local.load_template_config(template_id)

                defaults = {}
                for p in getattr(cfg_obj, "parameters", []):
                    val = getattr(p, "default", None)
                    defaults[p.name] = (val if val is not None else "")

                rs = ReportService(cfg_obj)

                with st.spinner("Génération du livrable à partir des données par défaut…"):
                    result = rs.generate_report(parameters=defaults, project_id=None)

                if result.get("success"):
                    st.success("✅ Livrable créé. Il apparaît dans l'historique ci-dessous.")
                    st.rerun()
                else:
                    st.error(f"Échec de la génération : {result.get('error', 'erreur inconnue')}")
            except Exception as e:
                st.error(f"Erreur lors de la génération : {e}")


    # Bouton Supprimer (existant)
    if st.button("🗑️ Supprimer", use_container_width=True, type="secondary", key="btn_delete_template"):
        st.session_state.delete_template_detail_id = template_id
        st.session_state.show_delete_modal_detail = True
        st.rerun()


st.divider()

# ===== Layout 2 colonnes
col_left, col_right = st.columns([1, 1], gap="large")

# ---------------------------------------------------------------------
# COLONNE GAUCHE : FICHIERS MASTER + TABLES DEMANDÉES + KPIs
# ---------------------------------------------------------------------
with col_left:
    st.subheader("📁 Fichiers master (lecture seule)")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Chemin PPT", value=str(ppt_path or ""), disabled=True)
        if st.button("📂 Ouvrir PPT", use_container_width=True, disabled=not ppt_path):
            if open_file(ppt_path):
                st.toast("PowerPoint ouvert")
    with c2:
        st.text_input("Chemin Excel", value=str(excel_path or ""), disabled=True)
        if st.button("📂 Ouvrir Excel", use_container_width=True, disabled=not excel_path):
            if open_file(excel_path):
                st.toast("Excel ouvert")

    st.divider()

    st.subheader("🧱 Tables demandées (colonnes minimales par gabarit)")

    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        usages = ts.list_gabarit_usages(template_id)

    if not usages:
        st.caption("Aucune table demandée pour l'instant.")
    else:
        rows = []
        # Pour chaque usage, calculer les tables requises + colonnes
        for u in usages:
            gname = u.get("gabarit_name")
            gver  = u.get("gabarit_version", "v1")
            req = ts.compute_required_tables_for_usage(template_id, gname, gver)
            # usage cible Excel
            tgt = u.get("excel_target") or {}
            sheet, table = tgt.get("sheet", ""), tgt.get("table", "")
            for (nm, ver), cols in req.items():
                rows.append({
                    "Gabarit": f"{nm} (v{ver})",
                    "Vers Excel": f"{sheet}/{table}",
                    "Colonnes requises": ", ".join(cols) if cols else "—",
                })

        import pandas as pd
        df_req = pd.DataFrame(rows)
        st.dataframe(df_req, use_container_width=True, hide_index=True)

    st.markdown("---")

    st.subheader("📄 Tables Excel créées")
    if usages:
        unique_tbls = sorted({((u.get("excel_target") or {}).get("sheet",""),
                            (u.get("excel_target") or {}).get("table","")) for u in usages})
        if unique_tbls:
            df_tbls = pd.DataFrame([{"Feuille": s, "Table": t} for (s, t) in unique_tbls])
            st.dataframe(df_tbls, use_container_width=True, hide_index=True)
        else:
            st.caption("Aucune table Excel définie.")


    st.divider()

    st.subheader("📈 KPIs de génération")
    k1, k2 = st.columns(2)
    with k1:
        st.metric("Total exécutions", stats.get("total_executions", 0))
        st.metric("Taux de succès", f"{stats.get('success_rate', 0)}%")
    with k2:
        st.metric("Durée moyenne", f"{stats.get('avg_execution_time_seconds', 0)}s")
        st.metric("Échecs", stats.get("failed_executions", 0))

# ---------------------------------------------------------------------
# COLONNE DROITE : HISTORIQUE DES TÉLÉCHARGEMENTS
# ---------------------------------------------------------------------
with col_right:
    st.subheader("⬇️ Historique des téléchargements")

    # Modèle ExecutionJob (historique)
    from backend.database.models import ExecutionJob  # pylint: disable=import-error

    # Récup données fraîches
    with DatabaseService.get_session() as db:
        recent = (
            db.query(ExecutionJob)
              .filter_by(template_id=template_id)
              .order_by(ExecutionJob.created_at.desc())
              .limit(20)
              .all()
        )
        jobs = [{
            "id": j.id,
            "date": j.created_at,
            "status": j.status,
            "duration": j.execution_time_seconds,
            "excel_path": j.output_excel_path,
            "ppt_path": j.output_ppt_path,
            "parameters": j.parameters,
            "error": j.error_message,
        } for j in recent]

    if not jobs:
        st.caption("Aucun téléchargement pour le moment.")
    else:
        with st.container(height=620):
            for job in jobs:
                # Ligne principale
                cst, cdate, cxl, cppt, cact = st.columns([1, 3, 2, 2, 1])

                with cst:
                    st.markdown("🔄" if job["status"] == "running"
                                else ("✅" if job["status"] == "completed" else "❌"))
                with cdate:
                    d = fmt_paris(job["date"])
                    dur = f" · {job['duration']:.1f}s" if job["duration"] else ""
                    st.markdown(f"**{d}**{dur}")

                with cxl:
                    ok = job["excel_path"] and Path(job["excel_path"]).exists()
                    if st.button("📊 Excel", key=f"job_x_{job['id']}",
                                 use_container_width=True,
                                 disabled=(not ok or job["status"] != "completed")):
                        if open_file(job["excel_path"]):
                            st.toast("Excel ouvert")

                with cppt:
                    ok = job["ppt_path"] and Path(job["ppt_path"]).exists()
                    if st.button("📄 PPT", key=f"job_p_{job['id']}",
                                 use_container_width=True,
                                 disabled=(not ok or job["status"] != "completed")):
                        if open_file(job["ppt_path"]):
                            st.toast("PowerPoint ouvert")

                with cact:
                    if st.button("🗑️", key=f"job_del_{job['id']}", use_container_width=True,
                                 help="Supprimer cette exécution (fichiers + traces)"):
                        with DatabaseService.get_session() as db_del:
                            ok = DatabaseService.delete_job_and_files(db_del, job["id"])
                        if ok:
                            st.success("Exécution supprimée")
                        else:
                            st.warning("Exécution introuvable")
                        st.rerun()

                # Détails
                with st.expander("📋 Détails"):
                    st.caption(f"**Heure** : {fmt_paris(job['date'])}")
                    if job["duration"]:
                        st.caption(f"**Durée** : {job['duration']:.2f} s")
                    if job["parameters"]:
                        st.caption("**Paramètres** :")
                        st.json(job["parameters"])
                    if job["error"]:
                        st.error(f"**Erreur** : {job['error']}")

if st.session_state.get('show_delete_modal_detail'):
    @st.dialog("⚠️ Confirmer la suppression")
    def confirm_delete_detail():
        st.warning(f"**Vous êtes sur le point de supprimer ce template :**")
        st.markdown(f"### {tpl_name} (v{tpl_version})")
        
        if tpl_desc:
            st.caption(tpl_desc)
        
        st.divider()
        
        # Avertissement si des exécutions existent
        with DatabaseService.get_session() as db:
            from backend.database.models import ExecutionJob
            count_jobs = db.query(ExecutionJob).filter_by(template_id=template_id).count()
        
        if count_jobs > 0:
            st.error(f"⚠️ Ce template a **{count_jobs}** exécution(s) dans l'historique. Elles seront également supprimées.")
        
        st.markdown("**Cette action est irréversible.** Tapez le nom exact du template pour confirmer :")
        
        confirmation = st.text_input(
            "Nom du template",
            key="delete_confirm_detail_input",
            placeholder=tpl_name
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Annuler", use_container_width=True, key="cancel_delete_detail"):
                st.session_state.show_delete_modal_detail = False
                if 'delete_template_detail_id' in st.session_state:
                    del st.session_state.delete_template_detail_id
                st.rerun()
        
        with col2:
            if st.button("Supprimer définitivement", type="primary", use_container_width=True, key="confirm_delete_detail"):
                if confirmation != tpl_name:
                    st.error("❌ Le nom ne correspond pas")
                else:
                    try:
                        with DatabaseService.get_session() as db:
                            service = TemplateService(db)
                            service.delete_template(template_id)
                        
                        st.success(f"✅ Template '{tpl_name}' supprimé")
                        st.session_state.show_delete_modal_detail = False
                        if 'delete_template_detail_id' in st.session_state:
                            del st.session_state.delete_template_detail_id
                        if 'selected_template_detail' in st.session_state:
                            del st.session_state.selected_template_detail
                        
                        st.switch_page("pages/2_📚_Bibliotheque.py")
                    
                    except Exception as e:
                        st.error(f"❌ Erreur lors de la suppression : {e}")
    
    confirm_delete_detail()