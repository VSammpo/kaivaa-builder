"""
Page de détail d'un livrable - Version optimisée
"""

import streamlit as st
from pathlib import Path
import sys
import subprocess
import platform
from datetime import datetime
import pandas as pd

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.services.dataset_service import load_csv, prepare_for_usage
from backend.services.excel_injection_service import inject_dataframe
from backend.services.gabarit_registry import list_gabarits, get_gabarit
from datetime import datetime
from zoneinfo import ZoneInfo

def fmt_paris(ts) -> str:
    """
    Affiche un datetime en Europe/Paris.
    - si tz-aware: convertit en Europe/Paris
    - si naïf: le considère déjà comme heure locale et le formate
    - si str ISO: essaie de parser
    """
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
        # sinon on suppose déjà local
        return ts.strftime("%d/%m/%Y %H:%M")
    return str(ts)

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
    
    def render_gabarits_section(template_id: int):
        st.subheader("🧱 Gabarits rattachés au livrable")

        DatabaseService.initialize()
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)

            # Liste des usages déjà attachés
            usages = ts.list_gabarit_usages(template_id)

            if usages:
                df = pd.DataFrame([{
                    "gabarit": f'{u.get("gabarit_name")} (v{u.get("gabarit_version")})',
                    "sheet": u.get("excel_target", {}).get("sheet", ""),
                    "table": u.get("excel_target", {}).get("table", ""),
                    "n_cols_enabled": len(u.get("columns_enabled", []))
                } for u in usages])
                st.dataframe(df, use_container_width=True, hide_index=True)
            else:
                st.caption("Aucun gabarit attaché pour l’instant.")

            st.divider()

            # Sélection d'un gabarit global
            gab_list = list_gabarits()
            if not gab_list:
                st.info("Crée d’abord des gabarits dans la page « 0_🧱_Gabarits_de_table ».")  # page que tu as déjà
                return

            labels = [f"{g.name} (v{g.version})" for g in gab_list]
            choice = st.selectbox("Choisir un gabarit", labels, key="gab_select")
            g = gab_list[labels.index(choice)]

            all_cols = [c.name for c in g.columns]

            # Valeurs existantes si déjà attaché
            existing = ts.get_gabarit_usage(template_id, g.name, g.version)
            default_enabled = existing.get("columns_enabled", []) if existing else []
            default_sheet = existing.get("excel_target", {}).get("sheet", "D001") if existing else "D001"
            default_table = existing.get("excel_target", {}).get("table", "") if existing else ""

            columns_enabled = st.multiselect(
                "Colonnes utilisées par CE livrable",
                options=all_cols,
                default=default_enabled,
                help="Coche uniquement les colonnes nécessaires à ce livrable."
            )

            c1, c2 = st.columns(2)
            with c1:
                sheet = st.text_input("Feuille Excel cible", value=default_sheet)
            with c2:
                table = st.text_input("Nom du tableau Excel (ListObject)", value=default_table)

            c3, c4 = st.columns(2)
            with c3:
                if st.button("💾 Enregistrer / Mettre à jour", key="gab_save"):
                    ts.upsert_gabarit_usage(
                        template_id=template_id,
                        gabarit_name=g.name,
                        gabarit_version=g.version,
                        columns_enabled=columns_enabled,
                        excel_sheet=sheet,
                        excel_table=table,
                    )
                    st.success(f"Gabarit {g.name} v{g.version} attaché au livrable.")
                    st.rerun()
            with c4:
                if existing and st.button("🗑️ Détacher ce gabarit", type="secondary", key="gab_delete"):
                    if ts.delete_gabarit_usage(template_id, g.name, g.version):
                        st.success("Gabarit détaché.")
                        st.rerun()
                    else:
                        st.warning("Aucune suppression effectuée.")

    def render_sources_preview_section(template_id: int):
        st.subheader("🔌 Sources & Preview (MVP CSV)")

        DatabaseService.initialize()
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            usages = ts.list_gabarit_usages(template_id)
            if not usages:
                st.caption("Aucun gabarit attaché → rattache un gabarit dans la section au-dessus.")
                return

            for u in usages:
                gab_label = f'{u.get("gabarit_name")} (v{u.get("gabarit_version")})'
                with st.expander(f"Source pour {gab_label}", expanded=False):
                    source_existing = ts.get_gabarit_source(template_id, u["gabarit_name"], u["gabarit_version"]) or {}
                    c1, c2, c3 = st.columns([3,1,1])
                    with c1:
                        path = st.text_input("CSV path", value=source_existing.get("path",""), key=f"src_path_{gab_label}")
                    with c2:
                        sep = st.text_input("sep", value=source_existing.get("sep",";"), key=f"src_sep_{gab_label}")
                    with c3:
                        enc = st.text_input("encoding", value=source_existing.get("encoding","utf-8-sig"), key=f"src_enc_{gab_label}")

                    colA, colB = st.columns(2)
                    with colA:
                        if st.button("💾 Enregistrer la source", key=f"save_src_{gab_label}"):
                            ts.upsert_gabarit_source(
                                template_id,
                                u["gabarit_name"], u["gabarit_version"],
                                {"type":"csv","path":path,"sep":sep,"encoding":enc}
                            )
                            st.success("Source enregistrée.")
                            st.rerun()
                    with colB:
                        if path and st.button("👀 Preview 1000", key=f"preview_{gab_label}"):
                            try:
                                df = load_csv({"type":"csv","path":path,"sep":sep,"encoding":enc})
                                df = df.head(1000)
                                dfp = prepare_for_usage(df, u.get("columns_enabled", []))
                                st.write("Shape:", dfp.shape)
                                st.dataframe(dfp.head(30), use_container_width=True)
                            except Exception as e:
                                st.error(f"Erreur preview: {e}")


    def render_validation_injection_section(template_id: int, excel_template_path: str):
        st.subheader("✅ Validation & Injection (MVP)")

        DatabaseService.initialize()
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            usages = ts.list_gabarit_usages(template_id)
            if not usages:
                st.caption("Aucun gabarit attaché.")
                return

            if not excel_template_path:
                st.warning("Aucun Excel master trouvé pour ce livrable.")
                return

            if st.button("📥 Injecter TOUTES les tables dans l’Excel master"):
                n_ok, n_err = 0, 0
                for u in usages:
                    try:
                        src = ts.get_gabarit_source(template_id, u["gabarit_name"], u["gabarit_version"])
                        if not src:
                            st.warning(f"Pas de source pour {u['gabarit_name']} v{u['gabarit_version']}")
                            continue

                        df = load_csv(src)
                        dfp = prepare_for_usage(df, u.get("columns_enabled", []))

                        target = u.get("excel_target",{})
                        inject_dataframe(
                            excel_template_path,
                            target.get("sheet","D001"),
                            target.get("table",""),
                            dfp[u.get("columns_enabled", [])] if u.get("columns_enabled") else dfp
                        )
                        n_ok += 1
                    except Exception as e:
                        n_err += 1
                        st.error(f"[{u.get('gabarit_name')}] Injection KO: {e}")
                st.success(f"Injections terminées: OK={n_ok}  KO={n_err}")


    
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
    
    render_gabarits_section(template_id)
    render_sources_preview_section(template_id)
    render_validation_injection_section(template_id, excel_template_path=excel_path)
    
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
                    # Affichage heure locale (Europe/Paris)
                    date_str = fmt_paris(job['date'])
                    duration_str = f" - {job['duration']:.1f}s" if job['duration'] else ""
                    st.markdown(f"**{date_str}**{duration_str}")


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
                    st.caption(f"**Heure de génération :** {fmt_paris(job['date'])}")

                    
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