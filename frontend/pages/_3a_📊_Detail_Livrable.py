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
        st.switch_page("pages/3_📚_Bibliotheque.py")
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

# ===== Navbar (5 boutons unifiés)
def render_template_subnav(active: str, template_id: int):
    cols = st.columns([1,1,1,1,1])
    with cols[0]:
        if st.button("← Retour bibliothèque", use_container_width=True):
            if "selected_template" in st.session_state:
                del st.session_state.selected_template
            st.switch_page("pages/3_📚_Bibliotheque.py")
    with cols[1]:
        st.button("🗂️ Détail du template", type="primary" if active=="detail" else "secondary", use_container_width=True)
    with cols[2]:
        if st.button("⚙️ Paramètres généraux", type=("primary" if active=="general" else "secondary"), use_container_width=True):
            st.session_state.selected_template = template_id
            st.switch_page("pages/_3b_➕_Form_Template.py")
    with cols[3]:
        if st.button("📑 Injection des données", type=("primary" if active=="inject" else "secondary"), use_container_width=True):
            st.session_state.selected_template = template_id
            st.switch_page("pages/_3c_📑_Tables_Template.py")
    with cols[4]:
        if st.button("🧾 Ajustement de la table", type=("primary" if active=="adjust" else "secondary"), use_container_width=True):
            st.session_state.selected_template = template_id
            st.switch_page("pages/_3d_🧾_Ajustement_Table.py")
    st.divider()


render_template_subnav("detail", template_id)

# ===== Titre + description + actions
col_title, col_actions = st.columns([4, 1])

with col_title:
    st.title(f"🗂️ Détail du template — {tpl_name} (v{tpl_version})")
    if tpl_desc:
        st.info(tpl_desc)

# ============= GESTION DU BOUTON "CRÉER LIVRABLE" =============
with col_actions:
    st.write("")
    st.write("")

    # ✅ SI GÉNÉRATION EN COURS → AFFICHER SPINNER
    if st.session_state.get('generating_report'):
        with st.spinner("⏳ Génération en cours..."):
            try:
                params_to_use = st.session_state.get('generation_params', {})
                
                with DatabaseService.get_session() as db:
                    ts_gen = TemplateService(db)
                    cfg_obj = ts_gen.load_template_config(template_id)
                
                rs = ReportService(cfg_obj)
                result = rs.generate_report(parameters=params_to_use, project_id=None)
                
                # ✅ DÉSACTIVER LE SPINNER
                st.session_state.generating_report = False
                if 'generation_params' in st.session_state:
                    del st.session_state.generation_params
                
                if result.get("success"):
                    st.success("✅ Livrable créé avec succès !")
                    st.balloons()
                    
                    with st.expander("📁 Fichiers générés", expanded=True):
                        st.code(result.get("excel_path"), language=None)
                        st.code(result.get("pptx_path"), language=None)
                    
                    st.caption(f"⏱️ Temps d'exécution : {result.get('execution_time_seconds', 0):.1f}s")
                    
                    import time
                    time.sleep(2)
                    st.rerun()
                else:
                    st.error(f"❌ Échec : {result.get('error', 'erreur inconnue')}")
            
            except Exception as e:
                st.session_state.generating_report = False
                if 'generation_params' in st.session_state:
                    del st.session_state.generation_params
                
                st.error(f"❌ Erreur lors de la génération : {e}")
                import traceback
                with st.expander("🐛 Détails techniques"):
                    st.code(traceback.format_exc())
    
    # ✅ SINON → AFFICHER LE BOUTON NORMAL
    else:
        if st.button("▶️ Créer livrable", use_container_width=True, type="primary", key="btn_build_modal"):
            st.session_state.show_generation_modal = True
            st.rerun()

    # Bouton Supprimer (inchangé)
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
                    st.markdown("📄" if job["status"] == "running"
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


# ============= MODALE DE GÉNÉRATION AVEC PARAMÈTRES =============
if st.session_state.get('show_generation_modal'):
    
    @st.dialog("🚀 Générer le livrable", width="large")
    def generation_modal():
        # ✅ IMPORTS EN DÉBUT DE FONCTION
        from backend.services.parameter_service import ParameterService
        
        st.markdown(f"### {tpl_name} (v{tpl_version})")
        
        if tpl_desc:
            st.caption(tpl_desc)
        
        st.divider()
        
        # ✅ CHARGEMENT À CHAQUE OUVERTURE
        with DatabaseService.get_session() as db:
            ts_local = TemplateService(db)
            cfg_obj = ts_local.load_template_config(template_id)
            params_config = cfg_obj.parameters
        
        if not params_config:
            st.info("ℹ️ Ce template n'a pas de paramètres configurés.")
            st.markdown("Le livrable sera généré avec les données par défaut.")
        else:
            # 🔄 BOUTON REFRESH (VERSION CORRIGÉE)
            col_title, col_refresh = st.columns([4, 1])
            with col_title:
                st.markdown("### ⚙️ Paramètres de génération")
            with col_refresh:
                if st.button("🔄 Recharger", key="refresh_params", help="Recharger les options depuis la config"):
                    # ✅ Vider le cache session pour forcer le rechargement
                    memory_key = f"last_params_{template_id}"
                    if memory_key in st.session_state:
                        del st.session_state[memory_key]
                    st.rerun()
            
            # ✅ CLÉ DE MÉMORISATION PAR TEMPLATE
            memory_key = f"last_params_{template_id}"
            
            # ✅ INITIALISER AVEC LES DERNIÈRES VALEURS OU LES DEFAULTS
            if memory_key not in st.session_state:
                st.session_state[memory_key] = {
                    p.name: ParameterService.get_default_value(p) 
                    for p in params_config
                }
            
            user_values = {}
            
            # 🐛 DEBUG EXPANDER (à supprimer après validation)
            with st.expander("🔍 Debug : Cache des paramètres", expanded=False):
                for p in params_config:
                    st.write(f"**{p.name}** ({p.type})")
                    cache_info = {
                        "options_mode": p.options_mode,
                        "has_cache": bool(p.options_cache),
                        "has_manual": bool(p.options_manual),
                        "has_source": bool(p.options_source)
                    }
                    
                    if p.options_cache:
                        cache_info["cache_size"] = len(p.options_cache.get('values', []))
                        cache_info["cache_preview"] = str(p.options_cache.get('values', [])[:3])
                    
                    st.json(cache_info)
            
            # ✅ CONTAINER SCROLLABLE
            with st.container(height=400):
                for param in params_config:
                    param_name = param.name
                    param_type = param.type
                    last_value = st.session_state[memory_key].get(param_name)
                    
                    st.markdown(f"**{param.name}**")
                    if param.description:
                        st.caption(param.description)
                    
                    # TYPE LISTE/SELECT
                    if param_type in ("liste", "select", "string") and param.options_mode != "none":
                        options = []
                        source_label = ""
                        
                        # 1️⃣ PRIORITÉ AU CACHE
                        if param.options_cache and isinstance(param.options_cache, dict):
                            cached_values = param.options_cache.get('values', [])
                            if cached_values:
                                options = list(cached_values)
                                source_label = f"💾 {len(options)} options en cache"
                        
                        # 2️⃣ SINON OPTIONS MANUELLES
                        if not options and param.options_manual and len(param.options_manual) > 0:
                            options = list(param.options_manual)
                            source_label = f"✍️ {len(options)} options manuelles"
                        
                        # 3️⃣ SINON RÉSOLUTION DYNAMIQUE
                        if not options and param.options_source:
                            try:
                                options = ParameterService._resolve_options_from_gabarit(param)
                                source_label = f"🔄 {len(options)} options calculées"
                            except Exception as e:
                                st.warning(f"⚠️ Erreur résolution : {e}")
                                source_label = "❌ Erreur de résolution"
                        
                        # Afficher la source
                        if source_label:
                            st.caption(source_label)
                        
                        # WIDGET SELON DISPONIBILITÉ
                        if not options or len(options) == 0:
                            st.warning("⚠️ Aucune option disponible - saisie libre activée")
                            user_values[param_name] = st.text_input(
                                f"{param_name} (saisie libre)",
                                value=str(last_value) if last_value else "",
                                key=f"modal_param_{param_name}",
                                label_visibility="collapsed"
                            )
                        else:
                            # ✅ SELECTBOX AVEC OPTIONS
                            try:
                                if last_value and str(last_value) in options:
                                    default_idx = options.index(str(last_value))
                                else:
                                    default_idx = 0
                            except (ValueError, TypeError):
                                default_idx = 0
                            
                            user_values[param_name] = st.selectbox(
                                param_name,
                                options=options,
                                index=default_idx,
                                key=f"modal_param_{param_name}",
                                label_visibility="collapsed",
                                help=f"🔍 Tapez pour rechercher parmi {len(options)} option(s)"
                            )
                    
                    # TYPE INTEGER
                    elif param_type == "integer":
                        user_values[param_name] = st.number_input(
                            param_name,
                            value=int(last_value) if last_value is not None else 0,
                            key=f"modal_param_{param_name}",
                            label_visibility="collapsed"
                        )
                    
                    # TYPE DATE
                    elif param_type == "date":
                        from datetime import datetime, date
                        
                        if isinstance(last_value, str):
                            try:
                                last_value = datetime.fromisoformat(last_value).date()
                            except Exception:
                                last_value = date.today()
                        elif not isinstance(last_value, date):
                            last_value = date.today()
                        
                        user_values[param_name] = st.date_input(
                            param_name,
                            value=last_value,
                            key=f"modal_param_{param_name}",
                            label_visibility="collapsed"
                        ).isoformat()
                    
                    # TYPE STRING
                    else:
                        user_values[param_name] = st.text_input(
                            param_name,
                            value=str(last_value) if last_value else "",
                            key=f"modal_param_{param_name}",
                            label_visibility="collapsed"
                        )
                    
                    st.markdown("")
        
        st.divider()
        

        # BOUTONS D'ACTION
        col1, col2 = st.columns(2)

        with col1:
            if st.button("❌ Annuler", use_container_width=True, key="cancel_generation"):
                st.session_state.show_generation_modal = False
                st.rerun()

        with col2:
            if st.button("🚀 Générer", type="primary", use_container_width=True, key="confirm_generation"):
                if params_config:
                    st.session_state[memory_key] = dict(user_values)
                
                # ✅ FERMER LA MODALE IMMÉDIATEMENT
                st.session_state.show_generation_modal = False
                
                # ✅ ACTIVER LE SPINNER SUR LE BOUTON PRINCIPAL
                st.session_state.generating_report = True
                st.session_state.generation_params = user_values if params_config else {}
                
                st.rerun()
    
    generation_modal()
    
# ============= MODALE SUPPRESSION =============
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
                        
                        st.switch_page("pages/3_📚_Bibliotheque.py")
                    
                    except Exception as e:
                        st.error(f"❌ Erreur lors de la suppression : {e}")
    
    confirm_delete_detail()