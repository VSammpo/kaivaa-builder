# frontend/pages/_1a_🗂️_Hub_Projet.py
import streamlit as st
from pathlib import Path
import sys
import subprocess
import platform
from loguru import logger

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService
from backend.services.template_service import TemplateService
from backend.database.models import ExecutionJob

st.set_page_config(page_title="Hub Projet", page_icon="🗂️", layout="wide")

# ===== NAVBAR PROJET (sans Configuration) =====
def render_project_subnav(active: str):
    cols = st.columns([1, 1, 1])
    
    with cols[0]:
        if st.button("← Projets", use_container_width=True):
            if 'selected_project_id' in st.session_state:
                del st.session_state.selected_project_id
            st.switch_page("pages/1_📁_Projets.py")
    
    with cols[1]:
        st.button("🗂️ Hub", 
                 type="primary" if active == "hub" else "secondary",
                 use_container_width=True)
    
    with cols[2]:
        if st.button("💾 Données", 
                    type="primary" if active == "data" else "secondary",
                    use_container_width=True):
            st.switch_page("pages/_1b_💾_Data_Projet.py")
    
    st.divider()

# ===== GUARD =====
if 'selected_project_id' not in st.session_state or not st.session_state.selected_project_id:
    st.error("Aucun projet sélectionné")
    if st.button("← Retour aux projets"):
        st.switch_page("pages/1_📁_Projets.py")
    st.stop()

project_id = st.session_state.selected_project_id

# ===== CHARGEMENT =====
DatabaseService.initialize()
with DatabaseService.get_session() as db:
    ps = ProjectService(db)
    ts = TemplateService(db)
    
    try:
        proj = ps.load_project(project_id)
    except FileNotFoundError:
        st.error(f"Projet introuvable : {project_id}")
        if st.button("← Retour aux projets"):
            st.switch_page("pages/1_📁_Projets.py")
        st.stop()

render_project_subnav("hub")

# ===== EN-TÊTE =====
col_title, col_add, col_settings = st.columns([3, 1, 1])

with col_title:
    st.title(f"🗂️ {proj.get('name', '(sans nom)')}")
    if proj.get('client_name'):
        st.caption(f"Client : {proj['client_name']}")

with col_add:
    st.write("")
    if st.button("➕ Ajouter livrable", use_container_width=True, type="primary"):
        st.session_state.show_add_deliverable_modal = True
        st.rerun()

with col_settings:
    st.write("")
    if st.button("⚙️ Paramètres", use_container_width=True):
        st.session_state.show_project_settings = True
        st.rerun()

if proj.get('description'):
    st.info(proj['description'])

st.divider()

# ===== STATISTIQUES GLOBALES =====
deliverables = proj.get("deliverables", [])
data_sources = proj.get("data_sources", [])

col_m1, col_m2, col_m3, col_m4 = st.columns(4)

with col_m1:
    st.metric("Livrables", len(deliverables))

with col_m2:
    functional_count = sum(1 for d in deliverables if d.get("is_functional", False))
    st.metric("Fonctionnels", f"{functional_count}/{len(deliverables)}")

with col_m3:
    client_sources = sum(1 for s in data_sources if s.get("source_type") == "client")
    st.metric("Sources client", client_sources)

with col_m4:
    last_gen_time = None
    
    if deliverables:
        template_ids_for_query = [d['template_id'] for d in deliverables]
        
        with DatabaseService.get_session() as db_query:
            try:
                # ✅ CORRECTION : Filtrer sur project_id
                last_job = db_query.query(ExecutionJob).filter(
                    ExecutionJob.template_id.in_(template_ids_for_query),
                    ExecutionJob.project_id == project_id
                ).order_by(ExecutionJob.created_at.desc()).first()
                
                if last_job:
                    last_gen_time = last_job.created_at
            except Exception as e:
                logger.warning(f"Erreur requête historique : {e}")
                last_gen_time = None
    
    if last_gen_time:
        from datetime import datetime
        from zoneinfo import ZoneInfo
        dt = last_gen_time.astimezone(ZoneInfo("Europe/Paris"))
        st.metric("Dernière génération", dt.strftime("%d/%m %H:%M"))
    else:
        st.metric("Dernière génération", "—")

st.divider()

# ===== LISTE DES LIVRABLES =====
st.subheader("📦 Livrables du projet")

if not deliverables:
    st.info("Aucun livrable configuré. Ajoutez-en un pour démarrer.")
else:
    cols_per_row = 2
    
    for i in range(0, len(deliverables), cols_per_row):
        cols = st.columns(cols_per_row)
        
        for j, col in enumerate(cols):
            idx = i + j
            if idx < len(deliverables):
                deliv = deliverables[idx]
                
                with col:
                    with st.container(border=True):
                        col_name, col_badge = st.columns([3, 1])
                        
                        with col_name:
                            st.markdown(f"### {deliv.get('template_name', 'Template inconnu')}")
                            st.caption(f"v{deliv.get('template_version', '1.0')}")
                        
                        with col_badge:
                            if deliv.get('is_functional', False):
                                st.markdown("🟢 **Prêt**")
                            else:
                                st.markdown("🔴 **Incomplet**")
                        
                        completion = deliv.get('completion_rate', 0.0)
                        st.progress(completion / 100.0)
                        st.caption(f"Complétude : {completion}%")
                        
                        st.markdown("")
                        ds_status = deliv.get('data_sources_status', {})
                        col_s1, col_s2 = st.columns(2)
                        
                        with col_s1:
                            st.caption(f"📁 Client : {ds_status.get('client', 0)}")
                        
                        with col_s2:
                            st.caption(f"🔄 Défaut : {ds_status.get('default', 0)}")
                        
                        if deliv.get('last_generated_at'):
                            from datetime import datetime
                            try:
                                dt = datetime.fromisoformat(deliv['last_generated_at'])
                                st.caption(f"🕒 Généré : {dt.strftime('%d/%m %H:%M')}")
                            except:
                                pass
                        
                        st.markdown("")
                        
                        col_btn1, col_btn2, col_btn3 = st.columns(3)
                        
                        with col_btn1:
                            if st.button("▶️ Générer", 
                                       key=f"gen_{deliv['template_id']}",
                                       use_container_width=True,
                                       disabled=not deliv.get('is_functional', False)):
                                st.session_state.generate_deliverable = deliv['template_id']
                                st.rerun()
                        
                        with col_btn2:
                            if st.button("⚙️ Config", 
                                       key=f"cfg_{deliv['template_id']}",
                                       use_container_width=True):
                                st.session_state.selected_deliverable_id = deliv['template_id']
                                st.switch_page("pages/_1c_⚙️_Config_Deliverable.py")
                        
                        with col_btn3:
                            if st.button("🗑️", 
                                       key=f"del_{deliv['template_id']}",
                                       use_container_width=True,
                                       help="Retirer ce livrable"):
                                st.session_state.remove_deliverable_id = deliv['template_id']
                                st.rerun()

st.markdown("---")

# ===== HISTORIQUE RÉCENT =====
st.subheader("📜 Dernières générations")

if not deliverables:
    st.info("Aucun livrable configuré. Ajoutez-en un pour voir l'historique des générations.")
else:
    template_ids = [d['template_id'] for d in deliverables]
    
    jobs_data = []
    
    with DatabaseService.get_session() as db_hist:
        try:
            # ✅ CORRECTION : Filtrer sur project_id
            recent_jobs = db_hist.query(ExecutionJob).filter(
                ExecutionJob.template_id.in_(template_ids),
                ExecutionJob.project_id == project_id
            ).order_by(ExecutionJob.created_at.desc()).limit(5).all()
            
            for job in recent_jobs:
                jobs_data.append({
                    'id': job.id,
                    'template_id': job.template_id,
                    'status': job.status,
                    'created_at': job.created_at,
                    'output_excel_path': job.output_excel_path,
                    'output_ppt_path': job.output_ppt_path
                })
        except Exception as e:
            logger.warning(f"Erreur chargement historique : {e}")
            jobs_data = []
    
    if jobs_data:
        for job in jobs_data:
            with st.container(border=True):
                col_info, col_status, col_actions = st.columns([3, 1, 2])
                
                with col_info:
                    tpl_name = next((d['template_name'] for d in deliverables 
                                   if d['template_id'] == job['template_id']), "Unknown")
                    st.markdown(f"**{tpl_name}**")
                    
                    from datetime import datetime
                    from zoneinfo import ZoneInfo
                    dt = job['created_at'].astimezone(ZoneInfo("Europe/Paris"))
                    st.caption(f"🕒 {dt.strftime('%d/%m/%Y %H:%M')}")
                
                with col_status:
                    if job['status'] == "completed":
                        st.markdown("✅ **OK**")
                    elif job['status'] == "failed":
                        st.markdown("❌ **Erreur**")
                    else:
                        st.markdown("⏳ **En cours**")
                
                with col_actions:
                    if job['status'] == "completed":
                        col_a1, col_a2 = st.columns(2)
                        
                        def open_file(filepath: str):
                            try:
                                abspath = str(Path(filepath).resolve())
                                if platform.system() == "Windows":
                                    subprocess.run(["cmd", "/c", "start", "", abspath], check=True)
                                elif platform.system() == "Darwin":
                                    subprocess.run(["open", abspath], check=True)
                                else:
                                    subprocess.run(["xdg-open", abspath], check=True)
                            except Exception as e:
                                st.error(f"Erreur : {e}")
                        
                        with col_a1:
                            if st.button("📊 Excel",
                                        key=f"open_xls_{job['id']}",
                                        use_container_width=True,
                                        disabled=(not bool(job.get('output_excel_path')) or job['status'] != "completed")):
                                open_file(job['output_excel_path'])

                        with col_a2:
                            if st.button("📄 PPT",
                                        key=f"open_ppt_{job['id']}",
                                        use_container_width=True,
                                        disabled=(not bool(job.get('output_ppt_path')) or job['status'] != "completed")):
                                open_file(job['output_ppt_path'])


    else:
        st.info("Aucune génération pour le moment. Configurez les données puis générez un livrable.")

# ===== MODALE AJOUT LIVRABLE =====
if st.session_state.get('show_add_deliverable_modal'):
    @st.dialog("➕ Ajouter un livrable", width="large")
    def add_deliverable_modal():
        st.markdown("### Sélectionnez un template")
        
        with DatabaseService.get_session() as db_modal:
            ts_local = TemplateService(db_modal)
            all_templates = ts_local.list_templates(active_only=True)
            
            templates_data = []
            for t in all_templates:
                templates_data.append({
                    'id': t.id,
                    'name': t.name,
                    'version': t.version,
                    'description': t.description
                })
        
        existing_ids = [d['template_id'] for d in deliverables]
        available = [t for t in templates_data if t['id'] not in existing_ids]
        
        if not available:
            st.warning("Tous les templates sont déjà ajoutés à ce projet")
            if st.button("← Retour"):
                st.session_state.show_add_deliverable_modal = False
                st.rerun()
            return
        
        for tpl in available:
            with st.container(border=True):
                col_info, col_btn = st.columns([4, 1])
                
                with col_info:
                    st.markdown(f"**{tpl['name']}**")
                    st.caption(f"v{tpl['version']}")
                    if tpl['description']:
                        desc = tpl['description'][:100] + "..." if len(tpl['description']) > 100 else tpl['description']
                        st.caption(desc)
                
                with col_btn:
                    if st.button("➕ Ajouter", key=f"add_{tpl['id']}", use_container_width=True):
                        try:
                            with DatabaseService.get_session() as db_add:
                                ps_local = ProjectService(db_add)
                                ps_local.add_deliverable(project_id, tpl['id'])
                                ps_local.update_deliverable_status(project_id, tpl['id'])
                            
                            st.success(f"✅ {tpl['name']} ajouté !")
                            st.session_state.show_add_deliverable_modal = False
                            
                            import time
                            time.sleep(1)
                            st.rerun()
                        
                        except Exception as e:
                            st.error(f"❌ Erreur : {e}")
                            import traceback
                            with st.expander("🔍 Détails"):
                                st.code(traceback.format_exc())
        
        st.divider()
        
        if st.button("❌ Annuler", use_container_width=True):
            st.session_state.show_add_deliverable_modal = False
            st.rerun()
    
    add_deliverable_modal()

# ===== MODALE PARAMÈTRES PROJET =====
if st.session_state.get('show_project_settings'):
    @st.dialog("⚙️ Paramètres du projet", width="large")
    def project_settings_modal():
        st.markdown("### Informations générales")
        
        name = st.text_input("Nom du projet", value=proj.get('name', ''))
        client_name = st.text_input("Nom du client", value=proj.get('client_name', ''))
        description = st.text_area("Description", value=proj.get('description', ''), height=100)
        
        st.divider()
        
        col_save, col_cancel = st.columns(2)
        
        with col_cancel:
            if st.button("❌ Annuler", use_container_width=True):
                st.session_state.show_project_settings = False
                st.rerun()
        
        with col_save:
            if st.button("💾 Enregistrer", type="primary", use_container_width=True):
                try:
                    proj['name'] = name
                    proj['client_name'] = client_name
                    proj['description'] = description
                    
                    with DatabaseService.get_session() as db_save:
                        ps_save = ProjectService(db_save)
                        ps_save.save_project(proj)
                    
                    st.success("✅ Paramètres enregistrés")
                    st.session_state.show_project_settings = False
                    
                    import time
                    time.sleep(1)
                    st.rerun()
                
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
    
    project_settings_modal()

# ===== MODALE GÉNÉRATION =====
if st.session_state.get('generate_deliverable'):
    @st.dialog("▶️ Générer un livrable", width="large")
    def generate_modal():
        # ✅ CORRECTION : Vérifier que la clé existe ET est valide
        template_id = st.session_state.get('generate_deliverable')
        
        if not template_id:
            st.error("Erreur : aucun livrable sélectionné")
            if st.button("Fermer"):
                if 'generate_deliverable' in st.session_state:
                    del st.session_state.generate_deliverable
                st.rerun()
            return
        
        # Charger les infos
        deliverable = next((d for d in deliverables if d['template_id'] == template_id), None)
        if not deliverable:
            st.error("Livrable introuvable")
            if 'generate_deliverable' in st.session_state:
                del st.session_state.generate_deliverable
            st.rerun()
            return
        
        st.markdown(f"### {deliverable['template_name']} v{deliverable['template_version']}")
        
        # Vérifier la complétude
        if not deliverable.get('is_functional'):
            st.error("⚠️ Ce livrable n'est pas prêt à être généré")
            st.markdown("**Problèmes détectés :**")
            
            status = None
            with DatabaseService.get_session() as db_check:
                ps_check = ProjectService(db_check)
                status = ps_check.compute_deliverable_status(project_id, template_id)
            
            if status and status.get('missing_gabarits'):
                st.markdown("**Gabarits manquants :**")
                for g in status['missing_gabarits']:
                    st.markdown(f"- {g}")
            
            st.divider()
            
            col_data, col_cancel = st.columns(2)
            
            with col_data:
                if st.button("💾 Configurer les données", use_container_width=True, type="primary"):
                    if 'generate_deliverable' in st.session_state:
                        del st.session_state.generate_deliverable
                    st.switch_page("pages/_1b_💾_Data_Projet.py")
            
            with col_cancel:
                if st.button("❌ Annuler", use_container_width=True):
                    if 'generate_deliverable' in st.session_state:
                        del st.session_state.generate_deliverable
                    st.rerun()
            
            return
        
        # Charger les paramètres du template
        with DatabaseService.get_session() as db_params:
            ts_params = TemplateService(db_params)
            tpl_config = ts_params.load_template_config(template_id)
        
        params = tpl_config.parameters if tpl_config else []
        custom_params = deliverable.get('custom_parameters', {})
        
        st.markdown("---")
        st.markdown("### 🎛️ Paramètres de génération")
        
        if not params:
            st.info("Ce template n'a pas de paramètres configurés")
        
        # Stocker les valeurs sélectionnées
        if 'generation_params' not in st.session_state:
            st.session_state.generation_params = {}
        
        for param in params:
            param_name = param.name
            
            with st.container(border=True):
                st.markdown(f"**{param_name}**")
                if param.description:
                    st.caption(param.description)
                
                # Récupérer la valeur par défaut (custom > template)
                default_value = custom_params.get(param_name, {}).get('default')
                if default_value is None:
                    from backend.services.parameter_service import ParameterService
                    default_value = ParameterService.get_default_value(param)
                
                # Widget selon le type
                if param.type in ("string", "liste", "select"):
                    # Récupérer les options
                    options = []
                    
                    if param.options_mode == "manual" and param.options_manual:
                        options = list(param.options_manual)
                    
                    elif param.options_mode == "from_column" and param.options_source:
                        # Résoudre depuis les données du projet
                        try:
                            col_name = param.options_source.get("column")
                            gab_name = param.options_source.get("gabarit")
                            gab_ver = param.options_source.get("version", "v1")
                            
                            with DatabaseService.get_session() as db_opts:
                                ps_opts = ProjectService(db_opts)
                                data_source = ps_opts.get_data_source(project_id, gab_name, gab_ver)
                            
                            if data_source:
                                from backend.services.dataset_service import _load_dataframe_from_source
                                
                                df = _load_dataframe_from_source(data_source.get("source_config"))
                                
                                if df is not None and col_name in df.columns:
                                    options = sorted(list(df[col_name].dropna().astype(str).unique()))[:500]
                        
                        except Exception as e:
                            logger.warning(f"Impossible de charger les options : {e}")
                    
                    # Cache si disponible
                    if not options and hasattr(param, 'options_cache') and param.options_cache:
                        cache_values = param.options_cache.get('values', [])
                        if cache_values:
                            options = list(cache_values)
                    
                    # Widget
                    if options:
                        try:
                            default_idx = options.index(str(default_value)) if str(default_value) in options else 0
                        except (ValueError, TypeError):
                            default_idx = 0
                        
                        value = st.selectbox(
                            "Valeur",
                            options=options,
                            index=default_idx,
                            key=f"gen_param_{param_name}",
                            label_visibility="collapsed"
                        )
                    else:
                        value = st.text_input(
                            "Valeur",
                            value=str(default_value) if default_value else "",
                            key=f"gen_param_{param_name}",
                            label_visibility="collapsed"
                        )
                
                elif param.type == "integer":
                    value = st.number_input(
                        "Valeur",
                        value=int(default_value) if default_value is not None else 0,
                        key=f"gen_param_{param_name}",
                        label_visibility="collapsed"
                    )
                
                elif param.type == "date":
                    from datetime import datetime, date
                    
                    if isinstance(default_value, str):
                        try:
                            default_value = datetime.fromisoformat(default_value).date()
                        except:
                            default_value = date.today()
                    elif not isinstance(default_value, date):
                        default_value = date.today()
                    
                    value = st.date_input(
                        "Valeur",
                        value=default_value,
                        key=f"gen_param_{param_name}",
                        label_visibility="collapsed"
                    ).isoformat()
                
                else:
                    value = st.text_input(
                        "Valeur",
                        value=str(default_value) if default_value else "",
                        key=f"gen_param_{param_name}",
                        label_visibility="collapsed"
                    )
                
                # Stocker la valeur
                st.session_state.generation_params[param_name] = value
        
        st.markdown("---")
        
        # Options de génération
        col_opt1, col_opt2 = st.columns(2)
        
        with col_opt1:
            generate_excel = st.checkbox("📊 Générer Excel", value=True)
        
        with col_opt2:
            generate_ppt = st.checkbox("📄 Générer PowerPoint", value=True)
        
        if not generate_excel and not generate_ppt:
            st.warning("⚠️ Sélectionnez au moins un format de sortie")
        
        st.divider()
        
        # Actions
        col_gen, col_cancel = st.columns(2)
        
        with col_cancel:
            if st.button("❌ Annuler", use_container_width=True):
                if 'generation_params' in st.session_state:
                    del st.session_state.generation_params
                if 'generate_deliverable' in st.session_state:
                    del st.session_state.generate_deliverable
                st.rerun()
        
        with col_gen:
            if st.button("▶️ Lancer la génération", 
                        type="primary", 
                        use_container_width=True,
                        disabled=not (generate_excel or generate_ppt)):
                try:
                    # Préparer les paramètres
                    generation_params = st.session_state.generation_params
                    
                    # Lancer la génération
                    with st.spinner("🔄 Génération en cours..."):
                        from backend.services.generation_service import GenerationService
                        
                        with DatabaseService.get_session() as db_gen:
                            gen_service = GenerationService(db_gen)
                            
                            result = gen_service.generate_deliverable(
                                project_id=project_id,
                                template_id=template_id,
                                parameters=generation_params,
                                generate_excel=generate_excel,
                                generate_ppt=generate_ppt
                            )
                    
                    # Succès
                    st.success("✅ Génération terminée avec succès !")
                    
                    # Afficher les chemins
                    if result.get('excel_path'):
                        st.markdown(f"📊 **Excel** : `{result['excel_path']}`")
                    
                    if result.get('ppt_path'):
                        st.markdown(f"📄 **PowerPoint** : `{result['ppt_path']}`")
                    
                    # Boutons d'ouverture
                    st.markdown("---")
                    
                    col_open1, col_open2, col_close = st.columns(3)
                    
                    def open_file(filepath: str):
                        try:
                            abspath = str(Path(filepath).resolve())
                            if platform.system() == "Windows":
                                subprocess.run(["cmd", "/c", "start", "", abspath], check=True)
                            elif platform.system() == "Darwin":
                                subprocess.run(["open", abspath], check=True)
                            else:
                                subprocess.run(["xdg-open", abspath], check=True)
                        except Exception as e:
                            st.error(f"Erreur : {e}")
                    
                    with col_open1:
                        if st.button("📂 Ouvrir Excel", use_container_width=True,
                                    disabled=not bool(result.get('excel_path'))):
                            open_file(result['excel_path'])

                    with col_open2:
                        if st.button("📂 Ouvrir PPT", use_container_width=True,
                                    disabled=not bool(result.get('ppt_path'))):
                            open_file(result['ppt_path'])


                    
                    with col_close:
                        if st.button("✓ Fermer", use_container_width=True):
                            if 'generation_params' in st.session_state:
                                del st.session_state.generation_params
                            if 'generate_deliverable' in st.session_state:
                                del st.session_state.generate_deliverable
                            st.rerun()
                
                except Exception as e:
                    st.error(f"❌ Erreur lors de la génération : {e}")
                    
                    import traceback
                    with st.expander("🔍 Détails de l'erreur"):
                        st.code(traceback.format_exc())
                    
                    st.divider()
                    
                    if st.button("← Retour", use_container_width=True):
                        if 'generation_params' in st.session_state:
                            del st.session_state.generation_params
                        if 'generate_deliverable' in st.session_state:
                            del st.session_state.generate_deliverable
                        st.rerun()
    
    generate_modal()
    
# ===== ACTION SUPPRESSION LIVRABLE =====
if 'remove_deliverable_id' in st.session_state:
    tid = st.session_state.remove_deliverable_id
    
    try:
        with DatabaseService.get_session() as db_remove:
            ps_remove = ProjectService(db_remove)
            ps_remove.remove_deliverable(project_id, tid)
        
        st.success("Livrable retiré")
        del st.session_state.remove_deliverable_id
        st.rerun()
    
    except Exception as e:
        st.error(f"Erreur : {e}")
        del st.session_state.remove_deliverable_id
