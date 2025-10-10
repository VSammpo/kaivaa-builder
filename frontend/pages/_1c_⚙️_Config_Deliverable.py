# frontend/pages/_1c_⚙️_Config_Deliverable.py
import streamlit as st
from pathlib import Path
import sys
import subprocess
import platform

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService
from backend.services.template_service import TemplateService
from backend.services.parameter_service import ParameterService

st.set_page_config(page_title="Config Livrable", page_icon="⚙️", layout="wide")

# ===== HELPER : OUVERTURE FICHIERS =====
def open_file(filepath: str) -> bool:
    """Ouvre un fichier avec l'application par défaut du système."""
    try:
        if not filepath:
            st.error("Chemin vide")
            return False
        
        abspath = str(Path(filepath).resolve())
        
        if not Path(abspath).exists():
            st.error(f"Fichier introuvable : {abspath}")
            return False
        
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

# ===== NAVBAR PROJET =====
def render_project_subnav(active: str):
    cols = st.columns([1, 1, 1])
    
    with cols[0]:
        if st.button("← Projets", use_container_width=True):
            if 'selected_project_id' in st.session_state:
                del st.session_state.selected_project_id
            if 'selected_deliverable_id' in st.session_state:
                del st.session_state.selected_deliverable_id
            st.switch_page("pages/1_📁_Projets.py")
    
    with cols[1]:
        if st.button("🗂️ Hub", 
                    type="primary" if active == "hub" else "secondary",
                    use_container_width=True):
            if 'selected_deliverable_id' in st.session_state:
                del st.session_state.selected_deliverable_id
            st.switch_page("pages/_1a_🗂️_Hub_Projet.py")
    
    with cols[2]:
        if st.button("💾 Données", 
                    type="primary" if active == "data" else "secondary",
                    use_container_width=True):
            if 'selected_deliverable_id' in st.session_state:
                del st.session_state.selected_deliverable_id
            st.switch_page("pages/_1b_💾_Data_Projet.py")
    
    st.divider()

# ===== GUARD PROJET =====
if 'selected_project_id' not in st.session_state or not st.session_state.selected_project_id:
    st.error("Aucun projet sélectionné")
    if st.button("← Retour aux projets"):
        st.switch_page("pages/1_📁_Projets.py")
    st.stop()

# ===== GUARD LIVRABLE =====
if 'selected_deliverable_id' not in st.session_state or not st.session_state.selected_deliverable_id:
    st.error("Aucun livrable sélectionné")
    if st.button("← Retour au Hub"):
        st.switch_page("pages/_1a_🗂️_Hub_Projet.py")
    st.stop()

project_id = st.session_state.selected_project_id
template_id = st.session_state.selected_deliverable_id

# ===== CHARGEMENT =====
DatabaseService.initialize()

# ✅ Extraire les données dans la session
with DatabaseService.get_session() as db:
    ps = ProjectService(db)
    ts = TemplateService(db)
    
    try:
        proj = ps.load_project(project_id)
        tpl_obj = ts.get_template(template_id)
        tpl_config = ts.load_template_config(template_id)
        
        # Extraire les données avant fermeture
        tpl = {
            'id': tpl_obj.id,
            'name': tpl_obj.name,
            'version': tpl_obj.version,
            'description': tpl_obj.description
        }
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        if st.button("← Retour au Hub"):
            st.switch_page("pages/_1a_🗂️_Hub_Projet.py")
        st.stop()

# Trouver le livrable dans le projet
deliverable = next((d for d in proj.get("deliverables", []) if d["template_id"] == template_id), None)
if not deliverable:
    st.error("Livrable introuvable dans le projet")
    if st.button("← Retour au Hub"):
        st.switch_page("pages/_1a_🗂️_Hub_Projet.py")
    st.stop()

render_project_subnav("config")

# ===== EN-TÊTE =====
col_title, col_status = st.columns([3, 1])

with col_title:
    st.title(f"⚙️ Configuration — {tpl['name']}")
    st.caption(f"Projet : {proj.get('name')} • v{tpl['version']}")

with col_status:
    st.write("")
    if deliverable.get("is_functional"):
        st.markdown("🟢 **Prêt**")
    else:
        st.markdown("🔴 **Incomplet**")

# ===== ONGLETS =====
tab_masters, tab_params, tab_stats = st.tabs([
    "📄 Masters",
    "🎛️ Paramètres",
    "📊 Statistiques"
])

# ==================== ONGLET 1 : MASTERS ====================
with tab_masters:
    st.subheader("📄 Fichiers masters du projet")
    st.caption("Ces fichiers sont des copies des masters du template, personnalisables pour ce projet uniquement")
    
    masters = deliverable.get("custom_masters", {})
    
    # PPT
    st.markdown("### 🎤 Présentation PowerPoint")
    
    col_ppt1, col_ppt2, col_ppt3 = st.columns([3, 1, 1])
    
    ppt_path = masters.get("ppt_path")
    
    with col_ppt1:
        if ppt_path:
            st.text_input("Chemin PPT", value=ppt_path, disabled=True, key="ppt_path_display")
        else:
            st.warning("Aucun master PPT configuré")
    
    with col_ppt2:
        if ppt_path and Path(ppt_path).exists():
            if st.button("📂 Ouvrir PPT", use_container_width=True):
                if open_file(ppt_path):
                    st.toast("PowerPoint ouvert", icon="✅")
        else:
            st.button("📂 Ouvrir PPT", use_container_width=True, disabled=True)
    
    with col_ppt3:
        if ppt_path:
            if st.button("🔄 Réinitialiser", key="reset_ppt", use_container_width=True):
                try:
                    new_path = ps.reset_master(project_id, template_id, "ppt")
                    
                    # Mettre à jour le livrable
                    for d in proj["deliverables"]:
                        if d["template_id"] == template_id:
                            d["custom_masters"]["ppt_path"] = new_path
                            break
                    
                    ps.save_project(proj)
                    
                    st.success("✅ Master PPT réinitialisé depuis le template")
                    st.rerun()
                
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
    
    st.markdown("---")
    
    # Excel
    st.markdown("### 📊 Fichier de données Excel")
    
    col_xls1, col_xls2, col_xls3 = st.columns([3, 1, 1])
    
    excel_path = masters.get("excel_path")
    
    with col_xls1:
        if excel_path:
            st.text_input("Chemin Excel", value=excel_path, disabled=True, key="excel_path_display")
        else:
            st.warning("Aucun master Excel configuré")
    
    with col_xls2:
        if excel_path and Path(excel_path).exists():
            if st.button("📂 Ouvrir Excel", use_container_width=True):
                if open_file(excel_path):
                    st.toast("Excel ouvert", icon="✅")
        else:
            st.button("📂 Ouvrir Excel", use_container_width=True, disabled=True)
    
    with col_xls3:
        if excel_path:
            if st.button("🔄 Réinitialiser", key="reset_excel", use_container_width=True):
                try:
                    new_path = ps.reset_master(project_id, template_id, "excel")
                    
                    # Mettre à jour le livrable
                    for d in proj["deliverables"]:
                        if d["template_id"] == template_id:
                            d["custom_masters"]["excel_path"] = new_path
                            break
                    
                    ps.save_project(proj)
                    
                    st.success("✅ Master Excel réinitialisé depuis le template")
                    st.rerun()
                
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
    
    st.markdown("---")
    
    st.info("""
    💡 **Comment modifier les masters ?**
    
    1. Cliquez sur "📂 Ouvrir" pour éditer le fichier
    2. Effectuez vos modifications (structure, mise en page, formules...)
    3. Enregistrez le fichier
    4. Les générations futures utiliseront cette version personnalisée
    
    ⚠️ **Attention** : Ces modifications n'affectent que ce projet. Le template original reste inchangé.
    """)

# ==================== ONGLET 2 : PARAMÈTRES ====================
with tab_params:
    st.subheader("🎛️ Paramètres du template")
    st.caption("Ajustez les valeurs par défaut et les options pour ce projet")
    
    params = tpl_config.parameters
    
    if not params:
        st.info("Ce template n'a pas de paramètres configurés")
    else:
        # Charger les paramètres personnalisés du projet
        custom_params = deliverable.get("custom_parameters", {})
        
        st.markdown("---")
        
        for param in params:
            param_name = param.name
            
            with st.container(border=True):
                col_name, col_type = st.columns([3, 1])
                
                with col_name:
                    st.markdown(f"**{param_name}**")
                    if param.description:
                        st.caption(param.description)
                
                with col_type:
                    st.caption(f"Type : {param.type}")
                
                # Récupérer la valeur personnalisée ou par défaut
                custom_value = custom_params.get(param_name, {}).get("default")
                if custom_value is None:
                    custom_value = ParameterService.get_default_value(param)
                
                # Widget selon le type
                if param.type in ("string", "liste", "select"):
                    # Récupérer les options
                    options = []
                    
                    if param.options_mode == "manual" and param.options_manual:
                        options = list(param.options_manual)
                    
                    elif param.options_mode == "from_column" and param.options_source:
                        # Essayer de résoudre depuis les données du projet
                        try:
                            col_name = param.options_source.get("column")
                            gab_name = param.options_source.get("gabarit")
                            gab_ver = param.options_source.get("version", "v1")
                            
                            # Charger la source de données du projet
                            data_source = ps.get_data_source(project_id, gab_name, gab_ver)
                            
                            if data_source:
                                # Charger les données et extraire les valeurs uniques
                                from backend.services.dataset_service import _load_dataframe_from_source
                                
                                df = _load_dataframe_from_source(data_source.get("source_config"))
                                
                                if df is not None and col_name in df.columns:
                                    options = sorted(list(df[col_name].dropna().astype(str).unique()))[:500]
                        
                        except Exception as e:
                            st.warning(f"⚠️ Impossible de charger les options depuis les données : {e}")
                    
                    # Cache si disponible
                    if not options and hasattr(param, 'options_cache') and param.options_cache:
                        cache_values = param.options_cache.get('values', [])
                        if cache_values:
                            options = list(cache_values)
                    
                    # Widget
                    if options:
                        try:
                            default_idx = options.index(str(custom_value)) if str(custom_value) in options else 0
                        except (ValueError, TypeError):
                            default_idx = 0
                        
                        new_value = st.selectbox(
                            "Valeur par défaut",
                            options=options,
                            index=default_idx,
                            key=f"param_{param_name}",
                            label_visibility="collapsed"
                        )
                    else:
                        new_value = st.text_input(
                            "Valeur par défaut",
                            value=str(custom_value) if custom_value else "",
                            key=f"param_{param_name}",
                            label_visibility="collapsed"
                        )
                
                elif param.type == "integer":
                    new_value = st.number_input(
                        "Valeur par défaut",
                        value=int(custom_value) if custom_value is not None else 0,
                        key=f"param_{param_name}",
                        label_visibility="collapsed"
                    )
                
                elif param.type == "date":
                    from datetime import datetime, date
                    
                    if isinstance(custom_value, str):
                        try:
                            custom_value = datetime.fromisoformat(custom_value).date()
                        except:
                            custom_value = date.today()
                    elif not isinstance(custom_value, date):
                        custom_value = date.today()
                    
                    new_value = st.date_input(
                        "Valeur par défaut",
                        value=custom_value,
                        key=f"param_{param_name}",
                        label_visibility="collapsed"
                    ).isoformat()
                
                else:
                    new_value = st.text_input(
                        "Valeur par défaut",
                        value=str(custom_value) if custom_value else "",
                        key=f"param_{param_name}",
                        label_visibility="collapsed"
                    )
                
                # Sauvegarder dans l'état
                if param_name not in custom_params:
                    custom_params[param_name] = {}
                
                custom_params[param_name]["default"] = new_value
        
        st.markdown("---")
        
        # Bouton enregistrer
        col_save1, col_save2 = st.columns(2)
        
        with col_save1:
            if st.button("💾 Enregistrer les paramètres", type="primary", use_container_width=True):
                try:
                    # Mettre à jour le livrable
                    for d in proj["deliverables"]:
                        if d["template_id"] == template_id:
                            d["custom_parameters"] = custom_params
                            break
                    
                    ps.save_project(proj)
                    
                    st.success("✅ Paramètres enregistrés pour ce projet")
                    st.balloons()
                    
                    import time
                    time.sleep(1)
                    st.rerun()
                
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
        
        with col_save2:
            if st.button("🔄 Réinitialiser", use_container_width=True):
                try:
                    # Vider les paramètres personnalisés
                    for d in proj["deliverables"]:
                        if d["template_id"] == template_id:
                            d["custom_parameters"] = {}
                            break
                    
                    ps.save_project(proj)
                    
                    st.info("Paramètres réinitialisés aux valeurs du template")
                    st.rerun()
                
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")

# ==================== ONGLET 3 : STATISTIQUES ====================
with tab_stats:
    st.subheader("📊 Statistiques du livrable")
    
    # Statut global
    col_s1, col_s2, col_s3 = st.columns(3)
    
    with col_s1:
        st.metric("Complétude", f"{deliverable.get('completion_rate', 0)}%")
    
    with col_s2:
        ds_status = deliverable.get("data_sources_status", {})
        st.metric("Sources client", ds_status.get("client", 0))
    
    with col_s3:
        st.metric("Sources défaut", ds_status.get("default", 0))
    
    st.markdown("---")
    
    # Gabarits requis
    st.markdown("### 🗂️ Gabarits requis")
    
    usages = ts.list_gabarit_usages(template_id)
    
    if not usages:
        st.info("Aucun gabarit requis par ce template")
    else:
        for u in usages:
            gname = u.get("gabarit_name")
            gver = u.get("gabarit_version", "v1")
            
            with st.container(border=True):
                col_g1, col_g2, col_g3 = st.columns([2, 1, 1])
                
                with col_g1:
                    st.markdown(f"**{gname}** (v{gver})")
                    
                    # Colonnes requises
                    cols = u.get("columns_enabled", [])
                    if cols:
                        st.caption(f"📊 {len(cols)} colonne(s)")
                
                with col_g2:
                    # Statut source
                    data_source = ps.get_data_source(project_id, gname, gver)
                    
                    if data_source:
                        if data_source.get("source_type") == "client":
                            st.markdown("🟢 **Client**")
                        else:
                            st.markdown("🟡 **Défaut**")
                    else:
                        # Vérifier si source par défaut existe
                        from backend.services.gabarit_registry import get_default_source
                        default_src = get_default_source(gname, gver)
                        
                        if default_src:
                            st.markdown("🟡 **Défaut**")
                        else:
                            st.markdown("🔴 **Manquant**")
                
                with col_g3:
                    # Statistiques si disponibles
                    if data_source and data_source.get("row_count"):
                        st.metric("Lignes", data_source["row_count"])
    
    st.markdown("---")
    
    # Dernière génération
    st.markdown("### 🕒 Dernière génération")
    
    if deliverable.get("last_generated_at"):
        from datetime import datetime
        try:
            dt = datetime.fromisoformat(deliverable["last_generated_at"])
            st.info(f"📅 Généré le {dt.strftime('%d/%m/%Y à %H:%M')}")
        except:
            st.caption("Date non disponible")
    else:
        st.info("Jamais généré")
    
    st.markdown("---")
    
    # Bouton recalculer statut
    if st.button("🔄 Recalculer le statut", use_container_width=True):
        try:
            ps.update_deliverable_status(project_id, template_id)
            st.success("✅ Statut mis à jour")
            st.rerun()
        
        except Exception as e:
            st.error(f"❌ Erreur : {e}")

# ===== ACTIONS GLOBALES =====
st.markdown("---")

col_back, col_gen = st.columns(2)

with col_back:
    if st.button("← Retour au Hub", use_container_width=True):
        if 'selected_deliverable_id' in st.session_state:
            del st.session_state.selected_deliverable_id
        st.switch_page("pages/_1a_🗂️_Hub_Projet.py")

with col_gen:
    if st.button("▶️ Générer ce livrable", 
                type="primary", 
                use_container_width=True,
                disabled=not deliverable.get("is_functional")):
        st.info("🚧 Génération en cours d'implémentation (Phase 6)")