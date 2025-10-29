# frontend/pages/_1b_💾_Data_Projet.py
import streamlit as st
from pathlib import Path
import sys
import pandas as pd

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService
from backend.services.gabarit_registry import get_gabarit, get_default_source
from code_editor import code_editor

st.set_page_config(page_title="Données Projet", page_icon="💾", layout="wide")

# ===== NAVBAR PROJET =====
def render_project_subnav(active: str):
    cols = st.columns([1, 1, 1])
    
    with cols[0]:
        if st.button("← Projets", use_container_width=True):
            if 'selected_project_id' in st.session_state:
                del st.session_state.selected_project_id
            if 'selected_deliverable_id' in st.session_state:
                del st.session_state.selected_deliverable_id
            st.switch_page("pages/4_📁_Projets.py")
    
    with cols[1]:
        if st.button("🗂️ Hub", 
                    type="primary" if active == "hub" else "secondary",
                    use_container_width=True):
            if 'selected_deliverable_id' in st.session_state:
                del st.session_state.selected_deliverable_id
            st.switch_page("pages/_4a_🗂️_Hub_Projet.py")
    
    with cols[2]:
        if st.button("💾 Données", 
                    type="primary" if active == "data" else "secondary",
                    use_container_width=True):
            if 'selected_deliverable_id' in st.session_state:
                del st.session_state.selected_deliverable_id
            st.switch_page("pages/_4b_💾_Data_Projet.py")
    
    st.divider()

# ===== GUARD =====
if 'selected_project_id' not in st.session_state or not st.session_state.selected_project_id:
    st.error("Aucun projet sélectionné")
    if st.button("← Retour aux projets"):
        st.switch_page("pages/4_📁_Projets.py")
    st.stop()

project_id = st.session_state.selected_project_id

# ===== CHARGEMENT =====
DatabaseService.initialize()
with DatabaseService.get_session() as db:
    ps = ProjectService(db)
    
    try:
        proj = ps.load_project(project_id)
    except FileNotFoundError:
        st.error(f"Projet introuvable : {project_id}")
        if st.button("← Retour aux projets"):
            st.switch_page("pages/4_📁_Projets.py")
        st.stop()

render_project_subnav("data")

# ===== EN-TÊTE =====
st.title(f"💾 Configuration des données — {proj.get('name', '(sans nom)')}")
st.caption("Configurez les sources de données pour chaque gabarit requis par vos livrables")

st.divider()

# ===== CALCUL DES GABARITS REQUIS =====
def compute_required_gabarits(proj: dict) -> dict:
    """
    Calcule l'union des gabarits requis par tous les livrables.
    INCLUT les gabarits d'enrichissement (dimensions).
    Retourne : {(gabarit_name, gabarit_version): {"columns": [...], "templates": [...]}}
    """
    from backend.services.template_service import TemplateService
    
    required = {}
    
    for deliv in proj.get("deliverables", []):
        template_id = deliv.get("template_id")
        if not template_id:
            continue
        
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            usages = ts.list_gabarit_usages(template_id)
        
        for u in usages:
            gname = u.get("gabarit_name")
            gver = u.get("gabarit_version", "v1")
            
            if not gname:
                continue
            
            key = (gname, gver)
            
            if key not in required:
                required[key] = {
                    "columns": [],
                    "templates": []
                }
            
            # Ajouter les colonnes de la table principale (union)
            cols = u.get("columns_enabled", [])
            if not cols:
                # Si aucune colonne spécifique, prendre toutes les colonnes du gabarit
                gab = get_gabarit(gname, gver)
                if gab:
                    cols = [c.name for c in gab.columns]
            
            for c in cols:
                if c not in required[key]["columns"]:
                    required[key]["columns"].append(c)
            
            # Ajouter le template
            tpl_name = deliv.get("template_name", f"Template #{template_id}")
            if tpl_name not in required[key]["templates"]:
                required[key]["templates"].append(tpl_name)
            
            # ========= NOUVEAU : Extraire les gabarits d'enrichissement =========
            enrichments = u.get("enrichments", [])
            for enrich in enrichments:
                if not isinstance(enrich, dict):
                    continue
                
                # Extraire les gabarits depuis le path
                # Format: [["SELL-IN", "CODE_CLIENT", "Dim_clients_fmcg", "CODE_CLIENT"], ...]
                path = enrich.get("path", [])
                for step in path:
                    if not isinstance(step, list) or len(step) < 3:
                        continue
                    
                    # step[2] est le gabarit de dimension
                    dim_name = step[2]
                    dim_ver = "v1"  # Par défaut v1, ou extraire si disponible
                    
                    dim_key = (dim_name, dim_ver)
                    
                    if dim_key not in required:
                        required[dim_key] = {
                            "columns": [],
                            "templates": []
                        }
                    
                    # Ajouter les colonnes demandées pour cet enrichissement
                    enrich_cols = enrich.get("columns", [])
                    for c in enrich_cols:
                        if c not in required[dim_key]["columns"]:
                            required[dim_key]["columns"].append(c)
                    
                    # Ajouter le template
                    if tpl_name not in required[dim_key]["templates"]:
                        required[dim_key]["templates"].append(tpl_name)
    
    return required

required_gabarits = compute_required_gabarits(proj)

if not required_gabarits:
    st.info("Aucun livrable configuré. Retournez au Hub pour ajouter des livrables.")
    if st.button("← Retour au Hub", use_container_width=True):
        st.switch_page("pages/_4a_🗂️_Hub_Projet.py")
    st.stop()

# ===== STATISTIQUES GLOBALES =====
col_m1, col_m2, col_m3 = st.columns(3)

data_sources = proj.get("data_sources", [])
configured_count = len(data_sources)
client_count = sum(1 for s in data_sources if s.get("source_type") == "client")
default_count = sum(1 for s in data_sources if s.get("source_type") == "default")

with col_m1:
    st.metric("Gabarits requis", len(required_gabarits))

with col_m2:
    st.metric("Sources configurées", f"{configured_count}/{len(required_gabarits)}")

with col_m3:
    st.metric("Sources client", client_count)

st.divider()

# ===== CONFIGURATION PAR GABARIT =====
st.subheader("🗂️ Configuration par gabarit")

for (gname, gver), info in sorted(required_gabarits.items()):
    # Charger le gabarit
    gab = get_gabarit(gname, gver)
    if not gab:
        continue
    
    # Récupérer la source configurée
    current_source = ps.get_data_source(project_id, gname, gver)
    
    # Vérifier si une source par défaut existe
    default_src = get_default_source(gname, gver)
    
    with st.expander(f"**{gname}** (v{gver})", expanded=not current_source):
        # Informations sur le gabarit
        col_info, col_status = st.columns([3, 1])
        
        with col_info:
            st.caption(f"📊 {len(info['columns'])} colonne(s) requise(s)")
            st.caption(f"📦 Utilisé par : {', '.join(info['templates'])}")
        
        with col_status:
            if current_source:
                if current_source.get("source_type") == "client":
                    st.markdown("🟢 **Client**")
                else:
                    st.markdown("🟡 **Défaut**")
            else:
                if default_src:
                    st.markdown("🟡 **Défaut**")
                else:
                    st.markdown("🔴 **Non configuré**")
        
        st.markdown("---")
        
        # Radio choix source
        current_type = "default"
        if current_source:
            current_type = current_source.get("source_type", "default")
        elif not default_src:
            current_type = "client"
        
        source_type = st.radio(
            "Source de données",
            ["default", "client"],
            index=0 if current_type == "default" else 1,
            format_func=lambda x: "🔄 Donnée par défaut du gabarit" if x == "default" else "📁 Donnée client",
            key=f"source_type_{gname}_{gver}",
            horizontal=True,
            disabled=not default_src and current_type == "default"
        )
        
        # === AFFICHAGE SOURCE PAR DÉFAUT ===
        if source_type == "default":
            if not default_src:
                st.warning("⚠️ Aucune donnée par défaut configurée pour ce gabarit. Sélectionnez 'Donnée client'.")
            else:
                st.success("✅ Utilise la donnée par défaut du gabarit")
                
                # Afficher aperçu si disponible
                from backend.services.gabarit_registry import get_default_preview
                preview = get_default_preview(gname, gver)
                
                if preview and preview.get("rows"):
                    with st.expander("👁️ Aperçu des données", expanded=False):
                        rows = preview.get("rows", [])
                        cols = preview.get("columns", [])
                        
                        if rows and cols:
                            df_preview = pd.DataFrame(rows, columns=cols)
                            st.dataframe(df_preview, use_container_width=True, height=300)
                            st.caption(f"📊 {len(rows)} lignes × {len(cols)} colonnes")
                
                # Bouton enregistrer
                if st.button("💾 Utiliser cette source", key=f"save_default_{gname}_{gver}", use_container_width=True):
                    try:
                        ps.set_data_source(project_id, gname, gver, "default", default_src)
                        
                        # Mettre à jour les statuts des livrables
                        for deliv in proj.get("deliverables", []):
                            ps.update_deliverable_status(project_id, deliv["template_id"])
                        
                        st.success("✅ Source par défaut configurée")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Erreur : {e}")
        
        # === CONFIGURATION SOURCE CLIENT ===
        else:
            st.markdown("### 📁 Configuration source client")
            
            # Clé de persistance unique par gabarit
            buffer_key = f"client_source_{gname}_{gver}"
            
            # Initialiser avec les valeurs existantes si disponibles
            if buffer_key not in st.session_state:
                if current_source and current_source.get("source_type") == "client":
                    cfg = current_source.get("source_config", {})
                    st.session_state[buffer_key] = {
                        "path": cfg.get("path", ""),
                        "sep": cfg.get("sep", ";"),
                        "encoding": cfg.get("encoding", "utf-8-sig"),
                        "python": cfg.get("python", "")
                    }
                else:
                    st.session_state[buffer_key] = {
                        "path": "",
                        "sep": ";",
                        "encoding": "utf-8-sig",
                        "python": ""
                    }
            
            # Formulaire configuration CSV
            col1, col2, col3 = st.columns([3, 1, 1])
            
            with col1:
                path = st.text_input(
                    "Chemin du fichier CSV",
                    value=st.session_state[buffer_key]["path"],
                    placeholder="C:/data/clients/fichier.csv",
                    key=f"path_{gname}_{gver}"
                )
                st.session_state[buffer_key]["path"] = path
            
            with col2:
                sep = st.text_input(
                    "Séparateur",
                    value=st.session_state[buffer_key]["sep"],
                    key=f"sep_{gname}_{gver}"
                )
                st.session_state[buffer_key]["sep"] = sep
            
            with col3:
                encoding = st.text_input(
                    "Encodage",
                    value=st.session_state[buffer_key]["encoding"],
                    key=f"encoding_{gname}_{gver}"
                )
                st.session_state[buffer_key]["encoding"] = encoding
            
            # Script Python (avec code_editor)
            with st.expander("🐍 Transformation Python (optionnel)", expanded=bool(st.session_state[buffer_key]["python"])):
                st.caption("💡 Variables : `df` (DataFrame), `pd` (pandas)")
                st.caption("⚠️ Le DataFrame doit être réassigné : `df = df[...]`")
                
                custom_buttons = [{
                    "name": "Copier", "feather": "Copy", "hasText": True,
                    "commands": ["copyAll"], "style": {"top": "0.46rem", "right": "0.4rem"}
                }]
                
                editor_result = code_editor(
                    st.session_state[buffer_key]["python"],
                    lang="python", height=200, theme="contrast", shortcuts="vscode",
                    focus=False, buttons=custom_buttons, allow_reset=True,
                    options={
                        "wrap": True, "showLineNumbers": True, "highlightActiveLine": True,
                        "enableLiveAutocompletion": True, "enableBasicAutocompletion": True
                    },
                    key=f"python_editor_{gname}_{gver}",
                    response_mode=["blur"]
                )
                
                # Capture du code
                if editor_result:
                    new_code = None
                    if isinstance(editor_result, dict):
                        new_code = (editor_result.get("text") 
                                   or editor_result.get("content") 
                                   or editor_result.get("code"))
                    elif isinstance(editor_result, str):
                        new_code = editor_result
                    
                    if isinstance(new_code, str):
                        st.session_state[buffer_key]["python"] = new_code
                
                st.caption(f"📝 {len(st.session_state[buffer_key]['python'])} caractères")
            
            st.markdown("---")
            
            # Actions
            col_preview, col_validate, col_save = st.columns(3)
            
            with col_preview:
                if st.button("👁️ Aperçu", key=f"preview_{gname}_{gver}", use_container_width=True, disabled=not path):
                    try:
                        # Charger un échantillon
                        if not Path(path).exists():
                            st.error(f"❌ Fichier introuvable : {path}")
                        else:
                            df = pd.read_csv(path, sep=sep, encoding=encoding, nrows=20)
                            
                            # Appliquer le script Python
                            code = st.session_state[buffer_key]["python"]
                            if code and code.strip():
                                loc = {"df": df.copy(), "pd": pd}
                                try:
                                    exec(code, {}, loc)
                                    if isinstance(loc.get("df"), pd.DataFrame):
                                        df = loc["df"]
                                except Exception as e:
                                    st.error(f"❌ Erreur script : {e}")
                            
                            st.success(f"✅ Aperçu chargé : {df.shape[0]} lignes × {df.shape[1]} colonnes")
                            st.dataframe(df, use_container_width=True, height=300)
                    
                    except Exception as e:
                        st.error(f"❌ Erreur : {e}")
            
            with col_validate:
                if st.button("✅ Valider", key=f"validate_{gname}_{gver}", use_container_width=True, disabled=not path):
                    try:
                        # Construire la config source
                        source_config = {
                            "type": "csv",
                            "path": path,
                            "sep": sep,
                            "encoding": encoding
                        }
                        if st.session_state[buffer_key]["python"].strip():
                            source_config["python"] = st.session_state[buffer_key]["python"]

                        # Sauvegarder temporairement (pour que validate lise la même conf)
                        ps.set_data_source(project_id, gname, gver, "client", source_config)

                        # Valider (retourne lignes/colonnes + diff de schéma)
                        result = ps.validate_data_source(project_id, gname, gver)

                        st.success("✅ Validation effectuée")

                        # KPIs
                        col_r1, col_r2 = st.columns(2)
                        with col_r1:
                            st.metric("Lignes", result.get("row_count", "—"))
                        with col_r2:
                            st.metric("Colonnes", len(result.get("columns", [])))

                        # Complétude
                        if result.get("columns_filled"):
                            with st.expander("📊 Complétude par colonne", expanded=False):
                                completeness = result["columns_filled"]
                                df_comp = pd.DataFrame([
                                    {"Colonne": col, "Complétude (%)": pct}
                                    for col, pct in sorted(completeness.items(), key=lambda x: x[1], reverse=True)
                                ])
                                st.dataframe(df_comp, use_container_width=True, hide_index=True, height=300)

                        # Diff de schéma
                        with st.expander("🧩 Écart au gabarit attendu", expanded=True):
                            exp_cols = result.get("expected_columns", []) or []
                            miss = result.get("missing_columns", []) or []
                            extra = result.get("extra_columns", []) or []
                            dtypes = result.get("dtype_mismatches", {}) or {}

                            c1, c2, c3 = st.columns(3)
                            with c1:
                                st.metric("Colonnes attendues", len(exp_cols))
                            with c2:
                                st.metric("Manquantes", len(miss))
                            with c3:
                                st.metric("En trop", len(extra))

                            if miss:
                                st.warning("Colonnes **manquantes** :", icon="⚠️")
                                st.code(", ".join(miss))

                            if extra:
                                st.info("Colonnes **en trop** :", icon="ℹ️")
                                st.code(", ".join(extra))

                            if dtypes:
                                st.error("Incohérences de **types** :", icon="❗")
                                df_types = pd.DataFrame([
                                    {"Colonne": k, "Attendu": v.get("expected"), "Trouvé": v.get("found")}
                                    for k, v in dtypes.items()
                                ])
                                st.dataframe(df_types, use_container_width=True, hide_index=True)
                            else:
                                st.success("Aucune incohérence de type détectée (ou type non spécifié dans le gabarit).")

                    except Exception as e:
                        st.error(f"❌ Erreur validation : {e}")
                        import traceback
                        with st.expander("🔍 Détails"):
                            st.code(traceback.format_exc())

            
            with col_save:
                if st.button("💾 Enregistrer", key=f"save_{gname}_{gver}", type="primary", use_container_width=True, disabled=not path):
                    try:
                        # Construire la config source
                        source_config = {
                            "type": "csv",
                            "path": path,
                            "sep": sep,
                            "encoding": encoding
                        }
                        
                        if st.session_state[buffer_key]["python"].strip():
                            source_config["python"] = st.session_state[buffer_key]["python"]
                        
                        # Sauvegarder
                        ps.set_data_source(project_id, gname, gver, "client", source_config)
                        
                        # Valider pour avoir les stats
                        ps.validate_data_source(project_id, gname, gver)
                        
                        # Mettre à jour les statuts des livrables
                        for deliv in proj.get("deliverables", []):
                            ps.update_deliverable_status(project_id, deliv["template_id"])
                        
                        st.success("✅ Source client enregistrée")
                        st.balloons()
                        
                        import time
                        time.sleep(1)
                        st.rerun()
                    
                    except Exception as e:
                        st.error(f"❌ Erreur : {e}")
            
            # Afficher les stats si source déjà validée
            if current_source and current_source.get("row_count"):
                st.markdown("---")
                st.markdown("### 📊 Statistiques de la source actuelle")
                
                col_s1, col_s2, col_s3 = st.columns(3)
                
                with col_s1:
                    st.metric("Lignes", current_source.get("row_count", "—"))
                
                with col_s2:
                    cols_filled = current_source.get("columns_filled", {})
                    st.metric("Colonnes", len(cols_filled))
                
                with col_s3:
                    if current_source.get("last_validated_at"):
                        from datetime import datetime
                        try:
                            dt = datetime.fromisoformat(current_source["last_validated_at"])
                            st.metric("Validée", dt.strftime("%d/%m %H:%M"))
                        except:
                            st.metric("Validée", "—")
                
                # Détails complétude
                if cols_filled:
                    with st.expander("📊 Complétude par colonne", expanded=False):
                        df_comp = pd.DataFrame([
                            {"Colonne": col, "Complétude (%)": pct}
                            for col, pct in sorted(cols_filled.items(), key=lambda x: x[1], reverse=True)
                        ])
                        
                        st.dataframe(df_comp, use_container_width=True, hide_index=True, height=300)

st.markdown("---")

# ===== ACTIONS GLOBALES =====
col_back, col_refresh = st.columns(2)

with col_back:
    if st.button("← Retour au Hub", use_container_width=True):
        st.switch_page("pages/_4a_🗂️_Hub_Projet.py")

with col_refresh:
    if st.button("🔄 Recalculer les statuts", use_container_width=True):
        try:
            with DatabaseService.get_session() as db:
                ps_refresh = ProjectService(db)
                
                for deliv in proj.get("deliverables", []):
                    ps_refresh.update_deliverable_status(project_id, deliv["template_id"])
            
            st.success("✅ Statuts mis à jour")
            st.rerun()
        
        except Exception as e:
            st.error(f"❌ Erreur : {e}")