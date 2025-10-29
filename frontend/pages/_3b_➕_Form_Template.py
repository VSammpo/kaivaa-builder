# pages/_2b_➕_Form_Template.py
"""
Page de création de templates - VERSION COMPLÈTE AVEC PARAMÈTRES ENRICHIS
"""

import streamlit as st
from pathlib import Path
import sys
import json
import pandas as pd

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.services.parameter_service import ParameterService
from backend.models.template_config import (
    TemplateConfig,
    ParameterConfig,
    DataSourceConfig,
    LoopConfig,
    ImageInjection,
    SlideMapping,
    TemplateTags,  # ✅ ajout
)

st.set_page_config(page_title="Paramètres généraux", page_icon="⚙️", layout="wide")
# --- Flash success (à mettre juste sous st.set_page_config) ---
if msg := st.session_state.pop("_flash_success", None):
    st.success(msg)
    st.toast(msg)


# --- Helpers de chemin/version ---
from pathlib import Path

def _templates_root_dir() -> Path:
    # déjà présent chez toi normalement ; sinon garde celui-ci
    from pathlib import Path
    return Path(__file__).resolve().parents[2] / "configuration" / "templates"

def _ensure_version_layout(name: str, version: str) -> Path:
    """
    Garantit l’arborescence versionnée et migre l'ancien fichier <version>.json
    qui pourrait traîner à la racine du template vers <Version>/config.json.
    """
    root = _templates_root_dir() / name
    version_dir = root / str(version)
    version_dir.mkdir(parents=True, exist_ok=True)

    legacy_json = root / f"{version}.json"
    cfg_path = version_dir / "config.json"

    if legacy_json.exists() and not cfg_path.exists():
        try:
            cfg_path.write_bytes(legacy_json.read_bytes())
            legacy_json.unlink()
        except Exception:
            pass  # on ne casse pas si la migration échoue

    return version_dir


def render_template_subnav(active: str, template_id: int | None):
    cols = st.columns([1,1,1,1,1])
    with cols[0]:
        if st.button("← Retour bibliothèque", use_container_width=True):
            if template_id:
                st.session_state.selected_template_detail = template_id
            if "selected_template" in st.session_state:
                del st.session_state.selected_template
            st.switch_page("pages/3_📚_Bibliotheque.py")
    with cols[1]:
        if st.button("🗂️ Détail du template", type=("primary" if active=="detail" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template_detail = template_id
            st.switch_page("pages/_3a_📊_Detail_Livrable.py")
    with cols[2]:
        st.button("⚙️ Paramètres généraux", type="primary" if active=="general" else "secondary", use_container_width=True)
    with cols[3]:
        if st.button("📑 Injection des données", type=("primary" if active=="inject" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_3c_📑_Tables_Template.py")
    with cols[4]:
        if st.button("🧾 Ajustement de la table", type=("primary" if active=="adjust" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_3d_🧾_Ajustement_Table.py")
    st.divider()

# États init
if 'parameters' not in st.session_state:
    st.session_state.parameters = []
if 'loops' not in st.session_state:
    st.session_state.loops = []
if 'images' not in st.session_state:
    st.session_state.images = {}
if 'mappings' not in st.session_state:
    st.session_state.mappings = []

# Guard + chargement
edit_mode = False
template_id_to_edit = None

if 'selected_template' not in st.session_state and 'selected_template_detail' in st.session_state:
    st.session_state.selected_template = st.session_state.selected_template_detail

if 'selected_template' in st.session_state and st.session_state.selected_template:
    edit_mode = True
    template_id_to_edit = st.session_state.selected_template
    with DatabaseService.get_session() as db:
        service = TemplateService(db)
        template_config = service.load_template_config(template_id_to_edit)
        template_db = service.get_template(template_id_to_edit)
        template_name = template_db.name
        template_version = template_db.version
        template_description = template_db.description
        template_card_image_path = template_db.card_image_path

    render_template_subnav("general", template_id_to_edit)
    st.title(f"⚙️ Paramètres généraux — {template_name} (v{template_version})")

    if not st.session_state.get('_template_loaded'):
        st.session_state.parameters = [p.model_dump() for p in template_config.parameters]
        st.session_state.loops = [
            {"loop_id": loop.loop_id, "slides": loop.slides, "sheet_name": loop.sheet_name}
            for loop in template_config.loops
        ]
        st.session_state.images = {
            slide_id: [
                {"type": img.type, "pattern": img.pattern, "default_path": img.default_path,
                 "position": img.position, "size": img.size, "background": img.background,
                 "loop_dependent": img.loop_dependent}
                for img in images
            ] for slide_id, images in template_config.image_injections.items()
        }
        st.session_state.mappings = [
            {"slide_id": m.slide_id, "sheet_name": m.sheet_name, "excel_range": m.excel_range, "has_header": m.has_header}
            for m in template_config.slide_mappings
        ]
        st.session_state._template_loaded = True
else:
    render_template_subnav("general", None)
    st.title("➕ Créer un nouveau template")
    template_name = ""
    template_version = "1.0"
    template_description = ""
    template_card_image_path = None

# === State key unique par template (évite la fuite A↔B) ===
tid_key = str(template_id_to_edit) if edit_mode else "new"
THEMES_KEY  = f"_tags_themes_{tid_key}"
FAMS_KEY    = f"_tags_families_{tid_key}"
W_THEME_KEY = f"new_theme_input_{tid_key}"
W_FAM_KEY   = f"new_family_input_{tid_key}"


# Sidebar récap
with st.sidebar:
    st.header("🧭 Récap")
    st.write(f"**Mode** : {'Édition' if edit_mode else 'Création'}")
    st.write(f"**Nom** : {template_name or '—'}")
    st.write(f"**Version** : {template_version if edit_mode else '—'}")
    st.write(f"**Source** : {template_config.data_source.type if edit_mode else '—'}")

    if edit_mode and template_card_image_path:
        try:
            st.image(template_card_image_path, caption="Image actuelle", use_container_width=True)
        except Exception:
            st.caption("Aperçu indisponible.")

    st.divider()
    st.caption(f"Paramètres : {len(st.session_state.parameters)}")
    st.caption(f"Boucles : {len(st.session_state.loops)}")
    st.caption(f"Images dynamiques : {sum(len(v) for v in st.session_state.images.values())}")
    st.caption(f"Mappings : {len(st.session_state.mappings)}")


# ===== ÉTAPE 1 : Informations générales =====
st.header("1️⃣ Informations générales")

col1, col2 = st.columns(2)

with col1:
    name = st.text_input(
        "Nom du template*", 
        value=template_name if edit_mode else "",
        placeholder="ex: BCE_INSEE",
        disabled=edit_mode
    )
    version = st.text_input(
        "Version", 
        value=template_version if edit_mode else "1.0"
    )

with col2:
    description = st.text_area(
        "Description", 
        value=template_description if edit_mode and template_description else "",
        placeholder="Description du template..."
    )

# === Tags : Thématiques & Familles ===
# (1) Charger les valeurs existantes pour alimenter les listes (union simple)
with DatabaseService.get_session() as db:
    _svc_for_tags = TemplateService(db)
    existing_tags = _svc_for_tags.list_all_tag_values()

# (2) Initialiser l'état persistant PROPRE AU TEMPLATE COURANT
if THEMES_KEY not in st.session_state or FAMS_KEY not in st.session_state:
    pref_themes, pref_families = [], []
    if edit_mode:
        with DatabaseService.get_session() as db:
            _svc = TemplateService(db)
            _cfg = _svc.get_config(template_id_to_edit) or {}
            _tags = (_cfg.get("tags") or {})
            pref_themes   = list(_tags.get("themes")   or [])
            pref_families = list(_tags.get("families") or [])
    st.session_state[THEMES_KEY] = pref_themes
    st.session_state[FAMS_KEY]   = pref_families

# (3) Options = existants ∪ sélection courante (sinon Streamlit n’affiche pas les valeurs non listées)
theme_options  = sorted(set(existing_tags.get("themes", []))   | set(st.session_state[THEMES_KEY]), key=str.lower)
family_options = sorted(set(existing_tags.get("families", [])) | set(st.session_state[FAMS_KEY]),   key=str.lower)

tag_col1, tag_col2 = st.columns(2)
with tag_col1:
    selected_themes = st.multiselect(
        "Thématiques",
        options=theme_options,
        default=st.session_state[THEMES_KEY],
        placeholder="Sélectionner des thématiques…",
        key=f"ms_themes_{tid_key}"
    )
    st.session_state[THEMES_KEY] = selected_themes

    new_theme = st.text_input("➕ Ajouter une thématique", value="", key=W_THEME_KEY)
    if st.button("Ajouter la thématique", key=f"btn_add_theme_{tid_key}"):
        v = (new_theme or "").strip()
        if v and v not in st.session_state[THEMES_KEY]:
            st.session_state[THEMES_KEY].append(v)
        st.rerun()

with tag_col2:
    selected_families = st.multiselect(
        "Familles",
        options=family_options,
        default=st.session_state[FAMS_KEY],
        placeholder="Sélectionner des familles…",
        key=f"ms_families_{tid_key}"
    )
    st.session_state[FAMS_KEY] = selected_families

    new_family = st.text_input("➕ Ajouter une famille", value="", key=W_FAM_KEY)
    if st.button("Ajouter la famille", key=f"btn_add_family_{tid_key}"):
        v = (new_family or "").strip()
        if v and v not in st.session_state[FAMS_KEY]:
            st.session_state[FAMS_KEY].append(v)
        st.rerun()

st.divider()

# ===== ÉTAPE 2 : Fichiers sources =====
st.header("2️⃣ Fichiers sources")

col1, col2 = st.columns(2)

with col1:
    upload_mode = st.radio(
        "Mode de création",
        ["Créer des fichiers vierges", "Uploader des fichiers existants"]
    )

with col2:
    card_image = st.file_uploader(
        "🖼️ Image de carte (optionnelle)", 
        type=['png', 'jpg', 'jpeg'],
        help="Image affichée dans la bibliothèque. Si vide, une image par défaut sera utilisée."
    )

ppt_file = None
excel_file = None

if upload_mode == "Uploader des fichiers existants":
    col1, col2 = st.columns(2)
    
    with col1:
        ppt_file = st.file_uploader("Template PowerPoint*", type=['pptx'])
    
    with col2:
        excel_file = st.file_uploader("Template Excel*", type=['xlsx'])

st.divider()

# ===== PARTIES NÉCESSAIRES =====
st.header("📋 Parties nécessaires")
st.caption("Activez uniquement les sections dont vous avez besoin pour ce template")

col1, col2, col3, col4 = st.columns(4)

with col1:
    show_params = st.checkbox(
        "Paramètres d'utilisation",
        value=len(st.session_state.parameters) > 0,
        help="Variables à renseigner lors de chaque génération (ex: période, enseigne, segment)"
    )

with col2:
    show_loops = st.checkbox(
        "Pages itératives",
        value=len(st.session_state.loops) > 0,
        help="Slides à dupliquer selon un tableau Loop (ex: une slide par produit)"
    )

with col3:
    show_images = st.checkbox(
        "Images dynamiques",
        value=len(st.session_state.images) > 0,
        help="Images injectées automatiquement selon des patterns (logos, photos produits)"
    )

with col4:
    show_mappings = st.checkbox(
        "Tableaux dynamiques",
        value=len(st.session_state.mappings) > 0,
        help="Données Excel à injecter dans des tableaux PowerPoint"
    )

st.divider()


# ===== ÉTAPE 3 : PARAMÈTRES (UI MODERNE) =====
if show_params:
    st.header("3️⃣ Paramètres d'utilisation")
    st.markdown("Variables demandées à l'utilisateur lors de chaque génération de livrable")
    
    # Liste des gabarits disponibles (pour options dynamiques)
    from backend.services.gabarit_registry import list_gabarits, get_gabarit
    all_gabarits = list_gabarits()
    gabarit_options = ["(aucun)"] + [f"{g.name} (v{g.version})" for g in all_gabarits]
    
    # Bouton ajout
    if st.button("➕ Ajouter un paramètre", use_container_width=True, type="primary"):
        st.session_state.parameters.append({
            "name": "",
            "type": "string",
            "required": True,
            "balise_ppt": "",
            "description": "",
            "default": None,
            "options_mode": "none",
            "options_manual": [],
            "options_source": None
        })
        st.rerun()
    
    st.markdown("")
    
    # Affichage des paramètres (style gabarit)
    to_delete = []
    for idx, param in enumerate(st.session_state.parameters):
        param_type_label = param.get('type', 'string')
        param_name_label = param.get('name') or '(sans nom)'
        
        with st.expander(f"**{idx+1}. {param_name_label}** — {param_type_label}", expanded=True):
            
            # Ligne 1 : Nom + Balise PPT
            col1, col2 = st.columns(2)
            with col1:
                param["name"] = st.text_input(
                    "Nom du paramètre *",
                    value=param.get("name", ""),
                    placeholder="ex: sous_marque",
                    key=f"param_name_{idx}"
                )
            with col2:
                param["balise_ppt"] = st.text_input(
                    "Balise PPT *",
                    value=param.get("balise_ppt", ""),
                    placeholder="ex: [Sous_Marque]",
                    key=f"param_balise_{idx}"
                )
            
            # Ligne 2 : Type + Obligatoire
            col1, col2 = st.columns(2)
            with col1:
                param["type"] = st.selectbox(
                    "Type",
                    ["string", "integer", "date", "liste"],
                    index=["string", "integer", "date", "liste"].index(param.get("type", "string")),
                    key=f"param_type_{idx}"
                )
            with col2:
                param["required"] = st.checkbox(
                    "Obligatoire",
                    value=param.get("required", True),
                    key=f"param_required_{idx}"
                )
            
            # Description
            param["description"] = st.text_area(
                "Description (optionnel)",
                value=param.get("description", ""),
                height=60,
                key=f"param_desc_{idx}"
            )
            
            # ✅ OPTIONS ENRICHIES (pour type select)
            if param["type"] in ["string", "liste"]:
                st.markdown("---")
                st.markdown("**💡 Options disponibles**")
                
                param["options_mode"] = st.radio(
                    "Source des options",
                    ["none", "manual", "from_column"],
                    index=["none", "manual", "from_column"].index(param.get("options_mode", "none")),
                    format_func=lambda x: {
                        "none": "Aucune (saisie libre)",
                        "manual": "Liste manuelle",
                        "from_column": "Depuis colonne de gabarit"
                    }[x],
                    key=f"param_optmode_{idx}",
                    horizontal=True
                )
                
                # MODE MANUEL
                if param["options_mode"] == "manual":
                    manual_text = st.text_area(
                        "Options (une par ligne)",
                        value="\n".join(param.get("options_manual", [])) if param.get("options_manual") else "",
                        height=100,
                        placeholder="BOMBAY\nGORDON'S\nTANQUERAY",
                        key=f"param_manual_{idx}"
                    )
                    param["options_manual"] = [line.strip() for line in manual_text.split("\n") if line.strip()]
                    param["options_source"] = None
                    
                    if param["options_manual"]:
                        st.caption(f"✓ {len(param['options_manual'])} option(s) définies")
                
                # MODE DEPUIS COLONNE
                elif param["options_mode"] == "from_column":
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Index par défaut si source existe déjà
                        current_source = param.get("options_source") or {}
                        current_gab_name = current_source.get("gabarit", "")
                        current_gab_ver = current_source.get("version", "v1")
                        
                        # Trouver l'index du gabarit actuel
                        default_index = 0
                        if current_gab_name:
                            for i, opt in enumerate(gabarit_options):
                                if opt != "(aucun)" and current_gab_name in opt and current_gab_ver in opt:
                                    default_index = i
                                    break
                        
                        selected_gab = st.selectbox(
                            "Gabarit source",
                            gabarit_options,
                            index=default_index,
                            key=f"param_gab_{idx}"
                        )
                    
                    # Récupérer les colonnes du gabarit sélectionné
                    column_options = ["(choisir)"]
                    if selected_gab != "(aucun)":
                        try:
                            gab_parts = selected_gab.split(" (v")
                            gab_name = gab_parts[0].strip()
                            gab_ver = gab_parts[1].rstrip(")").strip()
                            
                            gab_obj = get_gabarit(gab_name, gab_ver)
                            if gab_obj and gab_obj.columns:
                                column_options = ["(choisir)"] + [c.name for c in gab_obj.columns]
                        except Exception as e:
                            st.error(f"Erreur chargement colonnes : {e}")
                    
                    with col2:
                        # Index par défaut pour la colonne
                        current_col = current_source.get("column", "")
                        col_default_index = 0
                        if current_col and current_col in column_options:
                            col_default_index = column_options.index(current_col)
                        
                        selected_col = st.selectbox(
                            "Colonne source",
                            column_options,
                            index=col_default_index,
                            key=f"param_col_{idx}",
                            disabled=(selected_gab == "(aucun)")
                        )
                    
                        # Enregistrer la source
                        if selected_gab != "(aucun)" and selected_col != "(choisir)":
                            gab_parts = selected_gab.split(" (v")
                            gab_name = gab_parts[0].strip()
                            gab_ver = gab_parts[1].rstrip(")").strip()

                            param["options_source"] = {
                                "gabarit": gab_name,
                                "version": gab_ver,
                                "column": selected_col
                            }
                            param["options_manual"] = None

                            # === Calcul sur BOUTON + Mise en CACHE persistante ===
                            do_cache = st.button(
                                "⚡ Calculer & mettre en cache les propositions",
                                key=f"param_cache_{idx}",
                                use_container_width=True
                            )

                            if do_cache:
                                try:
                                    from backend.services.parameter_service import ParameterService, ParameterConfig
                                    tmp_param = ParameterConfig(**param)
                                    values = ParameterService._resolve_options_from_gabarit(tmp_param)

                                    if values:
                                        st.success(f"✓ {len(values)} valeur(s) unique(s) trouvées")
                                        with st.expander("Aperçu (jusqu'à 200)", expanded=False):
                                            st.write(", ".join(values[:200]))
                                    else:
                                        st.warning("Aucune valeur trouvée dans cette colonne")

                                    # 💾 PERSISTENCE immédiate dans le JSON
                                    from backend.services.template_service import TemplateService
                                    from backend.services.database_service import DatabaseService

                                    tid = template_id_to_edit if edit_mode else st.session_state.get("selected_template")
                                    if not tid:
                                        st.error("⚠️ Impossible de sauvegarder : template ID manquant")
                                    else:
                                        with DatabaseService.get_session() as db:
                                            svc = TemplateService(db)
                                            cfg = svc.get_config(tid)
                                            plist = list(cfg.get("parameters") or [])

                                            found = False
                                            for i, p in enumerate(plist):
                                                if p.get("name") == param.get("name"):
                                                    q = dict(p)
                                                    q["options_cache"] = {
                                                        "values": values[:],
                                                        "source": {"gabarit": gab_name, "version": gab_ver, "column": selected_col},
                                                    }
                                                    plist[i] = q
                                                    found = True
                                                    break
                                            if not found:
                                                newp = dict(param)
                                                newp["options_cache"] = {
                                                    "values": values[:],
                                                    "source": {"gabarit": gab_name, "version": gab_ver, "column": selected_col},
                                                }
                                                plist.append(newp)

                                            cfg["parameters"] = plist
                                            svc.update_config(tid, cfg)

                                        # ✅ CORRECTION : Mettre à jour le cache dans la session en cours
                                        param["options_cache"] = {
                                            "values": values[:],
                                            "source": {"gabarit": gab_name, "version": gab_ver, "column": selected_col},
                                        }
                                        
                                        st.toast("Propositions mises en cache ✅")
                                        st.rerun()  # ✅ Recharger pour afficher le selectbox
                                except Exception as e:
                                    st.error(f"Erreur cache : {e}")

                            # Affichage du cache existant (aucun calcul ici)
                            cache = param.get("options_cache")
                            if isinstance(cache, dict) and isinstance(cache.get("values"), list) and cache["values"]:
                                st.caption(f"Cache existant : {len(cache['values'])} valeur(s)")
                                with st.expander("Voir un aperçu (jusqu'à 200)", expanded=False):
                                    st.write(", ".join(list(cache["values"])[:200]))
                            else:
                                st.caption("Aucun cache enregistré pour ce paramètre.")

                
                # MODE AUCUNE
                else:
                    param["options_manual"] = None
                    param["options_source"] = None
            
            # === VALEUR PAR DÉFAUT ===
            st.markdown("---")
            st.markdown("**⚙️ Valeur par défaut**")

            mode = param.get("options_mode", "none")
            cache = param.get("options_cache")

            # 1️⃣ OPTIONS DEPUIS COLONNE (from_column)
            if mode == "from_column":
                # Vérifier si un cache existe
                has_cache = isinstance(cache, dict) and isinstance(cache.get("values"), list) and len(cache.get("values", [])) > 0
                
                if has_cache:
                    opts = list(cache["values"])
                    cur = param.get("default")
                    idx_default = opts.index(cur) if (cur in opts) else 0
                    param["default"] = st.selectbox(
                        "Valeur par défaut",
                        opts,
                        index=idx_default,
                        key=f"param_default_sel_{idx}",
                    )
                else:
                    # Aucun cache → afficher un avertissement + champ texte temporaire
                    st.info("⚠️ Aucun cache disponible. Cliquez d'abord sur '⚡ Calculer & mettre en cache les propositions'.")
                    param["default"] = st.text_input(
                        "Valeur par défaut (temporaire)",
                        value=str(param.get("default", "")) if param.get("default") else "",
                        key=f"param_default_nocache_{idx}",
                        help="Cette valeur sera utilisée en attendant le calcul du cache"
                    )

            # 2️⃣ OPTIONS MANUELLES
            elif mode == "manual":
                opts = [v for v in (param.get("options_manual") or []) if v]
                if opts:
                    cur = param.get("default")
                    idx_default = opts.index(cur) if (cur in opts) else 0
                    param["default"] = st.selectbox(
                        "Valeur par défaut",
                        opts,
                        index=idx_default,
                        key=f"param_default_manual_sel_{idx}",
                    )
                else:
                    param["default"] = st.text_input(
                        "Valeur par défaut",
                        value=str(param.get("default", "")) if param.get("default") else "",
                        key=f"param_default_manual_txt_{idx}",
                    )

            # 3️⃣ PAS D'OPTIONS → widgets selon le type
            else:
                if param.get("type") == "integer":
                    param["default"] = st.number_input(
                        "Valeur par défaut",
                        value=int(param.get("default", 0)) if param.get("default") is not None else 0,
                        key=f"param_default_int_{idx}",
                    )
                elif param.get("type") == "date":
                    from datetime import datetime, date
                    default_date = param.get("default")
                    if isinstance(default_date, str):
                        try:
                            default_date = datetime.fromisoformat(default_date).date()
                        except Exception:
                            default_date = date.today()
                    elif not isinstance(default_date, date):
                        default_date = date.today()
                    param["default"] = st.date_input(
                        "Valeur par défaut",
                        value=default_date,
                        key=f"param_default_date_{idx}",
                    ).isoformat()
                else:  # string ou liste sans options
                    param["default"] = st.text_input(
                        "Valeur par défaut",
                        value=str(param.get("default", "")) if param.get("default") else "",
                        key=f"param_default_text_{idx}",
                    )
            
            # Bouton suppression (UN SEUL, à la fin de l'expander)
            st.markdown("---")
            if st.button("🗑️ Supprimer ce paramètre", key=f"param_del_{idx}", use_container_width=True):
                to_delete.append(idx)
    
    # Traiter les suppressions
    if to_delete:
        for i in sorted(to_delete, reverse=True):
            del st.session_state.parameters[i]
        st.rerun()

    st.divider()

# ===== ÉTAPE 4 : Boucles =====
if show_loops:
    st.header("4️⃣ Boucles (édition en tableau)")
    st.markdown("`slides` doit être une liste de codes séparés par des virgules (ex: A001, A002)")

    loops_norm = []
    for loop in (st.session_state.loops or []):
        loops_norm.append({
            "loop_id": loop.get("loop_id", ""),
            "slides": ", ".join(loop.get("slides", [])) if isinstance(loop.get("slides"), list) else str(loop.get("slides", "")),
            "sheet_name": loop.get("sheet_name", "Boucles"),
        })
    loops_df = pd.DataFrame(loops_norm or [{"loop_id": "", "slides": "", "sheet_name": "Boucles"}])

    loops_editor = st.data_editor(
        loops_df,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "loop_id": st.column_config.TextColumn("ID boucle", help="Identifiant utilisé côté code"),
            "slides": st.column_config.TextColumn("Slides (A001, A002, ...)"),
            "sheet_name": st.column_config.TextColumn("Feuille Excel Loop", help="Par défaut: Boucles"),
        }
    )

    def _split_slides(s: str) -> list[str]:
        return [x.strip() for x in str(s or "").split(",") if x.strip()]

    st.session_state.loops = [
        {
            "loop_id": str(row.get("loop_id", "")).strip(),
            "slides": _split_slides(row.get("slides", "")),
            "sheet_name": str(row.get("sheet_name", "") or "Boucles").strip(),
        }
        for _, row in loops_editor.iterrows()
        if str(row.get("loop_id", "")).strip()
    ]

    st.divider()

# ===== ÉTAPE 5 : Configuration des images =====
if show_images:
    st.header("5️⃣ Configuration des images dynamiques")

    st.markdown("""
    Configurez les images à injecter dynamiquement dans les slides (logos, photos produits, fonds...).
    """)

    with st.expander("➕ Ajouter une configuration d'image"):
        col1, col2 = st.columns(2)
        
        with col1:
            img_slide_id = st.text_input("Slide ID", placeholder="ex: A001", key="new_img_slide")
            img_type = st.text_input("Type d'image", placeholder="ex: product_image", key="new_img_type")
            img_pattern = st.text_input("Pattern du chemin", placeholder="assets/{Marque}/{Produit}.png", key="new_img_pattern")
        
        with col2:
            img_default = st.text_input("Image par défaut (optionnel)", key="new_img_default")
            img_background = st.checkbox("Placer en arrière-plan", key="new_img_bg")
            img_loop = st.checkbox("Dépend d'une boucle", key="new_img_loop")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            img_left = st.number_input("Position Left", value=10, key="new_img_left")
        with col2:
            img_top = st.number_input("Position Top", value=10, key="new_img_top")
        with col3:
            img_max_width = st.number_input("Largeur max", value=100, key="new_img_width")
        with col4:
            img_max_height = st.number_input("Hauteur max", value=100, key="new_img_height")
        
        if st.button("➕ Ajouter cette image"):
            if img_slide_id and img_pattern:
                if img_slide_id not in st.session_state.images:
                    st.session_state.images[img_slide_id] = []
                
                st.session_state.images[img_slide_id].append({
                    "type": img_type,
                    "pattern": img_pattern,
                    "default_path": img_default if img_default else None,
                    "position": {"left": img_left, "top": img_top},
                    "size": {"max_width": img_max_width, "max_height": img_max_height},
                    "background": img_background,
                    "loop_dependent": img_loop
                })
                st.rerun()

    if st.session_state.images:
        st.markdown("**Images configurées :**")
        for slide_id, images in st.session_state.images.items():
            with st.expander(f"Slide {slide_id} ({len(images)} image(s))"):
                for idx, img in enumerate(images):
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.text(f"• {img['type']} : {img['pattern']}")
                    with col2:
                        if st.button("🗑️", key=f"del_img_{slide_id}_{idx}"):
                            st.session_state.images[slide_id].pop(idx)
                            if not st.session_state.images[slide_id]:
                                del st.session_state.images[slide_id]
                            st.rerun()

    st.divider()

# ===== ÉTAPE 6 : Mappings =====
if show_mappings:
    st.header("6️⃣ Mappings (édition en tableau)")

    mappings_df = pd.DataFrame(st.session_state.mappings or [],
                            columns=["slide_id", "sheet_name", "excel_range", "has_header"])

    if mappings_df.empty:
        mappings_df = pd.DataFrame([{"slide_id": "", "sheet_name": "Table", "excel_range": "A1:D10", "has_header": True}])

    mappings_editor = st.data_editor(
        mappings_df,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "slide_id": st.column_config.TextColumn("Slide ID", help="ex: A001"),
            "sheet_name": st.column_config.TextColumn("Feuille Excel", help="ex: Table"),
            "excel_range": st.column_config.TextColumn("Plage", help="ex: A1:D10"),
            "has_header": st.column_config.CheckboxColumn("En-tête"),
        }
    )
    st.session_state.mappings = [
        {
            "slide_id": str(row.get("slide_id", "")).strip(),
            "sheet_name": str(row.get("sheet_name", "") or "Table").strip(),
            "excel_range": str(row.get("excel_range", "") or "A1:D10").strip(),
            "has_header": bool(row.get("has_header", True)),
        }
        for _, row in mappings_editor.iterrows()
        if str(row.get("slide_id", "")).strip() and str(row.get("excel_range", "")).strip()
    ]

    st.divider()

# ===== BOUTON DE GÉNÉRATION =====
button_label = "💾 Mettre à jour le template" if edit_mode else "🚀 Créer le template"

if st.button(button_label, type="primary", use_container_width=True):
    if not name:
        st.error("Le nom du template est obligatoire")
    elif upload_mode == "Uploader des fichiers existants" and (not ppt_file or not excel_file) and not edit_mode:
        st.error("Uploadez les fichiers PowerPoint et Excel")
    else:
        try:
            # Normalisation avant TemplateConfig
            for loop in st.session_state.loops:
                if isinstance(loop.get("slides"), str):
                    loop["slides"] = [x.strip() for x in loop["slides"].split(",") if x.strip()]

            config = TemplateConfig(
                name=name,
                version=version,
                description=description,
                parameters=[ParameterConfig(**p) for p in st.session_state.parameters],
                data_source=DataSourceConfig(
                    type="excel",
                    required_tables=[]
                ),

                # ✅ SAUVEGARDE DES TAGS
                tags=TemplateTags(
                    themes=st.session_state[THEMES_KEY],
                    families=st.session_state[FAMS_KEY]
                ),


                loops=[LoopConfig(**loop) for loop in st.session_state.loops],
                image_injections={
                    slide_id: [ImageInjection(**img) for img in images]
                    for slide_id, images in st.session_state.images.items()
                },
                slide_mappings=[SlideMapping(**m) for m in st.session_state.mappings]
            )

            
            if edit_mode:
                # MODE MISE À JOUR
                with DatabaseService.get_session() as db:
                    service = TemplateService(db)
                    
                    updates = {
                        'version': version,
                        'description': description,
                        'config': config.model_dump(mode='json')
                    }
                    
                    service.update_template(
                        template_id=template_id_to_edit,
                        updates=updates,
                        user_id=1
                    )

                    # ✅ ÉCRITURE DISQUE DU config.json (source de vérité pour la Bibliothèque)
                    service.update_config(
                        template_id=template_id_to_edit,
                        new_config=config.model_dump(mode='json')
                    )

                    if card_image is not None:
                        service.save_card_image(
                            template_id=template_id_to_edit,
                            file_bytes=card_image.getvalue(),
                            original_filename=card_image.name
                        )

                
                st.session_state["_flash_success"] = f"✅ Template '{name}' mis à jour !"
                st.session_state.selected_template_detail = template_id_to_edit
                st.switch_page("pages/3_📚_Bibliotheque.py")
                st.stop()

                
                if '_template_loaded' in st.session_state:
                    del st.session_state._template_loaded
            
            else:
                # MODE CRÉATION
                import tempfile
                from PIL import Image

                temp_dir = Path(tempfile.gettempdir())
                ppt_path = None
                excel_path = None

                if upload_mode == "Uploader des fichiers existants":
                    if ppt_file:
                        ppt_path = temp_dir / ppt_file.name
                        with open(ppt_path, 'wb') as f:
                            f.write(ppt_file.getbuffer())
                    if excel_file:
                        excel_path = temp_dir / excel_file.name
                        with open(excel_path, 'wb') as f:
                            f.write(excel_file.getbuffer())
                else:
                    master_excel_path = project_root / "assets" / "master" / "master_template.xlsx"
                    master_ppt_path = project_root / "assets" / "master" / "master_template.pptx"
                    
                    if not master_excel_path.exists():
                        st.error(f"Master Excel introuvable : {master_excel_path}")
                        st.stop()
                    
                    if not master_ppt_path.exists():
                        st.error(f"Master PowerPoint introuvable : {master_ppt_path}")
                        st.stop()
                    
                    excel_path = master_excel_path
                    ppt_path = master_ppt_path

                with DatabaseService.get_session() as db:
                    service = TemplateService(db)
                    template = service.create_template(
                        config=config,
                        user_id=1,
                        ppt_source=ppt_path,
                        excel_source=excel_path
                    )
                    template_id = template.id

                    if card_image:
                        assets_dir = project_root / "assets" / "background" / "card"
                        assets_dir.mkdir(parents=True, exist_ok=True)

                        img = Image.open(card_image)
                        target_width, target_height = 300, 150
                        img_ratio = img.width / img.height
                        target_ratio = target_width / target_height

                        if img_ratio > target_ratio:
                            new_height = img.height
                            new_width = int(new_height * target_ratio)
                            left = (img.width - new_width) // 2
                            img_cropped = img.crop((left, 0, left + new_width, new_height))
                        else:
                            new_width = img.width
                            new_height = int(new_width / target_ratio)
                            top = (img.height - new_height) // 2
                            img_cropped = img.crop((0, top, new_width, top + new_height))

                        img_final = img_cropped.resize((target_width, target_height), Image.Resampling.LANCZOS)
                        image_path = assets_dir / f"{name}.png"
                        img_final.save(image_path, "PNG")

                        template.card_image_path = str(image_path)
                        db.commit()

                st.success(f"✅ Template '{name}' créé! (ID: {template_id})")
                st.balloons()
                
                st.session_state.parameters = []
                st.session_state.loops = []
                st.session_state.images = {}
                st.session_state.mappings = []

                st.session_state["_flash_success"] = f"✅ Template '{name}' créé ! (ID: {template_id})"
                st.session_state.selected_template_detail = template_id
                st.switch_page("pages/3_📚_Bibliotheque.py")
                st.stop()

        
        except Exception as e:
            st.error(f"Erreur : {e}")
            import traceback
            st.code(traceback.format_exc())