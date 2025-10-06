"""
Page de création de templates - VERSION COMPLÈTE
"""

import streamlit as st
from pathlib import Path
import sys
import json
import pandas as pd
from io import BytesIO
from pathlib import Path


project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.models.template_config import (
    TemplateConfig,
    ParameterConfig,
    DataSourceConfig,
    LoopConfig,
    ImageInjection,
    SlideMapping
)

st.set_page_config(page_title="Paramètres généraux", page_icon="⚙️", layout="wide")
def render_template_subnav(active: str, template_id: int | None):
    cols = st.columns([1,1,1,1,1])
    with cols[0]:
        if st.button("← Retour bibliothèque", use_container_width=True):
            if template_id:
                st.session_state.selected_template_detail = template_id
            if "selected_template" in st.session_state:
                del st.session_state.selected_template
            st.switch_page("pages/2_📚_Bibliotheque.py")
    with cols[1]:
        if st.button("🗂️ Détail du template", type=("primary" if active=="detail" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template_detail = template_id
            st.switch_page("pages/_2a_📊_Detail_Livrable.py")
    with cols[2]:
        st.button("⚙️ Paramètres généraux", type="primary" if active=="general" else "secondary", use_container_width=True)
    with cols[3]:
        if st.button("📑 Injection des données", type=("primary" if active=="inject" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_2b3_📑_Tables_Template.py")
    with cols[4]:
        if st.button("🧾 Ajustement de la table", type=("primary" if active=="adjust" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_2b4_🧾_Ajustement_Table.py")
    st.divider()

# --- États init (inchangés)
if 'parameters' not in st.session_state:
    st.session_state.parameters = []
if 'loops' not in st.session_state:
    st.session_state.loops = []
if 'images' not in st.session_state:
    st.session_state.images = {}
if 'mappings' not in st.session_state:
    st.session_state.mappings = []

# --- Guard + chargement
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
        st.session_state.parameters = [
            {"name": p.name, "type": p.type, "required": p.required, "balise_ppt": p.balise_ppt}
            for p in template_config.parameters
        ]
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


# === Récap compact dans la sidebar ===
with st.sidebar:
    st.header("🧭 Récap")
    st.write(f"**Mode** : {'Édition' if edit_mode else 'Création'}")
    st.write(f"**Nom** : {template_name or '—'}")
    st.write(f"**Version** : {template_version if edit_mode else '—'}")
    st.write(f"**Source** : {template_config.data_source.type if edit_mode else '—'}")

    # Aperçu image carte
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
        value=template_name if edit_mode else "",  # Utiliser template_name extrait
        placeholder="ex: BCE_INSEE",
        disabled=edit_mode
    )
    version = st.text_input(
        "Version", 
        value=template_version if edit_mode else "1.0"  # Utiliser template_version extrait
    )

with col2:
    description = st.text_area(
        "Description", 
        value=template_description if edit_mode and template_description else "",  # Utiliser template_description extrait
        placeholder="Description du template..."
    )

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
    # Upload image de carte
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
        value=False,
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



# ===== ÉTAPE 3 : Paramètres =====
if show_params:
    st.header("3️⃣ Paramètres")
    st.markdown("Ajoute/modifie directement dans le tableau. Les lignes vides sont ignorées.")

    # DataFrame source
    param_df = pd.DataFrame(st.session_state.parameters or [],
                            columns=["name", "type", "required", "balise_ppt"])

    # Valeurs par défaut si vide
    if param_df.empty:
        param_df = pd.DataFrame([{"name": "", "type": "string", "required": True, "balise_ppt": ""}])

    param_editor = st.data_editor(
        param_df,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "name": st.column_config.TextColumn("Nom", help="Nom interne du paramètre (ex: sous_marque)"),
            "type": st.column_config.SelectboxColumn("Type", options=["string", "integer", "date", "list"]),
            "required": st.column_config.CheckboxColumn("Obligatoire"),
            "balise_ppt": st.column_config.TextColumn("Balise PPT", help="ex: [SousMarque]"),
        }
    )

    # Sauvegarde dans la session (en nettoyant les lignes vides)
    st.session_state.parameters = [
        {
            "name": str(row.get("name", "")).strip(),
            "type": row.get("type") or "string",
            "required": bool(row.get("required", True)),
            "balise_ppt": str(row.get("balise_ppt", "")).strip()
        }
        for _, row in param_editor.iterrows()
        if str(row.get("name", "")).strip() and str(row.get("balise_ppt", "")).strip()
    ]


    st.divider()

# ===== ÉTAPE 4 : Boucles =====
if show_loops:
    st.header("4️⃣ Boucles (édition en tableau)")
    st.markdown("`slides` doit être une liste de codes séparés par des virgules (ex: A001, A002)")

    # Convertit les boucles actuelles en DF
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

    # Sauvegarde dans la session en listifiant slides
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

    # Afficher les images
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
                    type="excel",  # Valeur par défaut
                    required_tables=[]  # Liste vide
                ),
                loops=[LoopConfig(**loop) for loop in st.session_state.loops],
                image_injections={
                    slide_id: [ImageInjection(**img) for img in images]
                    for slide_id, images in st.session_state.images.items()
                },
                slide_mappings=[SlideMapping(**m) for m in st.session_state.mappings]
            )
            
            if edit_mode:
                # ========== MODE MISE À JOUR ==========
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

                    # Enregistrer nouvelle image si fournie
                    if card_image is not None:
                        service.save_card_image(
                            template_id=template_id_to_edit,
                            file_bytes=card_image.getvalue(),
                            original_filename=card_image.name
                        )
                
                st.success(f"✅ Template '{name}' mis à jour!")
                st.info("💡 Fichiers masters inchangés. Pour modifier PPT/Excel, créez une nouvelle version.")
                
                if '_template_loaded' in st.session_state:
                    del st.session_state._template_loaded
            
            else:
                # ========== MODE CRÉATION ==========
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
                    # Fichiers vierges : utiliser masters par défaut
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

                # Création template
                with DatabaseService.get_session() as db:
                    service = TemplateService(db)
                    template = service.create_template(
                        config=config,
                        user_id=1,
                        ppt_source=ppt_path,
                        excel_source=excel_path
                    )
                    template_id = template.id

                    # Gérer image de carte
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
                
                # Réinitialiser
                st.session_state.parameters = []
                st.session_state.loops = []
                st.session_state.images = {}
                st.session_state.mappings = []
        
        except Exception as e:
            st.error(f"Erreur : {e}")
            import traceback
            st.code(traceback.format_exc())