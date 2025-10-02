"""
Page de création de templates - VERSION COMPLÈTE
"""

import streamlit as st
from pathlib import Path
import sys
import json

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

st.set_page_config(page_title="Nouveau Template", page_icon="➕", layout="wide")

st.title("➕ Créer un nouveau template")

# Initialiser les états de session
if 'parameters' not in st.session_state:
    st.session_state.parameters = []
if 'loops' not in st.session_state:
    st.session_state.loops = []
if 'images' not in st.session_state:
    st.session_state.images = {}
if 'mappings' not in st.session_state:
    st.session_state.mappings = []

# ===== ÉTAPE 1 : Informations générales =====
st.header("1️⃣ Informations générales")

col1, col2 = st.columns(2)

with col1:
    name = st.text_input("Nom du template*", placeholder="ex: BCE_INSEE")
    version = st.text_input("Version", value="1.0")

with col2:
    description = st.text_area("Description", placeholder="Description du template...")

st.divider()

# ===== ÉTAPE 2 : Fichiers sources =====
st.header("2️⃣ Fichiers sources")

upload_mode = st.radio(
    "Mode de création",
    ["Créer des fichiers vierges", "Uploader des fichiers existants"]
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

# ===== ÉTAPE 3 : Paramètres =====
st.header("3️⃣ Paramètres")

with st.expander("➕ Ajouter un paramètre", expanded=len(st.session_state.parameters) == 0):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        param_name = st.text_input("Nom du paramètre", key="new_param_name")
    with col2:
        param_type = st.selectbox("Type", ["string", "integer", "date", "list"], key="new_param_type")
    with col3:
        param_balise = st.text_input("Balise PPT", placeholder="[NomParametre]", key="new_param_balise")
    
    if st.button("➕ Ajouter ce paramètre"):
        if param_name and param_balise:
            st.session_state.parameters.append({
                "name": param_name,
                "type": param_type,
                "required": True,
                "balise_ppt": param_balise
            })
            st.rerun()
        else:
            st.error("Nom et balise obligatoires")

# Afficher les paramètres
if st.session_state.parameters:
    st.markdown("**Paramètres configurés :**")
    for idx, param in enumerate(st.session_state.parameters):
        col1, col2 = st.columns([4, 1])
        with col1:
            st.text(f"• {param['name']} ({param['type']}) → {param['balise_ppt']}")
        with col2:
            if st.button("🗑️", key=f"del_param_{idx}"):
                st.session_state.parameters.pop(idx)
                st.rerun()

st.divider()

# ===== ÉTAPE 4 : Source de données =====
st.header("4️⃣ Source de données")

data_source_type = st.selectbox(
    "Type de source",
    ["excel", "postgresql", "mysql", "csv"]
)

if data_source_type == "excel":
    st.info("💡 Les tableaux structurés du fichier Excel seront utilisés comme source")
    table_names = st.text_input(
        "Noms des tableaux Excel",
        value="Performance",
        help="Noms des tableaux structurés (séparés par des virgules)"
    )
    tables_list = [t.strip() for t in table_names.split(',') if t.strip()]
else:
    required_tables = st.text_area(
        "Tables requises (une par ligne)",
        placeholder="observations\ndim_produits"
    )
    tables_list = [t.strip() for t in required_tables.split('\n') if t.strip()]

st.divider()

# ===== ÉTAPE 5 : Configuration des boucles =====
st.header("5️⃣ Configuration des boucles")

st.markdown("""
Les boucles permettent de dupliquer automatiquement des slides selon un tableau Loop dans Excel.
""")

with st.expander("➕ Ajouter une boucle", expanded=len(st.session_state.loops) == 0):
    col1, col2 = st.columns(2)
    
    with col1:
        loop_id = st.text_input("ID de la boucle", placeholder="ex: Entreprise", key="new_loop_id")
        loop_sheet = st.text_input("Feuille Excel", value="Charts_settings", key="new_loop_sheet")
    
    with col2:
        loop_slides = st.text_input("Slides concernées (séparées par virgules)", placeholder="A001, A002", key="new_loop_slides")
    
    if st.button("➕ Ajouter cette boucle"):
        if loop_id and loop_slides:
            slides_list = [s.strip() for s in loop_slides.split(',')]
            st.session_state.loops.append({
                "loop_id": loop_id,
                "slides": slides_list,
                "sheet_name": loop_sheet
            })
            st.rerun()

# Afficher les boucles
if st.session_state.loops:
    st.markdown("**Boucles configurées :**")
    for idx, loop in enumerate(st.session_state.loops):
        col1, col2 = st.columns([4, 1])
        with col1:
            st.text(f"• {loop['loop_id']} → Slides: {', '.join(loop['slides'])}")
        with col2:
            if st.button("🗑️", key=f"del_loop_{idx}"):
                st.session_state.loops.pop(idx)
                st.rerun()

st.divider()

# ===== ÉTAPE 6 : Configuration des images =====
st.header("6️⃣ Configuration des images dynamiques")

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

# ===== ÉTAPE 7 : Mapping des données =====
st.header("7️⃣ Mapping des données (optionnel)")

st.markdown("""
Configurez quelles données Excel doivent être injectées dans quels tableaux PowerPoint.
""")

with st.expander("➕ Ajouter un mapping de données"):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        map_slide_id = st.text_input("Slide ID", placeholder="A001", key="new_map_slide")
    with col2:
        map_sheet = st.text_input("Feuille Excel", placeholder="Table", key="new_map_sheet")
    with col3:
        map_range = st.text_input("Plage Excel", placeholder="A1:D10", key="new_map_range")
    
    map_header = st.checkbox("Première ligne = en-tête", value=True, key="new_map_header")
    
    if st.button("➕ Ajouter ce mapping"):
        if map_slide_id and map_sheet and map_range:
            st.session_state.mappings.append({
                "slide_id": map_slide_id,
                "sheet_name": map_sheet,
                "excel_range": map_range,
                "has_header": map_header
            })
            st.rerun()

# Afficher les mappings
if st.session_state.mappings:
    st.markdown("**Mappings configurés :**")
    for idx, mapping in enumerate(st.session_state.mappings):
        col1, col2 = st.columns([4, 1])
        with col1:
            st.text(f"• {mapping['slide_id']} ← {mapping['sheet_name']}!{mapping['excel_range']}")
        with col2:
            if st.button("🗑️", key=f"del_map_{idx}"):
                st.session_state.mappings.pop(idx)
                st.rerun()

st.divider()

# ===== BOUTON DE GÉNÉRATION =====
if st.button("🚀 Créer le template", type="primary", use_container_width=True):
    if not name:
        st.error("Le nom du template est obligatoire")
    elif not st.session_state.parameters:
        st.error("Ajoutez au moins un paramètre")
    elif upload_mode == "Uploader des fichiers existants" and (not ppt_file or not excel_file):
        st.error("Uploadez les fichiers PowerPoint et Excel")
    else:
        try:
            # Construire la config
            config = TemplateConfig(
                name=name,
                version=version,
                description=description,
                parameters=[ParameterConfig(**p) for p in st.session_state.parameters],
                data_source=DataSourceConfig(
                    type=data_source_type,
                    required_tables=tables_list
                ),
                loops=[LoopConfig(**loop) for loop in st.session_state.loops],
                image_injections={
                    slide_id: [ImageInjection(**img) for img in images]
                    for slide_id, images in st.session_state.images.items()
                },
                slide_mappings=[SlideMapping(**m) for m in st.session_state.mappings]
            )
            
            # Sauvegarder les fichiers uploadés
            import tempfile
            temp_dir = Path(tempfile.gettempdir())
            
            ppt_path = None
            excel_path = None
            
            if ppt_file:
                ppt_path = temp_dir / ppt_file.name
                with open(ppt_path, 'wb') as f:
                    f.write(ppt_file.getbuffer())
            
            if excel_file:
                excel_path = temp_dir / excel_file.name
                with open(excel_path, 'wb') as f:
                    f.write(excel_file.getbuffer())
            
            # Créer le template
            template_name = None
            template_id = None
            
            with DatabaseService.get_session() as db:
                service = TemplateService(db)
                template = service.create_template(
                    config=config,
                    user_id=1,
                    ppt_source=ppt_path,
                    excel_source=excel_path
                )
                template_name = template.name
                template_id = template.id
            
            st.success(f"✅ Template '{template_name}' créé avec succès! (ID: {template_id})")
            st.balloons()
            
            # Réinitialiser
            st.session_state.parameters = []
            st.session_state.loops = []
            st.session_state.images = {}
            st.session_state.mappings = []
            
        except Exception as e:
            st.error(f"Erreur lors de la création : {e}")
            import traceback
            st.code(traceback.format_exc())