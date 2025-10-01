"""
Page de création de templates
"""

import streamlit as st
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.models.template_config import (
    TemplateConfig,
    ParameterConfig,
    DataSourceConfig,
    LoopConfig
)

st.set_page_config(page_title="Nouveau Template", page_icon="➕", layout="wide")

st.title("➕ Créer un nouveau template")

# Étape 1 : Informations générales
st.header("1️⃣ Informations générales")

col1, col2 = st.columns(2)

with col1:
    name = st.text_input("Nom du template*", placeholder="ex: Suivi_Commercial")
    version = st.text_input("Version", value="1.0")

with col2:
    description = st.text_area("Description", placeholder="Description du template...")

st.divider()

# Étape 2 : Fichiers sources
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
        ppt_file = st.file_uploader("Template PowerPoint", type=['pptx'])
    
    with col2:
        excel_file = st.file_uploader("Template Excel", type=['xlsx'])

st.divider()

# Étape 3 : Paramètres
st.header("3️⃣ Paramètres")

if 'parameters' not in st.session_state:
    st.session_state.parameters = []

col1, col2 = st.columns([3, 1])

with col1:
    param_name = st.text_input("Nom du paramètre", key="param_name")
    param_type = st.selectbox("Type", ["string", "integer", "date", "list"], key="param_type")
    param_balise = st.text_input("Balise PPT", placeholder="ex: [Marque]", key="param_balise")

with col2:
    st.write("")
    st.write("")
    if st.button("➕ Ajouter paramètre"):
        if param_name and param_balise:
            st.session_state.parameters.append({
                "name": param_name,
                "type": param_type,
                "required": True,
                "balise_ppt": param_balise
            })
            st.success(f"Paramètre '{param_name}' ajouté")
            st.rerun()

# Afficher les paramètres
if st.session_state.parameters:
    st.markdown("**Paramètres ajoutés :**")
    for idx, param in enumerate(st.session_state.parameters):
        col1, col2 = st.columns([4, 1])
        with col1:
            st.text(f"• {param['name']} ({param['type']}) → {param['balise_ppt']}")
        with col2:
            if st.button("🗑️", key=f"del_param_{idx}"):
                st.session_state.parameters.pop(idx)
                st.rerun()

st.divider()

# Étape 4 : Source de données
st.header("4️⃣ Source de données")

data_source_type = st.selectbox(
    "Type de source",
    ["postgresql", "mysql", "excel", "csv"]
)

if data_source_type in ["postgresql", "mysql"]:
    required_tables = st.text_area(
        "Tables requises (une par ligne)",
        placeholder="observations\ndim_produits\ndim_drives"
    )
    tables_list = [t.strip() for t in required_tables.split('\n') if t.strip()]
else:
    tables_list = []

st.divider()

# Bouton de génération
if st.button("🚀 Créer le template", type="primary", use_container_width=True):
    if not name:
        st.error("Le nom du template est obligatoire")
    elif not st.session_state.parameters:
        st.error("Ajoutez au moins un paramètre")
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
                )
            )
            
            # Sauvegarder les fichiers uploadés
            ppt_path = None
            excel_path = None
            
            if ppt_file:
                import tempfile
                temp_dir = Path(tempfile.gettempdir())
                ppt_path = temp_dir / ppt_file.name
                with open(ppt_path, 'wb') as f:
                    f.write(ppt_file.getbuffer())
            
            if excel_file:
                import tempfile
                temp_dir = Path(tempfile.gettempdir())
                excel_path = temp_dir / excel_file.name
                with open(excel_path, 'wb') as f:
                    f.write(excel_file.getbuffer())
            
            # Créer le template
            template_name = None
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
            
        except Exception as e:
            st.error(f"Erreur lors de la création : {e}")