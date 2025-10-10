# frontend/pages/2_📚_Bibliotheque.py
# (Ancien: 1_📚_Bibliotheque.py)
# CHANGEMENTS: 
# - Numérotation 2 au lieu de 1
# - Navigation mise à jour vers nouvelles pages

"""
Page de la bibliothèque de templates
"""
from PIL import Image, ImageOps

import streamlit as st
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
import base64


st.set_page_config(page_title="Bibliothèque", page_icon="📚", layout="wide")

st.title("📚 Bibliothèque de Templates")
if msg := st.session_state.pop("_flash_success", None):
    st.success(msg)
    st.toast(msg)

# Filtres
col1, col2 = st.columns([3, 1])

with col1:
    search = st.text_input("🔍 Rechercher", placeholder="Nom du template...")

with col2:
    show_inactive = st.checkbox("Afficher inactifs", value=False)

st.divider()

colr1, colr2, colr3 = st.columns([1,1,6])
with colr1:
    if st.button("🔄 Rafraîchir"):
        st.rerun()


# Charger les templates
with DatabaseService.get_session() as db:
    service = TemplateService(db)
    templates = service.list_templates(active_only=not show_inactive)
    
    # Extraire toutes les infos dans la session
    templates_data = []
    for t in templates:
        templates_data.append({
            'id': t.id,
            'name': t.name,
            'version': t.version,
            'description': t.description,
            'ppt_path': t.ppt_template_path,
            'card_image_path': t.card_image_path,
            'is_active': t.is_active
        })

# Filtrer par recherche
if search:
    templates_data = [t for t in templates_data if search.lower() in t['name'].lower()]

# Affichage
if not templates_data:
    st.info("Aucun template trouvé. Créez-en un pour démarrer.")
    if st.button("➕ Nouveau template", type="primary", use_container_width=True):
        st.session_state.selected_template = None
        st.switch_page("pages/_2b_➕_Form_Template.py")
else:
    # Action en haut
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown(f"**{len(templates_data)} template(s) trouvé(s)**")
    with col2:
        if st.button("➕ Nouveau template", type="primary", use_container_width=True):
            st.session_state.selected_template = None
            st.switch_page("pages/_2b_➕_Form_Template.py")
    
    st.markdown("")
    
    # Grille de cartes (3 par ligne)
    cols_per_row = 3
    
    for i in range(0, len(templates_data), cols_per_row):
        cols = st.columns(cols_per_row)
        
        for j, col in enumerate(cols):
            idx = i + j
            if idx < len(templates_data):
                template = templates_data[idx]
                
                with col:
                    # Container pour la carte
                    with st.container(border=True):
                        # En-tête : nom + version
                        st.markdown(f"### {template['name']}")
                        st.caption(f"Version {template['version']}")
                        
                        # Image de carte
                        default_image = project_root / "assets" / "background" / "card" / "default.png"

                        image_to_show = None
                        if template['card_image_path'] and Path(template['card_image_path']).exists():
                            image_to_show = template['card_image_path']
                        elif default_image.exists():
                            image_to_show = str(default_image)

                        def afficher_image_carte(path: str, ratio: float = 16/9, radius_px: int = 8):
                            try:
                                p = Path(path)
                                if not p.exists():
                                    p = default_image
                                b64 = base64.b64encode(p.read_bytes()).decode("ascii")
                                padding_pct = 100 / ratio
                                st.markdown(f"""
                                <div style="position:relative;width:100%;padding-top:{padding_pct}%;
                                            overflow:hidden;border-radius:{radius_px}px;background:#10182014;">
                                <img src="data:image/png;base64,{b64}"
                                    style="position:absolute;inset:0;width:100%;height:100%;
                                            object-fit:cover;display:block;border-radius:{radius_px}px;">
                                </div>
                                """, unsafe_allow_html=True)
                            except Exception as e:
                                try:
                                    b64 = base64.b64encode(Path(default_image).read_bytes()).decode("ascii")
                                    st.markdown(f"""
                                    <div style="position:relative;width:100%;padding-top:{100/(16/9)}%;
                                                overflow:hidden;border-radius:8px;background:#10182014;">
                                    <img src="data:image/png;base64,{b64}"
                                        style="position:absolute;inset:0;width:100%;height:100%;
                                                object-fit:cover;display:block;border-radius:8px;">
                                    </div>
                                    """, unsafe_allow_html=True)
                                except:
                                    st.markdown("🖼️ *Aucune image*")

                        if image_to_show:
                            afficher_image_carte(image_to_show, ratio=16/9)
                        else:
                            st.markdown("🖼️ *Aucune image*")

                        
                        # Description (limitée à 100 caractères)
                        desc = template['description'] or "Aucune description"
                        if len(desc) > 100:
                            desc = desc[:97] + "..."
                        st.markdown(desc)
                        
                        st.markdown("")
                        
                        # Boutons
                        col_btn1, col_btn2 = st.columns(2)
                        
                        with col_btn1:
                            if st.button("📊 Ouvrir", key=f"open_{template['id']}", use_container_width=True):
                                st.session_state.selected_template_detail = template['id']
                                st.switch_page("pages/_2a_📊_Detail_Livrable.py")
                        
                        with col_btn2:
                            if st.button("🗑️ Supprimer", key=f"del_{template['id']}", 
                                       use_container_width=True, type="secondary"):
                                st.session_state.delete_template_id = template['id']
                                st.session_state.show_delete_modal = True
                                st.rerun()

st.divider()

# ===== MODAL DE CONFIRMATION SUPPRESSION =====
if st.session_state.get('show_delete_modal'):
    @st.dialog("⚠️ Confirmer la suppression")
    def confirm_delete():
        template_id = st.session_state.get('delete_template_id')
        
        # Récupérer le nom du template
        template_to_delete = next((t for t in templates_data if t['id'] == template_id), None)
        
        if template_to_delete:
            st.warning(f"**Vous êtes sur le point de supprimer le template :**")
            st.markdown(f"### {template_to_delete['name']} (v{template_to_delete['version']})")
            
            st.divider()
            st.markdown("**Cette action est irréversible.** Tapez le nom exact du template pour confirmer :")
            
            confirmation = st.text_input(
                "Nom du template",
                key="delete_confirm_input",
                placeholder=template_to_delete['name']
            )
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("Annuler", use_container_width=True):
                    st.session_state.show_delete_modal = False
                    if 'delete_template_id' in st.session_state:
                        del st.session_state.delete_template_id
                    st.rerun()
            
            with col2:
                if st.button("Supprimer définitivement", type="primary", use_container_width=True):
                    if confirmation != template_to_delete['name']:
                        st.error("❌ Le nom ne correspond pas")
                    else:
                        try:
                            with DatabaseService.get_session() as db:
                                service = TemplateService(db)
                                service.delete_template(template_id)
                            
                            st.success(f"✅ Template '{template_to_delete['name']}' supprimé")
                            st.session_state.show_delete_modal = False
                            if 'delete_template_id' in st.session_state:
                                del st.session_state.delete_template_id
                            st.rerun()
                        
                        except Exception as e:
                            st.error(f"❌ Erreur lors de la suppression : {e}")
    
    confirm_delete()