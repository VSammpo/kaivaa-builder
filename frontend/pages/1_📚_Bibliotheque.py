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

# Filtres
col1, col2 = st.columns([3, 1])

with col1:
    search = st.text_input("🔍 Rechercher", placeholder="Nom du template...")

with col2:
    show_inactive = st.checkbox("Afficher inactifs", value=False)

st.divider()

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
            'card_image_path': t.card_image_path,  # AJOUTER
            'is_active': t.is_active
        })

# Filtrer par recherche
if search:
    templates_data = [t for t in templates_data if search.lower() in t['name'].lower()]

# Affichage
if not templates_data:
    st.info("Aucun template trouvé. Créez-en un dans l'onglet 'Nouveau Template'.")
else:
    st.markdown(f"**{len(templates_data)} template(s) trouvé(s)**")
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
                        
                        # --- affichage image de carte avec lien cliquable ---
                        if st.button("", key=f"card_click_{template['id']}", use_container_width=True):
                            st.session_state.selected_template_detail = template['id']
                            st.switch_page("pages/4_📊_Detail_Livrable.py")

                        # --- affichage image de carte, normalisée 16:9 ---
                        default_image = project_root / "assets" / "background" / "card" / "default.png"

                        image_to_show = None
                        if template['card_image_path'] and Path(template['card_image_path']).exists():
                            image_to_show = template['card_image_path']
                        elif default_image.exists():
                            image_to_show = str(default_image)

                        def afficher_image_carte(path: str, ratio: float = 16/9, radius_px: int = 8):
                            """
                            Cadre à ratio fixe + object-fit:cover.
                            Empêche l'effet 'miniature' de st.image() au refresh.
                            """
                            try:
                                p = Path(path)
                                if not p.exists():
                                    p = default_image
                                b64 = base64.b64encode(p.read_bytes()).decode("ascii")
                                padding_pct = 100 / ratio  # 56.25% pour 16:9
                                st.markdown(f"""
                                <div style="position:relative;width:100%;padding-top:{padding_pct}%;
                                            overflow:hidden;border-radius:{radius_px}px;background:#10182014;">
                                <img src="data:image/png;base64,{b64}"
                                    style="position:absolute;inset:0;width:100%;height:100%;
                                            object-fit:cover;display:block;border-radius:{radius_px}px;">
                                </div>
                                """, unsafe_allow_html=True)
                            except Exception as e:
                                st.caption(f"🖼️ Image illisible ({e})")
                                # Fallback sûr sur l'image par défaut
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

                        # --- à l'endroit où tu affiches l'image de la carte ---
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
                        col_btn1, col_btn2, col_btn3 = st.columns(3)
                        
                        with col_btn1:
                            if st.button("▶️", key=f"gen_{template['id']}", help="Générer", use_container_width=True):
                                st.session_state.selected_template_for_generation = template['id']
                                st.switch_page("pages/3_▶️_Generer_Rapport.py")
                        
                        with col_btn2:
                            if st.button("✏️", key=f"edit_{template['id']}", help="Éditer", use_container_width=True):
                                st.session_state.selected_template = template['id']
                                st.switch_page("pages/2_➕_Nouveau_Template.py")
                        
                        with col_btn3:
                            if st.button("🗑️", key=f"del_{template['id']}", help="Supprimer", use_container_width=True):
                                if st.session_state.get(f"confirm_del_{template['id']}"):
                                    with DatabaseService.get_session() as db:
                                        service = TemplateService(db)
                                        service.delete_template(template['id'], hard_delete=False)
                                    st.success(f"Template '{template['name']}' désactivé")
                                    st.rerun()
                                else:
                                    st.session_state[f"confirm_del_{template['id']}"] = True
                                    st.warning("Cliquez à nouveau pour confirmer")