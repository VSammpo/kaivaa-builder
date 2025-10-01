"""
Page de la bibliothèque de templates
"""

import streamlit as st
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService

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
            'created_at': t.created_at.strftime('%d/%m/%Y'),
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
    
    for template in templates_data:
        template_id = template['id']
        template_name = template['name']
        
        with st.expander(f"📄 {template_name} (v{template['version']})"):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.markdown(f"**Description :** {template['description'] or '_Aucune description_'}")
                st.markdown(f"**Créé le :** {template['created_at']}")
                st.markdown(f"**Statut :** {'✅ Actif' if template['is_active'] else '❌ Inactif'}")
                
                # Statistiques
                with DatabaseService.get_session() as db:
                    service = TemplateService(db)
                    stats = service.get_template_stats(template_id)
                
                st.markdown(f"""
                **Statistiques :**
                - Exécutions totales : {stats['total_executions']}
                - Taux de succès : {stats['success_rate']}%
                - Temps moyen : {stats['avg_execution_time_seconds']}s
                """)
            
            with col2:
                st.markdown("**Actions**")
                
                if st.button("🔍 Détails", key=f"view_{template_id}"):
                    st.session_state.selected_template = template_id
                    st.switch_page("pages/2_➕_Nouveau_Template.py")
                
                if st.button("▶️ Générer", key=f"gen_{template_id}", disabled=True):
                    st.info("Fonctionnalité en cours de développement")
                
                if st.button("❌ Supprimer", key=f"del_{template_id}"):
                    if st.session_state.get(f"confirm_del_{template_id}"):
                        with DatabaseService.get_session() as db:
                            service = TemplateService(db)
                            service.delete_template(template_id, hard_delete=False)
                        st.success(f"Template '{template_name}' désactivé")
                        st.rerun()
                    else:
                        st.session_state[f"confirm_del_{template_id}"] = True
                        st.warning("Cliquez à nouveau pour confirmer")