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

# Filtrer par recherche
if search:
    templates = [t for t in templates if search.lower() in t.name.lower()]

# Affichage
if not templates:
    st.info("Aucun template trouvé. Créez-en un dans l'onglet 'Nouveau Template'.")
else:
    st.markdown(f"**{len(templates)} template(s) trouvé(s)**")
    
    for template in templates:
        with st.expander(f"📄 {template.name} (v{template.version})"):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.markdown(f"**Description :** {template.description or '_Aucune description_'}")
                st.markdown(f"**Créé le :** {template.created_at.strftime('%d/%m/%Y')}")
                st.markdown(f"**Statut :** {'✅ Actif' if template.is_active else '❌ Inactif'}")
                
                # Statistiques
                stats = service.get_template_stats(template.id)
                st.markdown(f"""
                **Statistiques :**
                - Exécutions totales : {stats['total_executions']}
                - Taux de succès : {stats['success_rate']}%
                - Temps moyen : {stats['avg_execution_time_seconds']}s
                """)
            
            with col2:
                st.markdown("**Actions**")
                
                if st.button("🔍 Détails", key=f"view_{template.id}"):
                    st.session_state.selected_template = template.id
                    st.switch_page("pages/2_➕_Nouveau_Template.py")
                
                if st.button("▶️ Générer", key=f"gen_{template.id}"):
                    st.session_state.selected_template = template.id
                    st.switch_page("pages/3_▶️_Generer_Rapport.py")
                
                if st.button("❌ Supprimer", key=f"del_{template.id}"):
                    if st.session_state.get(f"confirm_del_{template.id}"):
                        service.delete_template(template.id, hard_delete=False)
                        st.success(f"Template '{template.name}' désactivé")
                        st.rerun()
                    else:
                        st.session_state[f"confirm_del_{template.id}"] = True
                        st.warning("Cliquez à nouveau pour confirmer")