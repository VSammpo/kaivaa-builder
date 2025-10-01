"""
Page d'accueil de KAIVAA Builder
"""

import streamlit as st
from pathlib import Path
import sys

# Ajouter le backend au path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.database.models import Template

# Configuration de la page
st.set_page_config(
    page_title="KAIVAA Builder",
    page_icon="🎯",
    layout="wide"
)

# Initialiser la base de données
DatabaseService.create_tables()

# En-tête
st.title("🎯 KAIVAA Builder")
st.markdown("**Template Builder pour génération automatisée de présentations**")

st.divider()

# Stats globales
with DatabaseService.get_session() as db:
    total_templates = db.query(Template).count()
    active_templates = db.query(Template).filter_by(is_active=True).count()

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Templates créés", total_templates)

with col2:
    st.metric("Templates actifs", active_templates)

with col3:
    st.metric("En développement", "Phase 3")

st.divider()

# Guide rapide
st.header("Guide rapide")

st.markdown("""
### 🚀 Commencer

1. **📚 Bibliothèque** : Consultez vos templates existants
2. **➕ Nouveau Template** : Créez un nouveau template
3. **▶️ Générer Rapport** : Lancez une génération

### 🎨 Fonctionnalités

- ✅ Génération automatique de templates Excel/PowerPoint
- ✅ Configuration via interface web
- ✅ Support multi-sources de données (PostgreSQL, Excel, CSV)
- ✅ Gestion des boucles et répétitions de slides
- ✅ Injection d'images dynamiques
- 🚧 Custom tables SQL + Python (en cours)
- 🚧 Générateur de rapports par batch

### 📖 Documentation

- [GitHub](https://github.com/votre-username/kaivaa-builder)
- [Documentation API](https://docs.kaivaa-builder.com)
""")

# Footer
st.divider()
st.caption("KAIVAA Builder v0.1.0 - Développé pour SAMMPO")