# frontend/pages/3b_➕_Form_Gabarit.py
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.models.gabarits import TableGabarit, GabaritColumn
from backend.services.gabarit_registry import get_gabarit, upsert_gabarit

st.set_page_config(page_title="Formulaire Gabarit", page_icon="➕", layout="wide")

# Déterminer mode
edit_mode = False
gabarit = None

if 'selected_gabarit' in st.session_state and st.session_state.selected_gabarit:
    edit_mode = True
    gab_name, gab_version = st.session_state.selected_gabarit
    gabarit = get_gabarit(gab_name, gab_version)

# Titre
if edit_mode:
    st.title(f"✏️ Modifier le gabarit '{gabarit.name}'")
else:
    st.title("➕ Créer un gabarit")

if st.button("🔙 Retour", key="back_btn"):
    if edit_mode:
        st.switch_page("pages/3a_🧱_Detail_Gabarit.py")
    else:
        st.switch_page("pages/3_🧱_Gabarits.py")

st.divider()

# Valeurs par défaut
if edit_mode:
    name_init = gabarit.name
    version_init = gabarit.version
    desc_init = gabarit.description or ""
    cols_init = pd.DataFrame([c.model_dump() for c in gabarit.columns])
else:
    name_init = ""
    version_init = "v1"
    desc_init = ""
    cols_init = pd.DataFrame([{"name": "", "type": "text", "is_key": False}])

# Formulaire
with st.form("gabarit_form"):
    st.subheader("Informations générales")
    
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input(
            "Nom du gabarit*",
            value=name_init,
            disabled=edit_mode,
            help="Ex: SELL_OUT, PRICING_SOURCE"
        )
    with col2:
        version = st.text_input("Version", value=version_init, help="Ex: v1, v2")
    
    description = st.text_area("Description", value=desc_init, height=80)
    
    st.divider()
    st.subheader("Colonnes")
    st.caption("Types disponibles : text | number | integer | date | boolean")
    
    edited = st.data_editor(
        cols_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "name": st.column_config.TextColumn("Nom", help="Nom de la colonne"),
            "type": st.column_config.SelectboxColumn(
                "Type",
                options=["text", "number", "integer", "date", "boolean"]
            ),
            "is_key": st.column_config.CheckboxColumn(
                "Clé ?",
                help="Fait partie de la clé composite"
            )
        }
    )
    
    st.divider()
    
    # Méthodes (placeholder)
    with st.expander("⚙️ Méthodes (à venir)", expanded=False):
        st.info("Configuration des méthodes disponible prochainement")
    
    st.divider()
    
    # Bouton submit
    submitted = st.form_submit_button(
        "💾 Enregistrer" if edit_mode else "🚀 Créer",
        type="primary",
        use_container_width=True
    )
    
    if submitted:
        if not name.strip():
            st.error("Le nom est obligatoire")
        else:
            # Construire colonnes
            edited = edited.fillna("")
            cols = []
            seen = set()
            
            for r in edited.to_dict(orient="records"):
                n = (r.get("name") or "").strip()
                if not n or n in seen:
                    continue
                t = (r.get("type") or "text").strip().lower()
                is_key = bool(r.get("is_key", False))
                cols.append(GabaritColumn(name=n, type=t, is_key=is_key))
                seen.add(n)
            
            if not cols:
                st.error("Au moins une colonne est requise")
            else:
                try:
                    g = TableGabarit(
                        name=name.strip(),
                        version=version.strip(),
                        description=desc_init.strip() if desc_init else None,
                        columns=cols
                    )
                    upsert_gabarit(g)
                    st.success(f"Gabarit enregistré : {g.name} v{g.version}")
                    
                    # Redirection
                    st.session_state.selected_gabarit = (g.name, g.version)
                    st.switch_page("pages/3a_🧱_Detail_Gabarit.py")
                
                except Exception as e:
                    st.error(f"Erreur : {e}")