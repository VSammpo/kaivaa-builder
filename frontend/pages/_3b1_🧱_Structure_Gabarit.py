import streamlit as st
import pandas as pd
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.gabarit_registry import (
    get_gabarit, upsert_gabarit,
    set_role, get_role
)
from backend.models.gabarits import TableGabarit, GabaritColumn

st.set_page_config(page_title="Structure du gabarit", page_icon="🧱", layout="wide")

# Sélection gabarit
if "selected_gabarit" not in st.session_state or not st.session_state.selected_gabarit:
    st.error("Aucun gabarit sélectionné")
    if st.button("← Retour aux gabarits", use_container_width=True):
        st.switch_page("pages/3_🧱_Gabarits.py")
    st.stop()

gab_name, gab_version = st.session_state.selected_gabarit
gabarit = get_gabarit(gab_name, gab_version)

col_back, col_title = st.columns([1, 5])
with col_back:
    if st.button("← Détail gabarit", use_container_width=True):
        st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
with col_title:
    st.title(f"✏️ Structure : {gabarit.name} [{gabarit.version}]")

st.divider()
st.subheader("Structure des colonnes")

existing_role = get_role(gabarit.name, gabarit.version) or "mixed"

with st.form("structure_form"):
    col1, col2 = st.columns([2,1])
    with col1:
        st.text_input("Nom du gabarit", value=gabarit.name, disabled=True)
    with col2:
        st.text_input("Version", value=gabarit.version, disabled=True)

    role = st.selectbox("Rôle", ["fact","dimension","mixed"],
                        index=["fact","dimension","mixed"].index(existing_role),
                        help="Fact = événements • Dimension = référentiel • Mixed = hybride")

    st.caption("Types disponibles : text | number | integer | date | boolean")

    df_init = pd.DataFrame([c.model_dump() for c in gabarit.columns])
    edited = st.data_editor(
        df_init, num_rows="dynamic", use_container_width=True,
        column_config={
            "name": st.column_config.TextColumn("Nom *", width="medium"),
            "type": st.column_config.SelectboxColumn("Type", options=["text","number","integer","date","boolean"], width="small"),
            "is_key": st.column_config.CheckboxColumn("Clé ?", width="small"),
        },
        key="columns_editor_edit",
    )

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        save = st.form_submit_button("💾 Enregistrer la structure", type="primary", use_container_width=True)
    with c2:
        if st.form_submit_button("Annuler", use_container_width=True):
            st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")

    if save:
        rows = edited.fillna("").to_dict(orient="records")
        cols = []
        seen = set()
        for r in rows:
            n = (r.get("name") or "").strip()
            if not n or n in seen: 
                continue
            t = (r.get("type") or "text").strip().lower()
            is_key = bool(r.get("is_key", False))
            cols.append(GabaritColumn(name=n, type=t, is_key=is_key))
            seen.add(n)

        if not cols:
            st.error("❌ Au moins une colonne est requise")
        else:
            try:
                g = TableGabarit(
                    name=gabarit.name, version=gabarit.version,
                    description=gabarit.description or "",
                    columns=cols
                )
                upsert_gabarit(g)
                set_role(g.name, g.version, role)
                st.success("✅ Structure enregistrée")
                st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
            except Exception as e:
                st.error(f"❌ Erreur : {e}")
