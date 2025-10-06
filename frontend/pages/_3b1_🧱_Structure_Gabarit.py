# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

# ==== Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

# ==== Services & modèles
from backend.services.gabarit_registry import (
    get_gabarit, upsert_gabarit, set_role, get_role
)
from backend.models.gabarits import TableGabarit, GabaritColumn

st.set_page_config(page_title="Structure du gabarit", page_icon="🧱", layout="wide")

# ========= Navbar homogène (retour + sous-onglets)
def render_gabarit_subnav(active: str):
    # active ∈ {"structure","enrich","methods","default"}
    cols = st.columns([1, 1, 1, 1, 1])

    with cols[0]:
        if st.button("← Fiche gabarit", key=f"subnav_back_{active}", use_container_width=True):
            st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")

    with cols[1]:
        if st.button("📊 Structure", key=f"subnav_struct_{active}",
                     type=("primary" if active == "structure" else "secondary"),
                     use_container_width=True):
            if active != "structure":
                st.switch_page("pages/_3b1_🧱_Structure_Gabarit.py")

    with cols[2]:
        if st.button("🔗 Enrichissements", key=f"subnav_enrich_{active}",
                     type=("primary" if active == "enrich" else "secondary"),
                     use_container_width=True):
            if active != "enrich":
                st.switch_page("pages/_3b2_🔗_Enrichissements_Gabarit.py")

    with cols[3]:
        if st.button("⚙️ Méthodes", key=f"subnav_methods_{active}",
                     type=("primary" if active == "methods" else "secondary"),
                     use_container_width=True):
            if active != "methods":
                st.switch_page("pages/_3c_⚙️_Methodes_Gabarit.py")

    with cols[4]:
        if st.button("📁 Données par défaut", key=f"subnav_default_{active}",
                     type=("primary" if active == "default" else "secondary"),
                     use_container_width=True):
            if active != "default":
                st.switch_page("pages/_3b3_📁_Donnee_Par_Defaut.py")
    st.divider()

# ========= Déterminer mode (création vs édition)
create_mode = False
if "selected_gabarit" not in st.session_state or not st.session_state.selected_gabarit:
    create_mode = True
    gab_name, gab_version = "", "v1"
    gabarit = None
else:
    gab_name, gab_version = st.session_state.selected_gabarit
    gabarit = get_gabarit(gab_name, gab_version)

render_gabarit_subnav("structure")

st.title("➕ Créer un gabarit" if create_mode else f"✏️ Structure : {gab_name} [{gab_version}]")
st.divider()
st.subheader("Structure des colonnes")

existing_role = get_role(gab_name, gab_version) if not create_mode else "mixed"
existing_role = existing_role or "mixed"

with st.form("structure_form"):
    col1, col2 = st.columns([2,1])
    with col1:
        name_val = st.text_input("Nom du gabarit *",
                                 value=(gab_name if not create_mode else ""),
                                 disabled=not create_mode,
                                 placeholder="Ex: SELL_OUT")
    with col2:
        ver_val = st.text_input("Version *",
                                value=(gab_version if not create_mode else "v1"),
                                disabled=not create_mode)

    role = st.selectbox("Rôle", ["fact","dimension","mixed"],
                        index=["fact","dimension","mixed"].index(existing_role),
                        help="Fact = événements • Dimension = référentiel • Mixed = hybride")

    st.caption("Types disponibles : text | number | integer | date | boolean")

    # Table des colonnes
    df_init = (
        pd.DataFrame([c.model_dump() for c in gabarit.columns])
        if (gabarit and gabarit.columns) else
        pd.DataFrame([{"name":"", "type":"text", "is_key": False}])
    )
    edited = st.data_editor(
        df_init, num_rows="dynamic", use_container_width=True,
        column_config={
            "name": st.column_config.TextColumn("Nom *", width="medium"),
            "type": st.column_config.SelectboxColumn(
                "Type", options=["text","number","integer","date","boolean"], width="small"
            ),
            "is_key": st.column_config.CheckboxColumn("Clé ?", width="small"),
        },
        key="columns_editor_edit",
    )

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        save = st.form_submit_button(
            "🚀 Créer le gabarit" if create_mode else "💾 Enregistrer la structure",
            type="primary", use_container_width=True
        )
    with c2:
        if st.form_submit_button("Annuler", use_container_width=True):
            if create_mode:
                st.switch_page("pages/3_🧱_Gabarits.py")
            else:
                st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")

    if save:
        rows = edited.fillna("").to_dict(orient="records")
        cols, seen = [], set()
        for r in rows:
            n = (r.get("name") or "").strip()
            if not n or n in seen:
                continue
            t = (r.get("type") or "text").strip().lower()
            is_key = bool(r.get("is_key", False))
            cols.append(GabaritColumn(name=n, type=t, is_key=is_key))
            seen.add(n)

        if create_mode and not name_val.strip():
            st.error("❌ Le nom du gabarit est requis")
        elif not cols:
            st.error("❌ Au moins une colonne est requise")
        else:
            try:
                g = TableGabarit(
                    name=(name_val if create_mode else gab_name).strip(),
                    version=(ver_val if create_mode else gab_version).strip(),
                    description=gabarit.description if gabarit else "",
                    columns=cols
                )
                upsert_gabarit(g)
                set_role(g.name, g.version, role)
                st.session_state.selected_gabarit = (g.name, g.version)
                st.success("✅ Gabarit créé" if create_mode else "✅ Structure enregistrée")
                st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
            except Exception as e:
                st.error(f"❌ Erreur : {e}")
