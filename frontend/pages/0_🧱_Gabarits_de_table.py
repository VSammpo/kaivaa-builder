# frontend/pages/0_🧱_Gabarits_de_table.py
import streamlit as st
import pandas as pd

from backend.models.gabarits import TableGabarit, GabaritColumn
from backend.services.gabarit_registry import list_gabarits, upsert_gabarit, delete_gabarit, get_gabarit

st.set_page_config(page_title="Gabarits de table", layout="wide")
st.title("🧱 Gabarits de table (globaux) — MVP")

# Liste
gabarits = list_gabarits()
labels = [f"{g.name}  (v{g.version})" for g in gabarits]
choice = st.selectbox("Gabarit existant :", ["—"] + labels)
current = None
if choice != "—":
    idx = labels.index(choice)
    current = gabarits[idx]

col_left, col_right = st.columns([1,2], gap="large")

with col_left:
    st.subheader("📚 Registre")
    if gabarits:
        df = pd.DataFrame([{
            "name": g.name,
            "version": g.version,
            "n_cols": len(g.columns),
            "n_keys": sum(1 for c in g.columns if c.is_key),
            "description": (g.description or "")
        } for g in gabarits])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.caption("Aucun gabarit pour l’instant.")

with col_right:
    st.subheader("✏️ Créer / Éditer")

    mode = st.radio("Mode", ["➕ Nouveau", "Éditer"], horizontal=True)
    if mode == "Éditer" and not current:
        st.info("Sélectionne un gabarit à gauche pour l’éditer.")
        st.stop()

    if mode == "Éditer":
        name_init = current.name
        version_init = current.version
        desc_init = current.description or ""
        cols_init = pd.DataFrame([c.model_dump() for c in current.columns])
    else:
        name_init = ""
        version_init = "v1"
        desc_init = ""
        cols_init = pd.DataFrame([{"name": "", "type": "text", "is_key": False}])

    with st.form("gab_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("Nom du gabarit", value=name_init, help="Ex: SELL_OUT, PRICING_SOURCE… (global et réutilisable)")
        with c2:
            version = st.text_input("Version", value=version_init, help="Ex: v1")

        description = st.text_area("Description (optionnel)", value=desc_init)

        st.caption("Colonnes (types MVP : text | number | integer | date | boolean ; coche = clé composite)")
        edited = st.data_editor(
            cols_init,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "name": st.column_config.TextColumn("name"),
                "type": st.column_config.SelectboxColumn("type", options=["text","number","integer","date","boolean"]),
                "is_key": st.column_config.CheckboxColumn("clé ?", help="Inclure dans la clé composite")
            }
        )

        submitted = st.form_submit_button("💾 Enregistrer")
        if submitted:
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

            try:
                g = TableGabarit(name=name.strip(), version=version.strip(), description=(description or "").strip(), columns=cols)
                upsert_gabarit(g)
                st.success(f"Gabarit enregistré : {g.name} v{g.version}")
            except Exception as e:
                st.error(f"Erreur: {e}")

    if mode == "Éditer" and current:
        if st.button("🗑️ Supprimer ce gabarit", type="secondary"):
            ok = delete_gabarit(current.name, current.version)
            if ok:
                st.success(f"Supprimé : {current.name} v{current.version}")
            else:
                st.warning("Non supprimé.")
