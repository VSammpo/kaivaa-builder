# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

# ==== Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

# ==== Services
from backend.services.gabarit_registry import (
    get_gabarit, list_gabarits,
    get_relations, add_relation, delete_relation
)

st.set_page_config(page_title="Enrichissements", page_icon="🔗", layout="wide")

# ========= Navbar homogène
def render_gabarit_subnav(active: str):
    cols = st.columns([1, 1, 1, 1, 1])

    with cols[0]:
        if st.button("← Fiche gabarit", key=f"subnav_back_{active}", use_container_width=True):
            st.switch_page("pages/_1a_🧱_Detail_Gabarit.py")

    with cols[1]:
        if st.button("📊 Structure", key=f"subnav_struct_{active}",
                     type=("primary" if active == "structure" else "secondary"),
                     use_container_width=True):
            if active != "structure":
                st.switch_page("pages/_1b_🧱_Structure_Gabarit.py")

    with cols[2]:
        if st.button("🔗 Enrichissements", key=f"subnav_enrich_{active}",
                     type=("primary" if active == "enrich" else "secondary"),
                     use_container_width=True):
            if active != "enrich":
                st.switch_page("pages/_1c_🔗_Enrichissements_Gabarit.py")

    with cols[3]:
        if st.button("⚙️ Méthodes", key=f"subnav_methods_{active}",
                     type=("primary" if active == "methods" else "secondary"),
                     use_container_width=True):
            if active != "methods":
                st.switch_page("pages/_1e_⚙️_Methodes_Gabarit.py")

    with cols[4]:
        if st.button("📁 Données par défaut", key=f"subnav_default_{active}",
                     type=("primary" if active == "default" else "secondary"),
                     use_container_width=True):
            if active != "default":
                st.switch_page("pages/_1d_📁_Donnee_Par_Defaut.py")
    st.divider()

# ==== Sélection gabarit obligatoire
if "selected_gabarit" not in st.session_state or not st.session_state.selected_gabarit:
    st.error("Aucun gabarit sélectionné.")
    if st.button("← Retour aux gabarits", use_container_width=True):
        st.switch_page("pages/1_🧱_Gabarits.py")
    st.stop()

gab_name, gab_version = st.session_state.selected_gabarit
gabarit = get_gabarit(gab_name, gab_version)

render_gabarit_subnav("enrich")
st.title(f"🔗 Enrichissements : {gabarit.name} [{gabarit.version}]")
st.divider()

with st.expander("💡 Principe", expanded=False):
    st.markdown("""
Vous enrichissez **cette table** avec une **table de référence** (dimension).
- **Clé locale** : la colonne dans cette table
- **Clé d'enrichissement** : la colonne dans la table de référence
""")

# ==== Liste des enrichissements existants
existing_relations = get_relations(gabarit.name, gabarit.version)
if existing_relations:
    st.caption(f"📋 {len(existing_relations)} enrichissement(s)")
    for r in existing_relations:
        with st.container(border=True):
            col_info, col_del = st.columns([4,1])
            with col_info:
                st.markdown(f"**Depuis** `{r['to_gabarit']}` [{r.get('to_version','v1')}]")
                st.caption(f"Jointure : `{r['left_key']}` = `{r['right_key']}`")
            with col_del:
                if st.button("🗑️", key=f"del_enrich_{r['relation_id']}", use_container_width=True):
                    try:
                        delete_relation(
                            r["from_gabarit"], r.get("from_version", "v1"),
                            r["to_gabarit"], r.get("to_version", "v1"),
                            r["left_key"], r["right_key"],
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Suppression impossible : {e}")
else:
    st.info("Aucun enrichissement configuré.")

st.divider()
st.subheader("Ajouter un enrichissement")

# ==== Ajout d'un enrichissement
all_gabs = list_gabarits()
targets = [
    f"{g.name}|{g.version}"
    for g in all_gabs
    if not (g.name == gabarit.name and g.version == gabarit.version)
]

target = st.selectbox(
    "Table de référence",
    options=targets,
    index=0 if targets else None,
    format_func=lambda s: f"{s.split('|')[0]} [{s.split('|')[1]}]" if '|' in s else s,
)

local_cols = [c.name for c in (gabarit.columns or [])]
target_cols = []
if target:
    tgt_name, tgt_ver = target.split("|",1)
    tgt_gab = get_gabarit(tgt_name, tgt_ver)
    if tgt_gab:
        target_cols = [c.name for c in (tgt_gab.columns or [])]

col1, col2 = st.columns(2)
with col1:
    left_key = st.selectbox("Clé locale", options=local_cols, index=0 if local_cols else None)
with col2:
    if target_cols:
        right_key = st.selectbox("Clé de référence", options=target_cols, index=0)
    else:
        right_key = st.text_input("Clé de référence", placeholder="ex: siren")

if st.button("➕ Ajouter l'enrichissement", use_container_width=True, type="primary", disabled=not target):
    if not left_key or not right_key:
        st.error("Champs incomplets.")
    else:
        try:
            tgt_name, tgt_ver = target.split("|",1)
            add_relation(
                from_gabarit=gabarit.name, from_version=gabarit.version,
                to_gabarit=tgt_name, to_version=tgt_ver,
                left_key=left_key, right_key=right_key,
            )
            st.success("✅ Enrichissement ajouté")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
