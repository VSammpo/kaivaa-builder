# frontend/pages/_3b_➕_Form_Gabarit.py
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.models.gabarits import TableGabarit, GabaritColumn
from backend.services.gabarit_registry import (
    get_gabarit, upsert_gabarit, list_gabarits,
    get_relations, add_relation, delete_relation,
    set_role, get_role
)

st.set_page_config(page_title="Formulaire Gabarit", page_icon="➕", layout="wide")

# ---- Mode édition / création -----------------------------------------------------
edit_mode = False
gabarit = None
if "selected_gabarit" in st.session_state and st.session_state.selected_gabarit:
    edit_mode = True
    gab_name, gab_version = st.session_state.selected_gabarit
    gabarit = get_gabarit(gab_name, gab_version)

# ---- Titre & retour --------------------------------------------------------------
if edit_mode:
    st.title(f"✏️ Modifier le gabarit '{gabarit.name}'")
else:
    st.title("➕ Créer un gabarit")

if st.button("🔙 Retour", key="back_btn"):
    if edit_mode:
        st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
    else:
        st.switch_page("pages/3_🧱_Gabarits.py")

st.divider()

# ---- Valeurs par défaut ----------------------------------------------------------
if edit_mode:
    name_init = gabarit.name
    version_init = gabarit.version
    desc_init = gabarit.description or ""
    cols_init = pd.DataFrame([c.model_dump() for c in gabarit.columns])
    existing_role = get_role(gabarit.name, gabarit.version) or "mixed"
else:
    name_init = ""
    version_init = "v1"
    desc_init = ""
    cols_init = pd.DataFrame([{"name": "", "type": "text", "is_key": False}])
    existing_role = "mixed"

# ==============================================================================
# FORMULAIRE PRINCIPAL : infos + rôle + colonnes
# ==============================================================================
with st.form("gabarit_form"):
    st.subheader("Informations générales")
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input(
            "Nom du gabarit*",
            value=name_init,
            disabled=edit_mode,
            help="Ex: SELL_OUT, PRICING_SOURCE",
        )
    with col2:
        version = st.text_input("Version", value=version_init, help="Ex: v1, v2")

    description = st.text_area("Description", value=desc_init, height=80)

    st.divider()
    st.subheader("Rôle du gabarit")
    table_role = st.selectbox(
        "Rôle", ["fact", "dimension", "mixed"],
        index=["fact", "dimension", "mixed"].index(existing_role),
        help="Un gabarit 'fact' est une table d’événements ; 'dimension' un référentiel ; 'mixed' peut servir des deux côtés."
    )

    st.divider()
    st.subheader("Colonnes")
    st.caption("Types disponibles : text | number | integer | date | boolean")

    edited = st.data_editor(
        cols_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "name": st.column_config.TextColumn("Nom", help="Nom de la colonne"),
            "type": st.column_config.SelectboxColumn("Type",
                                                     options=["text", "number", "integer", "date", "boolean"]),
            "is_key": st.column_config.CheckboxColumn("Clé ?", help="Fait partie de la clé composite"),
        },
        key="columns_editor",
    )

    st.divider()
    with st.expander("⚙️ Méthodes (à venir)", expanded=False):
        st.info("Configuration des méthodes disponible prochainement")

    submitted = st.form_submit_button(
        "💾 Enregistrer" if edit_mode else "🚀 Créer",
        type="primary",
        use_container_width=True,
    )

    if submitted:
        if not name.strip():
            st.error("Le nom est obligatoire")
        else:
            # Construire la liste des colonnes à partir du data_editor
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
                        description=(description or "").strip(),  # <-- corrige l’ancien bug
                        columns=cols,
                    )
                    upsert_gabarit(g)
                    set_role(g.name, g.version, table_role)
                    st.success(f"Gabarit enregistré : {g.name} v{g.version}")

                    # Basculer en mode édition sur ce gabarit
                    st.session_state.selected_gabarit = (g.name, g.version)
                    st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")

                except Exception as e:
                    st.error(f"Erreur : {e}")

# ==============================================================================
# RELATIONS (catalogue) – en dehors du form principal (pas de st.button dans un form)
# ==============================================================================
st.divider()
st.subheader("Relations autorisées (catalogue)")

if not edit_mode:
    st.info("Enregistrez d'abord le gabarit pour déclarer des relations.")
else:
    # Recharger le gabarit actuel (au cas où update)
    gab = get_gabarit(gabarit.name, gabarit.version)
    local_cols = [c.name for c in (gab.columns or [])]

    # ---- Formulaire d'ajout d'une relation (séparé) ----------------------------
    with st.form("relation_form"):
        st.caption("Déclarez une clé ↔ clé depuis ce gabarit vers un gabarit cible (pas de type de jointure ici).")

        all_gabs = list_gabarits()
        choices = [
            f"{g.name}|{g.version}"
            for g in all_gabs
            if not (g.name == gab.name and g.version == gab.version)  # éviter self-join direct
        ]

        colA, colB = st.columns(2)
        with colA:
            target = st.selectbox("Gabarit cible (dimension)", choices, index=0 if choices else None, key="rel_target")
        with colB:
            left_key = st.selectbox("Clé locale (FROM)", local_cols, index=0 if local_cols else None, key="rel_left")

        # Colonnes du gabarit cible
        target_cols = []
        if target:
            tgt_name, tgt_ver = target.split("|")
            tgt_gab = get_gabarit(tgt_name, tgt_ver)
            if tgt_gab:
                target_cols = [c.name for c in tgt_gab.columns]

        if target_cols:
            right_key = st.selectbox("Clé côté cible (TO)", target_cols, key="rel_right_select")
        else:
            right_key = st.text_input("Clé côté cible (TO)", value="", placeholder="ex: siren, nafrev2", key="rel_right")

        add_ok = st.form_submit_button("➕ Ajouter la relation", use_container_width=True)
        if add_ok:
            if not (target and left_key and right_key):
                st.error("Veuillez sélectionner un gabarit cible et les deux clés.")
            else:
                tgt_name, tgt_ver = target.split("|")
                add_relation(
                    from_gabarit=gab.name, from_version=gab.version,
                    to_gabarit=tgt_name, to_version=tgt_ver,
                    left_key=left_key, right_key=right_key,
                )
                st.success("Relation ajoutée au gabarit.")
                st.rerun()

    # ---- Liste & suppression (en dehors de tout form) ---------------------------
    existing_relations = get_relations(gab.name, gab.version)
    if existing_relations:
        st.caption("Relations déclarées :")
        for r in existing_relations:
            with st.container(border=True):
                st.write(
                    f"**{r['from_gabarit']}[{r.get('from_version','v1')}]** "
                    f"`{r['left_key']}` = `{r['right_key']}` "
                    f"→ **{r['to_gabarit']}[{r.get('to_version','v1')}]**"
                )
                if st.button("Supprimer", key=f"del_rel_{r['relation_id']}"):
                    delete_relation(
                        r["from_gabarit"], r.get("from_version", "v1"),
                        r["to_gabarit"], r.get("to_version", "v1"),
                        r["left_key"], r["right_key"],
                    )
                    st.rerun()
    else:
        st.info("Aucune relation déclarée pour ce gabarit.")
