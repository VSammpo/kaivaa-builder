# frontend/pages/_3b_➕_Form_Gabarit.py
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

# === Path bootstrap ===
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

# === Backend imports ===
from backend.services.gabarit_registry import (
    get_gabarit, upsert_gabarit, list_gabarits,
    get_relations, add_relation, delete_relation,
    set_role, get_role,
    get_default_source, set_default_source, clear_default_source,
    set_default_preview  # <- nouveau import : on mémorise l'aperçu persistant
)
from backend.models.gabarits import TableGabarit, GabaritColumn

# Alignement & validation simple DF <-> colonnes attendues
try:
    from backend.services.dataset_service import align_df_to_expected_columns
except Exception:
    # Fallback minimal (au cas où le module bouge)
    def align_df_to_expected_columns(df: pd.DataFrame, expected_columns):
        expected = [c for c in (expected_columns or []) if isinstance(c, str) and c.strip()]
        cur_cols = list(df.columns)
        missing = [c for c in expected if c not in cur_cols]
        for c in missing:
            df[c] = pd.NA
        ordered = expected + [c for c in df.columns if c not in expected]
        return df[ordered], {"missing": missing, "extra": [c for c in cur_cols if c not in expected]}

st.set_page_config(page_title="Formulaire Gabarit", page_icon="➕", layout="wide")

# =====================================================================
# Helpers internes pour la donnée par défaut (preview / validation)
# =====================================================================
def _try_load_default_source(fmt: str, path: str, sep: str | None, enc: str | None, head: int = 20):
    """Charge un extrait du fichier selon fmt (csv/parquet). Retourne (df|None, err|None)."""
    try:
        p = Path(path)
        if not p.exists() or not p.is_file():
            return None, f"Fichier introuvable : {path}"
        if fmt == "csv":
            df = pd.read_csv(path, sep=sep or ";", encoding=enc or "utf-8-sig")
        elif fmt == "parquet":
            df = pd.read_parquet(path)
        else:
            return None, f"Format non supporté : {fmt}"
        return df.head(head), None
    except Exception as e:
        return None, str(e)

def _apply_python(df: pd.DataFrame, code: str | None):
    """Applique un snippet Python (df->df). Retourne (df|None, err|None)."""
    if not code or not isinstance(code, str) or not code.strip():
        return df, None
    try:
        loc = {"df": df, "pd": pd}
        exec(code, {}, loc)
        new_df = loc.get("df")
        if isinstance(new_df, pd.DataFrame):
            return new_df, None
        return df, None
    except Exception as e:
        return None, f"Erreur script Python : {e}"

def _profile_df(df: pd.DataFrame) -> dict:
    """Petit profiling similaire à la page Pipeline."""
    def _dtype(s: pd.Series) -> str:
        dt = str(s.dtype)
        if "datetime" in dt:
            return "date"
        if "int" in dt or "float" in dt or "decimal" in dt:
            return "num"
        return "text"
    return {
        "rows": int(df.shape[0]),
        "cols": int(df.shape[1]),
        "completeness": {c: float(1 - df[c].isna().mean()) for c in df.columns},
        "dtypes": {c: _dtype(df[c]) for c in df.columns},
    }

# =====================================================================
# Détermination mode
# =====================================================================
edit_mode = False
gabarit = None
if "selected_gabarit" in st.session_state and st.session_state.selected_gabarit:
    edit_mode = True
    gab_name, gab_version = st.session_state.selected_gabarit
    gabarit = get_gabarit(gab_name, gab_version)

# ---- Titre & retour
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

# ---- Valeurs par défaut (création vs édition)
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

# =====================================================================
# FORMULAIRE PRINCIPAL : infos + rôle + colonnes
# =====================================================================
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
            "type": st.column_config.SelectboxColumn("Type", options=["text", "number", "integer", "date", "boolean"]),
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
                        description=(description or "").strip(),
                        columns=cols,
                    )
                    upsert_gabarit(g)
                    set_role(g.name, g.version, table_role)
                    st.success(f"Gabarit enregistré : {g.name} v{g.version}")

                    # Basculer en mode édition
                    st.session_state.selected_gabarit = (g.name, g.version)
                    st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
                except Exception as e:
                    st.error(f"Erreur : {e}")

# =====================================================================
# ENRICHISSEMENTS (catalogue)
# =====================================================================
st.divider()
st.subheader("Enrichissements")

if not edit_mode:
    st.info("Enregistrez d'abord le gabarit pour définir des enrichissements.")
else:
    gab = get_gabarit(gabarit.name, gabarit.version)
    local_cols = [c.name for c in (gab.columns or [])]

    st.markdown(
        """
        *Principe :* vous **enrichissez** la table courante (**source**) avec une **table d’enrichissement** (dimension).
        On part **toujours** de la table de départ (celle que vous éditez), et on rattache une info complémentaire
        via une **clé locale** = **clé d’enrichissement**.

        **Exemple :** enrichir une table de faits de CA par `siren` avec le **nom d’entreprise** tiré d’une table dimension `SIREN`.
        """
    )

    # --- Sélection de la table d'enrichissement EN DEHORS du form (pour forcer le rerun immédiat)
    all_gabs = list_gabarits()
    options = [
        f"{g.name}|{g.version}"
        for g in all_gabs
        if not (g.name == gab.name and g.version == gab.version)
    ]

    if "_enrich_prev_target" not in st.session_state:
        st.session_state._enrich_prev_target = None

    target = st.selectbox(
        "Table d’enrichissement (dimension)",
        options=options,
        index=0 if options else None,
        key="enrich_target_select",
        format_func=lambda s: f"{s.split('|')[0]} [{s.split('|')[1]}]" if isinstance(s, str) and '|' in s else s
    )

    # Déterminer les colonnes de la table d’enrichissement
    target_cols = []
    tgt_name = tgt_ver = None
    if target:
        tgt_name, tgt_ver = target.split("|", 1)
        tgt_gab = get_gabarit(tgt_name, tgt_ver)
        if tgt_gab:
            target_cols = [c.name for c in (tgt_gab.columns or [])]

    # Reset du choix de la clé droite si la cible change
    if st.session_state._enrich_prev_target != target:
        if "enrich_right_key" in st.session_state:
            del st.session_state["enrich_right_key"]
        st.session_state._enrich_prev_target = target

    # --- Formulaire d'ajout d'enrichissement
    with st.form("enrichment_form"):
        st.caption("Mappez la **clé locale** (dans la table source) avec la **clé d’enrichissement** (dans la table sélectionnée).")
        colA, colB = st.columns(2)
        with colA:
            left_key = st.selectbox(
                "Clé dans la table **source** (cette table)",
                options=local_cols,
                index=0 if local_cols else None,
                key="enrich_left_key",
                help="Ex : siren de votre table de faits"
            )
        with colB:
            if target_cols:
                right_key = st.selectbox(
                    "Clé dans la table **d’enrichissement**",
                    options=target_cols,
                    index=0,
                    key="enrich_right_key",
                    help=f"Enrichissement : {tgt_name}[{tgt_ver}] · Ex : siren"
                )
            else:
                right_key = st.text_input(
                    "Clé dans la table **d’enrichissement**",
                    value="",
                    placeholder="ex: siren, nafrev2",
                    key="enrich_right_fallback"
                )

        add_ok = st.form_submit_button("➕ Ajouter un enrichissement", use_container_width=True)

        if add_ok:
            if not target:
                st.error("Veuillez choisir une table d’enrichissement.")
            elif not left_key or not (right_key or (isinstance(right_key, str) and right_key.strip())):
                st.error("Veuillez renseigner les deux clés (source et enrichissement).")
            else:
                try:
                    add_relation(
                        from_gabarit=gab.name, from_version=gab.version,   # table source (gauche)
                        to_gabarit=tgt_name, to_version=tgt_ver,          # table d’enrichissement (droite)
                        left_key=left_key,
                        right_key=(right_key if right_key else st.session_state.get("enrich_right_fallback", "")).strip(),
                    )
                    st.success("Enrichissement ajouté.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur lors de l'ajout : {e}")

    # --- Liste des enrichissements existants
    existing_relations = get_relations(gab.name, gab.version)
    if existing_relations:
        st.caption("Enrichissements déclarés :")
        for r in existing_relations:
            with st.container(border=True):
                # visuel orienté enrichissement : SOURCE.left_key  →  ENRICHISSEMENT.right_key
                st.write(
                    f"**Source** : {r['from_gabarit']}[{r.get('from_version','v1')}] · clé `#{r['left_key']}`"
                )
                st.write(
                    f"**Enrichissement** : {r['to_gabarit']}[{r.get('to_version','v1')}] · clé `#{r['right_key']}`"
                )
                if st.button("Supprimer", key=f"del_enrich_{r['relation_id']}"):
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
        st.info("Aucun enrichissement déclaré pour ce gabarit.")


# =====================================================================
# DONNÉE PAR DÉFAUT (catalogue) — preview, python, validation
# =====================================================================
st.divider()
st.subheader("Donnée par défaut (optionnelle)")

if not edit_mode:
    st.info("Enregistrez d'abord le gabarit pour définir une donnée par défaut.")
else:
    gab = get_gabarit(gabarit.name, gabarit.version)
    expected_cols = [c.name for c in (gab.columns or [])]
    current_default = get_default_source(gab.name, gab.version) or {}

    with st.expander("Configurer la donnée par défaut", expanded=False):
        use_default = st.checkbox("Activer une donnée par défaut", value=bool(current_default))

        # === Saisie des paramètres de source ===
        fmt = (current_default or {}).get("type", "csv")
        path_val = (current_default or {}).get("path", "")
        sep_val = (current_default or {}).get("sep", ";")
        enc_val = (current_default or {}).get("encoding", "utf-8-sig")
        code_init = (current_default or {}).get("python", "")

        colF1, colF2 = st.columns([1, 2])
        with colF1:
            fmt = st.selectbox("Format", ["csv", "parquet"], index=0 if fmt == "csv" else 1)
        with colF2:
            path = st.text_input(
                "Chemin du fichier (local ou réseau)",
                value=path_val,
                placeholder=r"Ex: C:\data\sources\siren.csv  ou  /data/sources/siren.parquet",
                help="Indique un chemin accessible par le serveur Streamlit. Aucun upload n'est effectué."
            )

        if fmt == "csv":
            colC1, colC2 = st.columns(2)
            with colC1:
                sep = st.text_input("Séparateur", value=sep_val)
            with colC2:
                enc = st.text_input("Encodage", value=enc_val)
        else:
            sep, enc = None, None

        # === Zone code Python (df -> df)
        st.subheader("Transformation Python (df → df)")
        python_code = st.text_area(
            "Code pandas",
            value=code_init,
            height=180,
            placeholder="# Exemples :\n# df = df.rename(columns={'A': 'B'})\n# df = df[df['col'] > 0]\n",
            key="default_src_code"
        )

        # === Boutons actions : Preview / Validation / Enregistrer ===
        colA, colB, colC, colD = st.columns([1, 1, 1, 1])

        # Feedback de statut (succès/échec)
        feedback_key = "default_src_feedback"

        # ---- Preview
        if colA.button("🔍 Preview (20)", use_container_width=True, disabled=not (use_default and path)):
            df, err = _try_load_default_source(fmt, path, sep, enc, head=20)
            if err:
                st.session_state[feedback_key] = ("error", f"Erreur chargement : {err}")
            else:
                # Appliquer le code Python (optionnel)
                df2, perr = _apply_python(df, python_code)
                if perr:
                    st.session_state[feedback_key] = ("error", perr)
                else:
                    st.caption("Aperçu des 20 premières lignes")
                    st.dataframe(df2, use_container_width=True)
                    st.caption("Profiling simple")
                    st.json(_profile_df(df2))
                    st.session_state[feedback_key] = ("success", f"Chargement OK · shape={df2.shape}")

        # ---- Validation structure non bloquante
        if colB.button("✅ Valider structure", use_container_width=True, disabled=not (use_default and path)):
            df, err = _try_load_default_source(fmt, path, sep, enc, head=2000)
            if err:
                st.session_state[feedback_key] = ("error", f"Erreur chargement : {err}")
            else:
                df2, perr = _apply_python(df, python_code)
                if perr:
                    st.session_state[feedback_key] = ("error", perr)
                else:
                    aligned, warns = align_df_to_expected_columns(df2.copy(), expected_cols)
                    st.success("Validation effectuée (non bloquante).")
                    st.json({
                        "expected": expected_cols,
                        "missing_columns": warns.get("missing", []),
                        "extra_columns": warns.get("extra", []),
                        "profile": _profile_df(df2),
                    })
                    st.session_state[feedback_key] = ("success", "Validation OK")

        # ---- Enregistrer la config par défaut (et mémoriser l'aperçu 20 lignes)
        if colC.button("💾 Enregistrer la donnée par défaut", use_container_width=True, disabled=not (use_default and path)):
            # 1) Charger un tout petit échantillon (20 lignes) + appliquer le code Python
            df20, err = _try_load_default_source(fmt, path, sep, enc, head=20)
            if err:
                st.session_state[feedback_key] = ("error", f"Erreur chargement pour sauvegarde: {err}")
            else:
                df20, perr = _apply_python(df20, python_code)
                if perr:
                    st.session_state[feedback_key] = ("error", perr)
                else:
                    # 2) Sauvegarder la source
                    src = {"type": fmt, "path": str(Path(path).resolve())}
                    if fmt == "csv":
                        src.update({
                            "sep": sep or ";",
                            "encoding": enc or "utf-8-sig",
                        })
                    if python_code and python_code.strip():
                        src["python"] = python_code

                    set_default_source(gab.name, gab.version, src)

                    # 3) Mémoriser l’aperçu persistant (20 lignes)
                    try:
                        sample = df20.head(20)
                        set_default_preview(
                            gab.name, gab.version,
                            rows=sample.to_dict(orient="records"),
                            columns=list(sample.columns)
                        )
                    except Exception as e:
                        st.warning(f"Source enregistrée, mais l'aperçu n'a pas pu être mémorisé : {e}")

                    st.success("Donnée par défaut enregistrée avec aperçu persistant (20 lignes).")
                    st.session_state[feedback_key] = ("success", "Enregistré")
                    st.rerun()

        # ---- Retirer
        if colD.button("🧹 Retirer la donnée par défaut", use_container_width=True, disabled=not bool(current_default)):
            clear_default_source(gab.name, gab.version)
            st.success("Donnée par défaut retirée.")
            st.session_state[feedback_key] = ("success", "Supprimée")
            st.rerun()

        # ---- Indicateur de succès/échec global
        fb = st.session_state.get(feedback_key)
        if fb:
            kind, msg = fb
            if kind == "success":
                st.success(msg)
            else:
                st.error(msg)
