# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from code_editor import code_editor

# ==== Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

# ==== Services
from backend.services.gabarit_registry import (
    get_gabarit, get_default_source, set_default_source,
    clear_default_source, set_default_preview
)

# align helper (fallback si service absent)
try:
    from backend.services.dataset_service import align_df_to_expected_columns
except Exception:
    def align_df_to_expected_columns(df: pd.DataFrame, expected_columns):
        expected = [c for c in (expected_columns or []) if isinstance(c, str) and c.strip()]
        cur_cols = list(df.columns)
        missing = [c for c in expected if c not in cur_cols]
        for c in missing:
            df[c] = pd.NA
        ordered = expected + [c for c in df.columns if c not in expected]
        return df[ordered], {"missing": missing, "extra": [c for c in cur_cols if c not in expected]}

st.set_page_config(page_title="Donnée par défaut", page_icon="📁", layout="wide")

# ========= Navbar homogène
def render_gabarit_subnav(active: str):
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

# ==== Sélection gabarit obligatoire
if "selected_gabarit" not in st.session_state or not st.session_state.selected_gabarit:
    st.error("Aucun gabarit sélectionné.")
    if st.button("← Retour aux gabarits", use_container_width=True):
        st.switch_page("pages/3_🧱_Gabarits.py")
    st.stop()

gab_name, gab_version = st.session_state.selected_gabarit
gabarit = get_gabarit(gab_name, gab_version)

render_gabarit_subnav("default")
st.title(f"📁 Donnée par défaut : {gabarit.name} [{gabarit.version}]")
st.divider()

# ==== Helpers locaux
def _try_load_source(fmt: str, path: str, sep: str | None, enc: str | None, head: int = 20):
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
    """
    Applique un script Python sur un DataFrame.
    Le script DOIT réassigner `df` pour que les modifications soient prises en compte.
    """
    if not code or not isinstance(code, str) or not code.strip():
        return df, None
    try:
        loc = {"df": df.copy(), "pd": pd}
        exec(code, {}, loc)
        new_df = loc.get("df")
        
        if not isinstance(new_df, pd.DataFrame):
            return None, (
                f"Le script doit retourner un DataFrame (pas {type(new_df).__name__}). "
                "Utilisez 'df = df[[...]]' (double crochets) pour garder un DataFrame"
            )
        
        return new_df, None
    
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        return None, f"Erreur script Python : {e}\n{tb}"

expected_cols = [c.name for c in (gabarit.columns or [])]
current_default = get_default_source(gabarit.name, gabarit.version) or {}

# ✅ INITIALISER le code Python avec une clé unique par gabarit
buffer_key = f"python_code_buffer_{gab_name}_{gab_version}"
if buffer_key not in st.session_state:
    st.session_state[buffer_key] = current_default.get("python", "")

use_default = st.checkbox("Activer une donnée par défaut", value=bool(current_default))

if use_default:
    col1, col2 = st.columns([1,3])
    with col1:
        fmt = st.selectbox("Format", ["csv","parquet"],
                           index=(0 if current_default.get("type") == "csv" else (1 if current_default.get("type")=="parquet" else 0)))
    with col2:
        path = st.text_input(
            "Chemin du fichier",
            value=current_default.get("path", ""),
            placeholder=r"Ex: C:\data\sources\fichier.csv",
            help="Chemin accessible par le serveur"
        )

    if fmt == "csv":
        csep, cenc = st.columns(2)
        with csep:
            sep = st.text_input("Séparateur", value=current_default.get("sep", ";"))
        with cenc:
            enc = st.text_input("Encodage", value=current_default.get("encoding", "utf-8-sig"))
    else:
        sep, enc = None, None

    expanded = bool(st.session_state.get(buffer_key) or current_default.get("python"))
    with st.form(f"default_data_form_{gab_name}_{gab_version}", clear_on_submit=False, border=True):

        with st.expander("Transformation Python (optionnel)", expanded=expanded):
            st.caption("💡 Variables disponibles : `df` (DataFrame), `pd` (pandas)")
            st.caption("⚠️ Vous devez réassigner `df` (ex. `df = df[['col1','col2']]` ou tout calcul renvoyant un DataFrame)")

            custom_buttons = [{
                "name": "Copier", "feather": "Copy", "hasText": True,
                "commands": ["copyAll"], "style": {"top": "0.46rem", "right": "0.4rem"}
            }]

            editor_result = code_editor(
                st.session_state[buffer_key],
                lang="python", height=300, theme="contrast", shortcuts="vscode",
                focus=False, buttons=custom_buttons, allow_reset=True,
                options={
                    "wrap": True, "showLineNumbers": True, "highlightActiveLine": True,
                    "enableLiveAutocompletion": True, "enableBasicAutocompletion": True,
                },
                key=f"python_code_editor_{gab_name}_{gab_version}",
                response_mode=["submit", "blur"]  # <-- capture AVANT le rerun
            )

            # Extraction robuste du contenu vers le buffer
            if editor_result:
                new_code = None
                if isinstance(editor_result, dict):
                    new_code = (editor_result.get("text")
                                or editor_result.get("content")
                                or editor_result.get("code"))
                elif isinstance(editor_result, str):
                    new_code = editor_result
                if isinstance(new_code, str):
                    st.session_state[buffer_key] = new_code

            current_code = st.session_state[buffer_key]
            st.caption(f"🔍 Code capturé : {len(current_code)} caractères")

        st.divider()
        col_preview, col_validate, col_save = st.columns(3)
        with col_preview:
            do_preview = st.form_submit_button("👁️ Aperçu", use_container_width=True, disabled=not path)
        with col_validate:
            do_validate = st.form_submit_button("✅ Valider", use_container_width=True, disabled=not path)
        with col_save:
            do_save = st.form_submit_button("💾 Enregistrer", type="primary", use_container_width=True, disabled=not path)

            pass

    # --- TRAITEMENT DES ACTIONS APRÈS LE FORM (la valeur de l'éditeur est déjà dans le buffer) ---
    if do_preview:
        with st.spinner("Chargement..."):
            df, err = _try_load_source(fmt, path, sep, enc, head=20)
            if err:
                st.error(f"❌ {err}")
            else:
                current_code = st.session_state[buffer_key]
                if current_code.strip():
                    st.info(f"🔍 Application du script ({len(current_code)} caractères)")
                df2, perr = _apply_python(df, current_code)
                if perr:
                    st.error(f"❌ {perr}")
                else:
                    st.success(f"✅ Aperçu chargé : {df2.shape[0]} lignes × {df2.shape[1]} colonnes")
                    st.dataframe(df2, use_container_width=True, height=300)

    if do_validate:
        with st.spinner("Validation..."):
            df, err = _try_load_source(fmt, path, sep, enc, head=100)
            if err:
                st.error(f"❌ {err}")
            else:
                current_code = st.session_state[buffer_key]
                df2, perr = _apply_python(df, current_code)
                if perr:
                    st.error(f"❌ {perr}")
                else:
                    aligned, warns = align_df_to_expected_columns(df2.copy(), expected_cols)
                    if warns.get("missing"):
                        st.warning(f"⚠️ Colonnes manquantes : {', '.join(warns['missing'])}")
                    if warns.get("extra"):
                        st.info(f"ℹ️ Colonnes supplémentaires : {', '.join(warns['extra'])}")
                    if not warns.get("missing") and not warns.get("extra"):
                        st.success("✅ Structure parfaitement alignée")

    if do_save:
        with st.spinner("Enregistrement..."):
            df20, err = _try_load_source(fmt, path, sep, enc, head=20)
            if err:
                st.error(f"❌ {err}")
            else:
                current_code = st.session_state[buffer_key]
                df20_transformed, perr = _apply_python(df20, current_code)
                if perr:
                    st.error(f"❌ {perr}")
                else:
                    src = {"type": fmt, "path": str(Path(path).resolve())}
                    if fmt == "csv":
                        src.update({"sep": sep or ";", "encoding": enc or "utf-8-sig"})
                    if current_code and current_code.strip():
                        src["python"] = current_code

                    set_default_source(gabarit.name, gabarit.version, src)

                    sample = df20_transformed.head(20)
                    set_default_preview(
                        gabarit.name, gabarit.version,
                        rows=sample.to_dict(orient="records"),
                        columns=list(sample.columns)
                    )
                    st.success("✅ Donnée par défaut enregistrée avec aperçu")
                    st.rerun()

    # --- Bouton Retirer (hors formulaire) ---
    with st.container():
        if st.button("🗑️ Retirer", use_container_width=True, disabled=not current_default):
            clear_default_source(gabarit.name, gabarit.version)
            if buffer_key in st.session_state:
                del st.session_state[buffer_key]
            st.success("✅ Donnée par défaut retirée")
            st.rerun()


else:
    if current_default:
        st.info("La donnée par défaut est désactivée. Cochez la case ci-dessus pour la reconfigurer.")