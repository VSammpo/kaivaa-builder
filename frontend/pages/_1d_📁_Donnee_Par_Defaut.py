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

# ---- Utils ----
def _make_json_safe(obj):
    import numpy as np
    import pandas as pd
    if isinstance(obj, dict):
        return {k: _make_json_safe(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_make_json_safe(item) for item in obj]
    elif isinstance(obj, (pd.Timestamp, pd.DatetimeTZDtype)):
        return obj.isoformat() if pd.notna(obj) else None
    elif isinstance(obj, (np.integer, np.floating)):
        return obj.item()
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    try:
        import pandas as pd  # noqa
        if pd.isna(obj):
            return None
    except Exception:
        pass
    return obj

def _ns(key: str, gab_name: str, gab_version: str) -> str:
    """Namespacer pour les clés de session (évite les collisions inter-gabarits)."""
    return f"{key}__{gab_name}__{gab_version}"

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

# ==== Détecter si on a changé de gabarit pour forcer le rechargement
k_last_gab = "last_viewed_gabarit"
current_gab_id = f"{gab_name}|{gab_version}"
if k_last_gab in st.session_state and st.session_state[k_last_gab] != current_gab_id:
    # On a changé de gabarit, nettoyer tous les flags hydrated
    keys_to_clean = [k for k in st.session_state.keys() if "hydrated_once__" in k]
    for k in keys_to_clean:
        del st.session_state[k]
st.session_state[k_last_gab] = current_gab_id

# ==== Clés de session namespacées
k_enabled = _ns("default_enabled", gab_name, gab_version)
k_mode    = _ns("default_mode", gab_name, gab_version)              # "file" | "python_only"
k_fmt     = _ns("default_fmt", gab_name, gab_version)               # "csv" | "parquet" | "python_only"
k_path    = _ns("default_path", gab_name, gab_version)
k_sep     = _ns("default_sep", gab_name, gab_version)
k_enc     = _ns("default_enc", gab_name, gab_version)
k_codebuf = _ns("python_code_buffer", gab_name, gab_version)
k_hydrated= _ns("hydrated_once", gab_name, gab_version)
k_post    = _ns("postsave_state", gab_name, gab_version)
k_flash   = _ns("flash_msg", gab_name, gab_version)


# ==== 1) APPLIQUER LES ÉTATS "PENDING" AVANT TOUT WIDGET
# (on peut mettre à jour session_state ici sans erreur)
if k_post in st.session_state:
    pending = st.session_state[k_post]
    for kk, vv in pending.items():
        st.session_state[kk] = vv
    del st.session_state[k_post]

# ==== 1b) Afficher un flash éventuel (après application du postsave_state, avant widgets)
if k_flash in st.session_state:
    st.success(st.session_state[k_flash])
    del st.session_state[k_flash]


# ==== 2) Charger l'état depuis le JSON au premier chargement de page
# La clé est de détecter si on arrive sur la page (pas de k_hydrated) ou si on est déjà dessus
current_default = get_default_source(gabarit.name, gabarit.version) or {}

# Recharger UNIQUEMENT si on vient d'arriver sur la page (k_hydrated absent)
# ET qu'il n'y a pas de postsave en cours
if k_hydrated not in st.session_state and k_post not in st.session_state:
    if current_default:
        mode = "file" if current_default.get("path") else (
            "python_only" if current_default.get("type") in {"python_only", "script"} else "file"
        )
        st.session_state[k_enabled] = True
        st.session_state[k_mode]    = mode
        st.session_state[k_fmt]     = current_default.get("type") if current_default.get("type") in {"csv","parquet"} else ("python_only" if mode=="python_only" else "csv")
        st.session_state[k_path]    = current_default.get("path", "")
        st.session_state[k_sep]     = current_default.get("sep", ";")
        st.session_state[k_enc]     = current_default.get("encoding", "utf-8-sig")
        st.session_state[k_codebuf] = current_default.get("python", "")
    else:
        st.session_state[k_enabled] = False
        st.session_state[k_mode]    = "file"
        st.session_state[k_fmt]     = "csv"
        st.session_state[k_path]    = ""
        st.session_state[k_sep]     = ";"
        st.session_state[k_enc]     = "utf-8-sig"
        st.session_state[k_codebuf] = ""
    st.session_state[k_hydrated] = True

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

# ==== Bandeau résumé du default enregistré
with st.expander("Résumé de la donnée par défaut enregistrée (lecture seule)", expanded=bool(current_default)):
    if current_default:
        colA, colB, colC, colD = st.columns([2,2,2,2])
        with colA:
            st.write("**Type** :", current_default.get("type", "—"))
            st.write("**Séparateur** :", current_default.get("sep", "—"))
        with colB:
            st.write("**Encodage** :", current_default.get("encoding", "—"))
            st.write("**Chemin** :", current_default.get("path", "—"))
        with colC:
            code_len = len(current_default.get("python", "") or "")
            st.write("**Script Python** :", f"{code_len} caractères")
        with colD:
            if st.button("↩️ Charger dans le formulaire", use_container_width=True):
                # On pousse dans la session PUIS rerun (widgets pas encore créés au prochain run)
                st.session_state[_ns("postsave_state", gab_name, gab_version)] = {
                    k_enabled: True,
                    k_mode: "file" if current_default.get("path") else "python_only",
                    k_fmt:  current_default.get("type") if current_default.get("type") in {"csv","parquet"} else ("python_only" if not current_default.get("path") else "csv"),
                    k_path: current_default.get("path", ""),
                    k_sep:  current_default.get("sep", ";"),
                    k_enc:  current_default.get("encoding", "utf-8-sig"),
                    k_codebuf: current_default.get("python", ""),
                }
                st.rerun()
    else:
        st.info("Aucune donnée par défaut enregistrée pour ce gabarit.")

# ==== UI principale
use_default = st.checkbox("Activer une donnée par défaut", key=k_enabled, value=st.session_state.get(k_enabled, False))

# Choix du mode
if use_default:
    radio_index = 0 if st.session_state.get(k_mode) == "file" else 1
    mode_label = st.radio(
        "Mode de création",
        ["📁 Fichier source", "🐍 Script Python uniquement"],
        index=radio_index,
        horizontal=True,
        key=_ns("radio_mode", gab_name, gab_version)
    )
    st.session_state[k_mode] = "file" if mode_label == "📁 Fichier source" else "python_only"
    use_file = (st.session_state[k_mode] == "file")

    if use_file:
        col1, col2 = st.columns([1,3])
        with col1:
            fmt_index = 0 if st.session_state.get(k_fmt, "csv") == "csv" else 1
            fmt_label = st.selectbox("Format", ["csv","parquet"], index=fmt_index, key=_ns("sel_fmt", gab_name, gab_version))
            st.session_state[k_fmt] = fmt_label

        with col2:
            st.text_input(
                "Chemin du fichier",
                value=st.session_state.get(k_path, ""),
                placeholder=r"Ex: C:\data\sources\fichier.csv",
                help="Chemin accessible par le serveur",
                key=k_path  # clé directe pour que la valeur se mette à jour
            )

        if st.session_state[k_fmt] == "csv":
            csep, cenc = st.columns(2)
            with csep:
                st.text_input("Séparateur", value=st.session_state.get(k_sep, ";"), key=k_sep)
            with cenc:
                st.text_input("Encodage", value=st.session_state.get(k_enc, "utf-8-sig"), key=k_enc)
        else:
            # on laisse les anciennes valeurs, elles ne seront pas utilisées
            pass
    else:
        st.info("💡 Mode script Python : créez un DataFrame `df` directement dans le code")
        # on ne vide pas ici pour éviter un set après widget — la sauvegarde s'occupera de tout

    expanded = bool(st.session_state.get(k_codebuf))
    with st.form(f"default_data_form_{gab_name}_{gab_version}", clear_on_submit=False, border=True):

        with st.expander("Transformation Python" + (" (obligatoire)" if not use_file else " (optionnel)"), expanded=expanded):
            st.caption("💡 Variables disponibles : `df` (DataFrame si fichier), `pd` (pandas)")
            if not use_file:
                st.caption("⚠️ Vous devez créer `df` (ex. `df = pd.DataFrame({...})`)")
            else:
                st.caption("⚠️ Vous devez réassigner `df` (ex. `df = df[['col1','col2']]`)")

            custom_buttons = [{
                "name": "Copier", "feather": "Copy", "hasText": True,
                "commands": ["copyAll"], "style": {"top": "0.46rem", "right": "0.4rem"}
            }]

            editor_result = code_editor(
                st.session_state.get(k_codebuf, ""),
                lang="python", height=300, theme="contrast", shortcuts="vscode",
                focus=False, buttons=custom_buttons, allow_reset=True,
                options={
                    "wrap": True, "showLineNumbers": True, "highlightActiveLine": True,
                    "enableLiveAutocompletion": True, "enableBasicAutocompletion": True,
                },
                key=_ns("python_code_editor", gab_name, gab_version),
                response_mode=["submit", "blur"]
            )

            if editor_result:
                new_code = None
                if isinstance(editor_result, dict):
                    new_code = (editor_result.get("text")
                                or editor_result.get("content")
                                or editor_result.get("code"))
                elif isinstance(editor_result, str):
                    new_code = editor_result
                if isinstance(new_code, str):
                    st.session_state[k_codebuf] = new_code

            current_code = st.session_state.get(k_codebuf, "")
            st.caption(f"📝 Code capturé : {len(current_code)} caractères")

        st.divider()

        can_preview = (use_file and bool(st.session_state.get(k_path))) or (not use_file and bool(st.session_state.get(k_codebuf, "").strip()))
        col_preview, col_validate, col_save = st.columns(3)
        with col_preview:
            do_preview = st.form_submit_button("👁️ Aperçu", use_container_width=True, disabled=not can_preview)
        with col_validate:
            do_validate = st.form_submit_button("✅ Valider", use_container_width=True, disabled=not can_preview)
        with col_save:
            do_save = st.form_submit_button("💾 Enregistrer", type="primary", use_container_width=True, disabled=not can_preview)

    # --- TRAITEMENT APRÈS FORM ---
    if use_file:
        fmt = st.session_state.get(k_fmt, "csv")
        path = st.session_state.get(k_path, "")
        sep  = st.session_state.get(k_sep, ";") if fmt == "csv" else None
        enc  = st.session_state.get(k_enc, "utf-8-sig") if fmt == "csv" else None
    else:
        fmt, path, sep, enc = "python_only", "", None, None

    if do_preview:
        with st.spinner("Chargement..."):
            if use_file and path:
                df, err = _try_load_source(fmt, path, sep, enc, head=20)
                if err:
                    st.error(f"❌ {err}")
                    df = None
            else:
                df = pd.DataFrame()

            if df is not None:
                current_code = st.session_state.get(k_codebuf, "")
                if current_code.strip():
                    st.info(f"📝 Application du script ({len(current_code)} caractères)")
                    df2, perr = _apply_python(df, current_code)
                    if perr:
                        st.error(f"❌ {perr}")
                    else:
                        st.success(f"✅ Aperçu chargé : {df2.shape[0]} lignes × {df2.shape[1]} colonnes")
                        st.dataframe(df2, use_container_width=True, height=300)
                elif use_file:
                    st.success(f"✅ Aperçu chargé : {df.shape[0]} lignes × {df.shape[1]} colonnes")
                    st.dataframe(df, use_container_width=True, height=300)

    if do_validate:
        with st.spinner("Validation..."):
            if use_file and path:
                df, err = _try_load_source(fmt, path, sep, enc, head=100)
                if err:
                    st.error(f"❌ {err}")
                    df = None
            else:
                df = pd.DataFrame()

            if df is not None:
                current_code = st.session_state.get(k_codebuf, "")
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
            if use_file and path:
                df20, err = _try_load_source(fmt, path, sep, enc, head=20)
                if err:
                    st.error(f"❌ {err}")
                    df20 = None
            else:
                df20 = pd.DataFrame()

            if df20 is not None:
                current_code = st.session_state.get(k_codebuf, "")
                df20_transformed, perr = _apply_python(df20, current_code)
                if perr:
                    st.error(f"❌ {perr}")
                else:
                    # Construire la source (ce qui sera persisté EN DISQUE)
                    if use_file and path:
                        src = {"type": st.session_state.get(k_fmt, "csv"), "path": str(Path(path).resolve())}
                        if st.session_state.get(k_fmt, "csv") == "csv":
                            src.update({"sep": st.session_state.get(k_sep, ";"), "encoding": st.session_state.get(k_enc, "utf-8-sig")})
                    else:
                        src = {"type": "python_only"}

                    if current_code and current_code.strip():
                        src["python"] = current_code

                    # 1) Persistance JSON
                    set_default_source(gabarit.name, gabarit.version, src)

                    # 2) Sauvegarde du preview (20 lignes)
                    sample = df20_transformed.head(20)
                    safe_rows = _make_json_safe(sample.to_dict(orient="records"))
                    set_default_preview(
                        gabarit.name, gabarit.version,
                        rows=safe_rows,
                        columns=list(sample.columns)
                    )

                    # 3) Préparer l'état post-save puis rerun (APPLIQUÉ AVANT WIDGETS AU PROCHAIN RUN)
                    pending_state = {
                        k_enabled: True,
                        k_mode: "file" if use_file else "python_only",
                        k_fmt:  src.get("type", "csv"),
                        k_path: src.get("path", ""),
                        k_sep:  src.get("sep", ";") if src.get("type") == "csv" else st.session_state.get(k_sep, ";"),
                        k_enc:  src.get("encoding", "utf-8-sig") if src.get("type") == "csv" else st.session_state.get(k_enc, "utf-8-sig"),
                        k_codebuf: src.get("python", ""),
                    }
                    st.session_state[k_post] = pending_state
                    st.session_state[k_flash] = "✅ Donnée par défaut enregistrée avec aperçu"
                    st.rerun()


# --- Bouton Retirer (hors formulaire) ---
if use_default:
    with st.container():
        if st.button("🗑️ Retirer", use_container_width=True, disabled=not (current_default or st.session_state.get(k_enabled))):
            clear_default_source(gabarit.name, gabarit.version)
            # On nettoie via un état post-save vide pour éviter le set après widget
            st.session_state[k_post] = {
                k_enabled: False,
                k_mode: "file",
                k_fmt: "csv",
                k_path: "",
                k_sep: ";",
                k_enc: "utf-8-sig",
                k_codebuf: "",
            }
            st.session_state[k_flash] = "✅ Donnée par défaut retirée"
            st.rerun()