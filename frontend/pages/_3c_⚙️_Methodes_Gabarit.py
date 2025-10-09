# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from code_editor import code_editor

# ==== Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

# ==== Services (adapter si vos modules diffèrent)
from backend.services.gabarit_registry import (
    get_gabarit,
    list_methods_for_gabarit,   # ← au lieu de get_methods_for_gabarit
    upsert_method_for_gabarit,
    delete_method_for_gabarit,
    reorder_methods_for_gabarit,
    get_default_preview,
)

# Moteur d'exécution des méthodes
try:
    from backend.services.method_executor import apply_method, MethodExecutionError
except Exception as e:
    # Fallback de secours si le module backend n'est pas importable
    class MethodExecutionError(Exception):
        pass

    def apply_method(df, method_dict, params_dict):
        """Fallback basique : exécute directement le code Python stocké dans la méthode."""
        code = method_dict.get("code", "")
        outcol = method_dict.get("output_column", "result")
        loc = {"df": df.copy(), "pd": pd, "params": params_dict}
        try:
            exec(code, {}, loc)
            new_df = loc.get("df")
            if not isinstance(new_df, pd.DataFrame):
                raise MethodExecutionError("Le code de la méthode n'a pas produit un DataFrame nommé df.")
            return new_df
        except Exception as e:
            raise MethodExecutionError(str(e))

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

# ========= Sélection gabarit obligatoire
if "selected_gabarit" not in st.session_state or not st.session_state.selected_gabarit:
    st.error("Aucun gabarit sélectionné.")
    if st.button("← Retour aux gabarits", use_container_width=True):
        st.switch_page("pages/3_🧱_Gabarits.py")
    st.stop()

gab_name, gab_version = st.session_state.selected_gabarit
gabarit = get_gabarit(gab_name, gab_version)

render_gabarit_subnav("methods")
st.title(f"⚙️ Méthodes du gabarit — {gabarit.name} [{gabarit.version}]")

# ========= Helpers généraux =========
def _load_default_sample(gab):
    """Charge l'aperçu de donnée par défaut (20 lignes) si présent."""
    try:
        prev = get_default_preview(gab.name, gab.version) or {}
        rows = prev.get("rows") or []
        cols = prev.get("columns") or []
        if rows and cols:
            df = pd.DataFrame(rows, columns=cols)
            return df, None
        return None, None
    except Exception as e:
        return None, str(e)

def _get_sample_for_editor(gab):
    sample, _ = _load_default_sample(gab)
    return sample if isinstance(sample, pd.DataFrame) else pd.DataFrame()

def _infer_required_columns(code: str) -> list[str]:
    """Très simple : repère df['col'] ou df["col"] dans le code."""
    import re
    if not code: return []
    pat = r"df\[\s*['\"]([^'\"]+)['\"]\s*\]"
    return sorted(set(re.findall(pat, code)))

# ========= Paramètres dynamiques =========
def _safe_eval_options_python(code: str, df: pd.DataFrame) -> list[str]:
    if not code or not isinstance(code, str):
        return []
    env = {"df": df, "pd": pd}
    try:
        loc = {}
        exec(code, env, loc)
        options = loc.get("options", None)
        if options is None and "get_options" in loc and callable(loc["get_options"]):
            options = loc["get_options"](df)
        if isinstance(options, (list, tuple)):
            return [str(x) for x in options][:500]
        return []
    except Exception:
        return []

def _default_params_for_method(schema, df):
    vals = {}
    for spec in schema or []:
        nm = spec.get("name")
        tp = spec.get("type", "text")
        dv = spec.get("default", None)

        # Options pour select/select_multi
        opts = []
        if tp in ("select", "select_multi"):
            if spec.get("options"):
                if isinstance(spec["options"], list):
                    opts = [str(x) for x in spec["options"]]
                elif isinstance(spec["options"], str):
                    opts = [s.strip() for s in spec["options"].split(",") if s.strip()]
            if not opts and spec.get("source_column") and isinstance(df, pd.DataFrame):
                col = spec["source_column"]
                if col in df.columns:
                    opts = sorted(list(df[col].dropna().astype(str).unique()))[:500]
            if not opts and spec.get("options_python"):
                opts = _safe_eval_options_python(spec.get("options_python"), df)

        if tp == "number":
            try:
                vals[nm] = float(dv) if dv not in (None, "") else 0.0
            except Exception:
                vals[nm] = 0.0
        elif tp == "boolean":
            vals[nm] = str(dv).lower() in ("1", "true", "yes", "y", "on")
        elif tp == "select":
            if opts:
                vals[nm] = dv if str(dv) in [str(o) for o in opts] else opts[0]
            else:
                vals[nm] = str(dv) if dv is not None else ""
        elif tp == "select_multi":
            cur = dv if isinstance(dv, list) else ([dv] if dv not in (None, "") else [])
            cur = [str(x) for x in cur]
            if opts:
                cur = [x for x in cur if x in [str(o) for o in opts]]
                if not cur and opts:
                    cur = [str(opts[0])]
            vals[nm] = cur
        else:
            vals[nm] = "" if dv is None else str(dv)
    return vals

def _param_schema_from_form() -> list[dict]:
    out = []
    for p in st.session_state.form_param_list:
        name = (p.get("name") or "").strip()
        if not name: 
            continue
        item = {
            "name": name,
            "type": (p.get("type") or "text").strip(),
            "default": p.get("default"),
        }
        lbl = (p.get("label") or "").strip()
        if lbl:
            item["label"] = lbl

        if item["type"] in ("select", "select_multi"):
            src_mode = p.get("options_mode") or "manual"
            if src_mode == "manual":
                opts_csv = (p.get("options_text") or "").strip()
                if opts_csv:
                    item["options"] = [s.strip() for s in opts_csv.split(",") if s.strip()]
            elif src_mode == "column":
                col = (p.get("source_column") or "").strip()
                if col:
                    item["source_column"] = col
            elif src_mode == "python":
                py = (p.get("options_python") or "").strip()
                if py:
                    item["options_python"] = py
        out.append(item)
    return out

# ========= State init =========
st.session_state.setdefault("current_view", "list")
st.session_state.setdefault("editing_method_id", None)
st.session_state.setdefault("form_param_list", [])

# ========= Récupération des méthodes =========
methods = list_methods_for_gabarit(gabarit.name, gabarit.version) or []


# ========= ROUTEUR DE VUES =========
view = st.session_state.current_view

# --------------------------------- VUE LISTE ---------------------------------
if view == "list":
    st.subheader("Méthodes (ordre d'exécution)")
    c_top = st.container()
    with c_top:
        if st.button("➕ Nouvelle méthode", type="primary", use_container_width=True):
            st.session_state.editing_method_id = None
            st.session_state.form_param_list = []
            st.session_state.current_view = "edit"
            st.rerun()

    if not methods:
        st.info("Aucune méthode pour ce gabarit.")
    else:
        for m in sorted(methods, key=lambda x: x.get("order", 1)):
            with st.container(border=True):
                header = st.columns([5, 1, 1, 1, 1])
                with header[0]:
                    st.markdown(f"**{m.get('name','(sans nom)')}**  →  `out: {m.get('output_column','result')}`")
                    if m.get("description"):
                        st.caption(m["description"])
                    req = m.get("required_columns") or []
                    if req:
                        st.caption("Colonnes requises : " + ", ".join(f"`{c}`" for c in req))
                with header[1]:
                    if st.button("⬆️", key=f"up_{m['id']}", use_container_width=True, help="Monter"):
                        ids = [x["id"] for x in methods]; idx = ids.index(m["id"])
                        if idx > 0:
                            ids[idx-1], ids[idx] = ids[idx], ids[idx-1]
                            reorder_methods_for_gabarit(gabarit.name, gabarit.version, ids)
                        st.rerun()
                with header[2]:
                    if st.button("⬇️", key=f"down_{m['id']}", use_container_width=True, help="Descendre"):
                        ids = [x["id"] for x in methods]; idx = ids.index(m["id"])
                        if idx < len(ids)-1:
                            ids[idx+1], ids[idx] = ids[idx], ids[idx+1]
                            reorder_methods_for_gabarit(gabarit.name, gabarit.version, ids)
                        st.rerun()
                with header[3]:
                    if st.button("🧪 Tester", key=f"test_{m['id']}", use_container_width=True):
                        st.session_state.testing_method_id = m["id"]
                        st.session_state.current_view = "test"
                        st.rerun()

                # ➕ AJOUTER CETTE COLONNE SI ABSENTE
                with header[4]:
                    if st.button("✏️ Éditer", key=f"edit_{m['id']}", use_container_width=True):
                        # Pré-remplir la construction de param_schema pour l’UI
                        st.session_state.form_param_list = []
                        for p in (m.get("param_schema") or []):
                            entry = {
                                "name": p.get("name",""),
                                "label": p.get("label",""),
                                "type": p.get("type","text"),
                                "default": p.get("default",""),
                                "options_mode": "manual",
                                "options_text": "",
                                "source_column": "",
                                "options_python": "",
                            }
                            if entry["type"] in ("select","select_multi"):
                                if p.get("options_python"):
                                    entry["options_mode"] = "python"
                                    entry["options_python"] = p.get("options_python","")
                                elif p.get("source_column"):
                                    entry["options_mode"] = "column"
                                    entry["source_column"] = p.get("source_column","")
                                else:
                                    entry["options_mode"] = "manual"
                                    opts = p.get("options",[])
                                    entry["options_text"] = ",".join(opts) if isinstance(opts,list) else str(opts or "")
                            st.session_state.form_param_list.append(entry)

                        st.session_state.editing_method_id = m["id"]
                        st.session_state.current_view = "edit"
                        st.rerun()

            # Ligne actions secondaires
            cols = st.columns([1,1,3,3,2])
            with cols[0]:
                if st.button("🗑️ Supprimer", key=f"del_{m['id']}", use_container_width=True):
                    delete_method_for_gabarit(gabarit.name, gabarit.version, m["id"])
                    st.rerun()

        st.markdown("---")
        if st.button("▶️ Tester TOUTES les méthodes", use_container_width=True):
            st.session_state.current_view = "test_all"
            st.rerun()

# --------------------------------- VUE EDIT ----------------------------------
elif view == "edit":
    is_edit = st.session_state.editing_method_id is not None
    st.subheader("Création de méthode" if not is_edit else "Édition de méthode")

    # Pré-champs
    name = st.text_input("Nom de la méthode *", value="" if not is_edit else next((m["name"] for m in methods if m["id"]==st.session_state.editing_method_id), ""))
    outcol = st.text_input("Colonne de sortie *", value="" if not is_edit else next((m["output_column"] for m in methods if m["id"]==st.session_state.editing_method_id), ""))
    desc = st.text_area("Description (optionnel)", value="" if not is_edit else next((m.get("description","") for m in methods if m["id"]==st.session_state.editing_method_id), ""), height=80)

    # ✅ Code avec code_editor
    st.markdown("### 📝 Formule Python")
    st.caption("💡 Variables disponibles : `df` (DataFrame), `pd` (pandas), `params` (dict des paramètres)")
    st.caption("⚠️ Le DataFrame peut être modifié in-place (ex. `df['nouvelle_colonne'] = ...`) ; "
            "mais vous pouvez aussi écrire simplement une **expression** ou une **variable `value`**.")
    st.caption(f"💡 Si vous tapez **juste une expression**, elle sera automatiquement affectée à "
            f"`df['{(outcol or 'result').strip()}']`.")


    # ---- INITIALISATION PERSISTANTE (buffer + snapshot pour le badge) ----
    code_buffer_key = f"method_code_buffer_{st.session_state.editing_method_id or 'new'}"
    saved_snapshot_key = f"method_saved_snapshot_{st.session_state.editing_method_id or 'new'}"

    if code_buffer_key not in st.session_state:
        if is_edit:
            current_method = next((m for m in methods if m["id"] == st.session_state.editing_method_id), None)
            st.session_state[code_buffer_key] = current_method.get("code", "") if current_method else ""
        else:
            st.session_state[code_buffer_key] = ""

    if saved_snapshot_key not in st.session_state:
        # première valeur "officiellement sauvée" affichée comme référence
        st.session_state[saved_snapshot_key] = st.session_state[code_buffer_key]

    def _render_status_badge():
        is_dirty = st.session_state[code_buffer_key] != st.session_state[saved_snapshot_key]
        label = "🟡 Édition en cours" if is_dirty else "🟢 Sauvegardé"
        st.markdown(
            """
            <style>
            .pill{display:inline-block;padding:.2rem .5rem;border-radius:999px;
                font-size:.85rem;font-weight:600;border:1px solid rgba(0,0,0,.1);}
            .pill.saved{background:#e8fff0;}
            .pill.dirty{background:#fff8e6;}
            </style>
            """,
            unsafe_allow_html=True
        )
        st.markdown(f'<span class="pill {"dirty" if is_dirty else "saved"}">{label}</span>', unsafe_allow_html=True)

    _render_status_badge()

    # ---- ÉCHANTILLON POUR APERÇU/TEST (évite variables non définies) ----
    _editor_sample = _get_sample_for_editor(gabarit)          # helper défini plus haut dans ce fichier
    _sample_cols = list(_editor_sample.columns) if not _editor_sample.empty else []

    # ---- ÉDITEUR + BOUTONS regroupés dans UN FORM (synchro garantie) ----
    custom_buttons = [
        {"name":"Copier","feather":"Copy","hasText":True,"commands":["copyAll"],"style":{"top":"0.46rem","right":"0.4rem"}}
    ]

    with st.form(f"method_edit_form_{code_buffer_key}", clear_on_submit=False, border=True):
        editor_result = code_editor(
            st.session_state[code_buffer_key],
            lang="python",
            height=300,
            theme="contrast",
            shortcuts="vscode",
            focus=False,
            buttons=custom_buttons,
            allow_reset=True,
            options={
                "wrap": True,
                "showLineNumbers": True,
                "highlightActiveLine": True,
                "enableLiveAutocompletion": True,
                "enableBasicAutocompletion": True,
            },
            key=f"method_code_editor_{st.session_state.editing_method_id or 'new'}",
            response_mode=["submit","blur"]  # << capture AVANT le rerun
        )

        # Capture robuste -> met à jour le buffer AVANT de traiter les clics
        if editor_result:
            new_code = None
            if isinstance(editor_result, dict):
                new_code = editor_result.get("text") or editor_result.get("content") or editor_result.get("code")
            elif isinstance(editor_result, str):
                new_code = editor_result
            if isinstance(new_code, str):
                st.session_state[code_buffer_key] = new_code

        st.caption(f"🔍 Code capturé : {len(st.session_state[code_buffer_key])} caractères")

        c1, c2, c3 = st.columns(3)
        with c1:
            do_preview = st.form_submit_button("👁️ Aperçu", use_container_width=True)
        with c2:
            do_test = st.form_submit_button("🧪 Tester", use_container_width=True)
        with c3:
            do_save = st.form_submit_button("💾 Enregistrer", type="primary", use_container_width=True)

    # ---- TRAITEMENT DES ACTIONS (après le form) ----
    code = st.session_state[code_buffer_key]
    try:
        req_cols = _infer_required_columns(code)  # helper déjà présent plus haut chez toi
        if not isinstance(req_cols, list):
            req_cols = list(req_cols) if req_cols is not None else []
    except NameError:
        req_cols = []
    except Exception:
        req_cols = []


    # Colonnes référencées dans le code (ex. df['x'] → 'x')
    req_cols = _infer_required_columns(code)  # helper défini plus haut

    # petit utilitaire : s’assure que la colonne de sortie existe pour l’aperçu
    def _ensure_outcol(df: pd.DataFrame, out: str | None) -> pd.DataFrame:
        if not isinstance(df, pd.DataFrame):
            df = pd.DataFrame()
        oc = (out or "").strip() or "result"
        if oc not in df.columns:
            df = df.copy()
            df[oc] = None
        return df

    # ---- APERÇU (rendu visible) ----
    if do_preview:
        try:
            df_prev = _editor_sample.copy()
            # On appelle le moteur comme en prod (supporte df[out]=..., value=..., expression seule)
            method_dict = {
                "name": name or "(preview)",
                "description": desc or "",
                "output_column": (outcol or "result").strip(),
                "param_schema": _param_schema_from_form(),
                "required_columns": req_cols,
                "code": code or "",
            }
            params_vals = _default_params_for_method(method_dict["param_schema"], df_prev)
            df_prev = apply_method(df_prev, method_dict, params_vals)

            st.success("Aperçu généré sur la donnée par défaut.")
            if outcol and outcol in df_prev.columns:
                st.dataframe(df_prev[[outcol]].head(20), use_container_width=True)
            else:
                st.dataframe(df_prev.head(20), use_container_width=True)
        except MethodExecutionError as e:
            st.error(f"Erreur pendant l’aperçu : {e}")
        except Exception as e:
            st.error(f"Erreur inattendue pendant l’aperçu : {e}")


    # ---- TEST (sandbox + log visible) ----
    if do_test:
        try:
            df_test = _editor_sample.copy()
            method_dict = {
                "name": name or "(test)",
                "description": desc or "",
                "output_column": (outcol or "result").strip(),
                "param_schema": _param_schema_from_form(),
                "required_columns": req_cols,
                "code": code or "",
            }
            params_vals = _default_params_for_method(method_dict["param_schema"], df_test)
            df_test = apply_method(df_test, method_dict, params_vals)

            st.success("Test exécuté (sandbox).")
            with st.expander("Voir le DataFrame test"):
                st.dataframe(df_test.head(50), use_container_width=True)
        except MethodExecutionError as e:
            st.error(f"Erreur pendant le test : {e}")
        except Exception as e:
            st.error(f"Erreur inattendue pendant le test : {e}")


    # ---- SAVE (upsert + badge → 'Sauvegardé') ----
    if do_save:
        try:
            # Ordre courant (préserve l’ordre si édition)
            current_order = None
            if is_edit:
                for mm in methods:
                    if mm["id"] == st.session_state.editing_method_id:
                        current_order = mm.get("order", 1)
                        break

            saved = upsert_method_for_gabarit(
                gabarit.name, gabarit.version,
                name=name.strip(),
                description=(desc or "").strip(),
                output_column=outcol.strip(),
                param_schema=_param_schema_from_form(),
                required_columns=req_cols,         # << remplace 'inferred'
                code=code,
                order=current_order,
                method_id=st.session_state.editing_method_id
            )

            # Maj du snapshot => le badge passe au vert
            st.session_state[saved_snapshot_key] = st.session_state[code_buffer_key]
            st.success("Méthode enregistrée ✅")
            st.toast("Sauvegarde effectuée", icon="💾")
            _render_status_badge()

        except Exception as e:
            st.error(f"Échec sauvegarde : {e}")



    to_delete = []
    for idx, p in enumerate(st.session_state.form_param_list):
        with st.expander(f"Paramètre #{idx+1} — {(p.get('name') or 'sans nom')}", expanded=True):
            colA, colB, colC = st.columns([2,2,1])
            with colA:
                p["name"] = st.text_input("Nom *", value=p.get("name",""), key=f"p_name_{idx}")
            with colB:
                p["label"] = st.text_input("Label", value=p.get("label",""), key=f"p_label_{idx}")
            with colC:
                p["type"]  = st.selectbox("Type", ["text","number","boolean","select","select_multi"],
                                          index=["text","number","boolean","select","select_multi"].index(p.get("type","text")),
                                          key=f"p_type_{idx}")

            if p["type"] in ("select","select_multi"):
                st.caption("Source des options")
                mode = st.radio("Mode", ["manual","column","python"],
                                index=["manual","column","python"].index(p.get("options_mode","manual")),
                                key=f"p_mode_{idx}", horizontal=True, label_visibility="collapsed")
                p["options_mode"] = mode

                if mode == "manual":
                    p["options_text"] = st.text_area(
                        "Options (séparées par des virgules)",
                        value=p.get("options_text",""),
                        key=f"p_opts_text_{idx}", height=80
                    )
                    opts_preview = [s.strip() for s in (p.get("options_text") or "").split(",") if s.strip()]
                    if opts_preview:
                        st.caption(f"Prévisualisation : {len(opts_preview)} option(s)")

                elif mode == "column":
                    p["source_column"] = st.selectbox(
                        "Colonne source",
                        options=_sample_cols if _sample_cols else ["(aucune donnée par défaut disponible)"],
                        index=(_sample_cols.index(p.get("source_column")) if p.get("source_column") in _sample_cols else 0) if _sample_cols else 0,
                        key=f"p_source_col_{idx}", disabled=not _sample_cols
                    )
                    if _sample_cols and p.get("source_column") in _sample_cols:
                        uniques = sorted(list(_editor_sample[p["source_column"]].dropna().astype(str).unique()))
                        st.caption(f"Prévisualisation : {min(len(uniques), 500)} valeur(s) unique(s)")

                elif mode == "python":
                    st.caption("Le code doit définir `options = [...]` ou `def get_options(df): ...`")
                    p["options_python"] = st.text_area(
                        "Code Python des options", value=p.get("options_python",""),
                        key=f"p_opts_py_{idx}", height=140,
                        placeholder="options = sorted(df['col'].astype(str).unique().tolist())[:200]"
                    )
                    if st.button("▶️ Tester le code", key=f"p_opts_py_test_{idx}"):
                        preview = _safe_eval_options_python(p.get("options_python",""), _editor_sample)
                        st.caption(f"Prévisualisation : {len(preview)} option(s)")
                        if preview[:10]:
                            st.write(preview[:10])

                # Valeur(s) par défaut selon les options
                st.markdown("**Valeur par défaut**")
                widget_opts = []
                if mode == "manual":
                    widget_opts = [s.strip() for s in (p.get("options_text") or "").split(",") if s.strip()]
                elif mode == "column" and _sample_cols and p.get("source_column") in _sample_cols:
                    widget_opts = sorted(list(_editor_sample[p["source_column"]].dropna().astype(str).unique()))[:500]
                elif mode == "python" and p.get("options_python"):
                    widget_opts = _safe_eval_options_python(p.get("options_python",""), _editor_sample)

                if p["type"] == "select":
                    cur = p.get("default", "")
                    p["default"] = st.selectbox("Défaut", options=widget_opts if widget_opts else [cur],
                                                index=(widget_opts.index(cur) if cur in widget_opts else 0) if widget_opts else 0,
                                                key=f"p_def_sel_{idx}")
                else:
                    cur = p.get("default", [])
                    if not isinstance(cur, list): cur = [cur] if cur not in (None,"") else []
                    p["default"] = st.multiselect("Défaut (multi)", options=widget_opts,
                                                  default=[x for x in cur if x in widget_opts],
                                                  key=f"p_def_ms_{idx}")

            elif p["type"] == "number":
                p["default"] = st.number_input("Valeur par défaut",
                                               value=float(p.get("default") or 0.0),
                                               key=f"p_def_num_{idx}")
            elif p["type"] == "boolean":
                p["default"] = st.checkbox("Valeur par défaut",
                                           value=bool(p.get("default") in (True, "true", "True", 1, "1", "on")),
                                           key=f"p_def_bool_{idx}")
            else:
                p["default"] = st.text_input("Valeur par défaut", value=str(p.get("default") or ""), key=f"p_def_text_{idx}")

            if st.button("🗑️ Supprimer ce paramètre", key=f"p_del_{idx}"):
                to_delete.append(idx)

    if to_delete:
        for i in sorted(to_delete, reverse=True):
            del st.session_state.form_param_list[i]
        st.rerun()
    # ---------------- /PARAMS DYNAMIQUES ----------------

    st.markdown("---")
    if st.button("💾 Enregistrer", type="primary", use_container_width=True, key="save_method_btn"):
        if not name.strip():
            st.error("Le nom de la méthode est requis")
        elif not outcol.strip():
            st.error("Le nom de la colonne de sortie est requis")
        else:
            try:
                current_order = None
                if is_edit:
                    for mm in methods:
                        if mm["id"] == st.session_state.editing_method_id:
                            current_order = mm.get("order", 1)
                            break

                saved = upsert_method_for_gabarit(
                    gabarit.name, gabarit.version,
                    name=name.strip(),
                    description=desc.strip(),
                    output_column=outcol.strip(),
                    param_schema=_param_schema_from_form(),
                    required_columns=req_cols,
                    code=code,
                    order=current_order,
                    method_id=st.session_state.editing_method_id
                )
                st.success(f"Méthode '{saved['name']}' enregistrée")
                
                # ✅ Nettoyer le buffer après sauvegarde
                if code_buffer_key in st.session_state:
                    del st.session_state[code_buffer_key]
                
                st.session_state.current_view = "list"
                st.rerun()
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement : {e}")

    if st.button("↩️ Annuler", use_container_width=True):
        # ✅ Nettoyer le buffer si on annule
        if code_buffer_key in st.session_state:
            del st.session_state[code_buffer_key]
        st.session_state.current_view = "list"
        st.rerun()

# ------------------------------- VUE TEST 1 -----------------------------------
elif view == "test":
    mid = st.session_state.get("testing_method_id")
    m = next((x for x in methods if x["id"] == mid), None)
    if not m:
        st.warning("Méthode introuvable.")
        st.session_state.current_view = "list"
        st.rerun()

    st.subheader(f"🧪 Test : {m.get('name','(sans nom)')}")
    sample, err = _load_default_sample(gabarit)
    if err:
        st.error(err)
    if sample is None:
        st.info("Aucune donnée par défaut n'est configurée pour ce gabarit.")
    else:
        if st.button("▶️ Exécuter le test", type="primary", use_container_width=True):
            try:
                with st.spinner("Exécution en cours..."):
                    pvals = _default_params_for_method(m.get("param_schema") or [], sample)
                    out = apply_method(sample, m, pvals)
                st.success("Méthode appliquée avec succès")
                if pvals:
                    st.caption("Paramètres utilisés : " + ", ".join([f"{k}={v}" for k,v in pvals.items()]))
                st.dataframe(out, use_container_width=True, height=500)
            except MethodExecutionError as e:
                st.error(f"Erreur d'exécution : {e}")
            except Exception as e:
                st.error(f"Erreur inattendue : {e}")

    st.markdown("---")
    if st.button("↩️ Retour aux méthodes", use_container_width=True):
        st.session_state.current_view = "list"
        st.rerun()

# ----------------------------- VUE TEST ALL -----------------------------------
elif view == "test_all":
    st.subheader("🧪 Test de la chaîne complète de méthodes")
    sample, err = _load_default_sample(gabarit)
    if err:
        st.error(err)
    if sample is None:
        st.info("Aucune donnée par défaut n'est configurée.")
    else:
        if st.button("▶️ Lancer le test complet", type="primary", use_container_width=True):
            try:
                df = sample.copy()
                for m in sorted(methods, key=lambda x: x.get("order", 1)):
                    pvals = _default_params_for_method(m.get("param_schema") or [], df)
                    df = apply_method(df, m, pvals)
                st.success("Chaîne exécutée avec succès")
                st.dataframe(df, use_container_width=True, height=500)
            except MethodExecutionError as e:
                st.error(f"Erreur d'exécution : {e}")
            except Exception as e:
                st.error(f"Erreur inattendue : {e}")

    st.markdown("---")
    if st.button("↩️ Retour aux méthodes", use_container_width=True):
        st.session_state.current_view = "list"
        st.rerun()