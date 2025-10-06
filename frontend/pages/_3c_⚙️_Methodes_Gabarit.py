# frontend/pages/_3c_⚙️_Methodes_Gabarit.py
import re
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

# ==== Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))
st.set_page_config(page_title="Méthodes du gabarit", page_icon="⚙️", layout="wide")

# ==== Services
from backend.services.gabarit_registry import (
    get_gabarit, list_gabarits, get_default_source,
    list_methods_for_gabarit, upsert_method_for_gabarit,
    delete_method_for_gabarit, reorder_methods_for_gabarit,
)
from backend.services.method_executor import apply_method, MethodExecutionError

# ========================= CSS =========================
st.markdown("""
<style>
.badge{ 
    display:inline-block; 
    padding:2px 8px; 
    border-radius:12px; 
    background:#eef2ff; 
    margin-right:6px; 
    font-size:12px;
    font-weight: 500;
}
.badge-warn{ background:#fff3cd; color: #856404; }
.badge-ok{ background:#d4edda; color: #155724; }
</style>
""", unsafe_allow_html=True)


# ========================= Helpers =========================

# --- Helper: exécuter un code python qui renvoie une liste d'options ---
def _safe_eval_options_python(code: str, df: pd.DataFrame) -> list[str]:
    if not code or not isinstance(code, str):
        return []
    env = {"df": df, "pd": pd}
    try:
        loc = {}
        exec(code, env, loc)
        # Convention : l'utilisateur doit définir une variable 'options'
        options = loc.get("options", None)
        if options is None and "get_options" in loc and callable(loc["get_options"]):
            options = loc["get_options"](df)
        if isinstance(options, (list, tuple)):
            return [str(x) for x in options][:500]
        return []
    except Exception:
        return []

# --- Helper: charger un sample juste pour l'éditeur (colonnes / preview) ---
def _get_sample_for_editor(gabarit):
    sample, _err = _load_default_sample(gabarit)
    # sample peut être None; on retourne un DF vide mais avec .columns
    if sample is None:
        sample = pd.DataFrame()
    return sample


def _infer_required_columns(code: str) -> list[str]:
    if not code:
        return []
    try:
        cols = set()
        for m in re.finditer(r"df\[\s*(['\"])(.+?)\1\s*\]", code):
            cols.add(m.group(2).strip())
        for m in re.finditer(r"df\.loc\[\s*:\s*,\s*(['\"])(.+?)\1\s*\]", code):
            cols.add(m.group(2).strip())
        return sorted([c for c in cols if c])
    except re.error:
        return []

def _load_default_sample(gab):
    src = get_default_source(gab.name, gab.version) or {}
    if not src:
        return None, "Aucune donnée par défaut configurée."
    t, path = src.get("type"), src.get("path")
    try:
        if t == "csv":
            df = pd.read_csv(path, sep=src.get("sep",";"), encoding=src.get("encoding","utf-8-sig"))
        elif t == "parquet":
            df = pd.read_parquet(path)
        elif t == "excel":
            xl = pd.ExcelFile(path)
            sheet = src.get("sheet_name")
            df = xl.parse(sheet) if sheet else xl.parse(xl.sheet_names[0])
        else:
            return None, f"Type de source inconnu: {t}"
        if src.get("python"):
            loc={"df": df, "pd": pd}
            exec(src["python"], {}, loc)
            df = loc.get("df", df)
        return df.head(20), None
    except Exception as e:
        return None, f"Erreur lecture source par défaut: {e}"

def _default_params_for_method(schema, df):
    vals = {}
    for spec in schema or []:
        nm = spec.get("name")
        tp = spec.get("type", "text")
        dv = spec.get("default", None)

        # Préparer options si besoin (select / select_multi)
        opts = []
        if tp in ("select", "select_multi"):
            # source 1: manuel (options)
            if spec.get("options"):
                if isinstance(spec["options"], list):
                    opts = [str(x) for x in spec["options"]]
                elif isinstance(spec["options"], str):
                    opts = [s.strip() for s in spec["options"].split(",") if s.strip()]
            # source 2: colonne
            if not opts and spec.get("source_column") and isinstance(df, pd.DataFrame):
                col = spec["source_column"]
                if col in df.columns:
                    opts = sorted(list(df[col].dropna().astype(str).unique()))[:500]
            # source 3: python
            if not opts and spec.get("options_python"):
                opts = _safe_eval_options_python(spec.get("options_python"), df)

        # Attribution du default en fonction du type
        if tp == "number":
            try:
                vals[nm] = float(dv) if dv not in (None, "") else 0.0
            except Exception:
                vals[nm] = 0.0
        elif tp == "boolean":
            vals[nm] = str(dv).lower() in ("1", "true", "yes", "y", "on")
        elif tp == "select":
            # valeur par défaut cohérente avec les options
            if opts:
                vals[nm] = dv if str(dv) in [str(o) for o in opts] else opts[0]
            else:
                vals[nm] = str(dv) if dv is not None else ""
        elif tp == "select_multi":
            # liste de valeurs
            cur = dv if isinstance(dv, list) else ([dv] if dv not in (None, "") else [])
            cur = [str(x) for x in cur]
            if opts:
                # filtrer sur les options valides
                cur = [x for x in cur if x in [str(o) for o in opts]]
                if not cur and opts:
                    cur = [str(opts[0])]
            vals[nm] = cur
        else:
            # text et tout le reste par défaut
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

        # Sources d'options pour select / select_multi
        if item["type"] in ("select", "select_multi"):
            src_mode = p.get("options_mode") or "manual"
            if src_mode == "manual":
                opts_csv = (p.get("options_text") or "").strip()
                if opts_csv:
                    item["options"] = [s.strip() for s in opts_csv.split(",") if s.strip()]
            elif src_mode == "column":
                col = (p.get("source_column") or "").strip()
                if col:
                    # on stocke sous le champ standard utilisé côté exécution
                    item["source_column"] = col
            elif src_mode == "python":
                py = (p.get("options_python") or "").strip()
                if py:
                    item["options_python"] = py

        out.append(item)
    return out

def _apply_all_methods_on_sample(gabarit, sample: pd.DataFrame):
    df = sample.copy(); warns=[]; errs=[]
    for m in list_methods_for_gabarit(gabarit.name, gabarit.version):
        try:
            miss = [c for c in (m.get("required_columns") or []) if c not in df.columns]
            if miss:
                warns.append(f"{m['name']}: colonnes manquantes {miss}")
                continue
            pvals = _default_params_for_method(m.get("param_schema") or [], df)
            df = apply_method(df, m, pvals)
        except MethodExecutionError as e:
            warns.append(f"{m['name']}: {e}")
        except Exception as e:
            errs.append(f"{m['name']}: {e}")
    return df, warns, errs

# ========================= State =========================
def _ensure_state():
    ss = st.session_state
    ss.setdefault("current_view", "list")
    ss.setdefault("editing_method_id", None)
    ss.setdefault("form_name", "")
    ss.setdefault("form_output_col", "")
    ss.setdefault("form_desc", "")
    ss.setdefault("form_code", "")
    ss.setdefault("form_param_list", [])
    ss.setdefault("testing_method", None)
    ss.setdefault("last_gabarit_key", None)

_ensure_state()

st.title("⚙️ Méthodes du gabarit")

sel = st.session_state.get("selected_gabarit")
gabarit = get_gabarit(*sel) if sel else None
if not gabarit:
    gs = list_gabarits()
    opts = [f"{g.name}|{g.version}" for g in gs]
    pick = st.selectbox("Sélectionner un gabarit", opts) if opts else None
    if pick:
        n,v = pick.split("|",1)
        gabarit = get_gabarit(n,v)

if not gabarit:
    st.info("Choisissez d'abord un gabarit depuis sa page de détail.")
    st.stop()

gab_key = f"{gabarit.name}|{gabarit.version}"
if st.session_state.last_gabarit_key != gab_key:
    st.session_state.current_view = "list"
    st.session_state.editing_method_id = None
    st.session_state.testing_method = None
    st.session_state.last_gabarit_key = gab_key

st.markdown(f"**Gabarit :** `{gabarit.name}` • **Version :** `{gabarit.version}`")
g_cols = [c.name for c in (gabarit.columns or [])]
key_cols = [c.name for c in (gabarit.columns or []) if getattr(c,"is_key",False)]

col_info, col_back = st.columns([4, 1])
with col_back:
    if st.button("⬅️ Retour au gabarit", use_container_width=True):
        st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")

st.divider()

methods = list_methods_for_gabarit(gabarit.name, gabarit.version)

# ========================= VUE SIMPLE SANS TABS =========================
if st.session_state.current_view == "list":
    # LISTE DES MÉTHODES
    col_new, col_test_all = st.columns([1, 1])
    
    with col_new:
        if st.button("➕ Nouvelle méthode", type="primary", use_container_width=True):
            st.session_state.current_view = "edit"
            st.session_state.editing_method_id = None
            st.session_state.form_name = ""
            st.session_state.form_output_col = ""
            st.session_state.form_desc = ""
            st.session_state.form_code = ""
            st.session_state.form_param_list = []
            st.rerun()
    
    with col_test_all:
        if st.button("🧪 Tester toutes les méthodes", use_container_width=True):
            st.session_state.current_view = "test_all"
            st.rerun()
    
    st.write("")
    
    if not methods:
        st.info("💡 Aucune méthode définie. Créez votre première méthode pour commencer !")
    else:
        st.subheader(f"📊 {len(methods)} méthode(s) configurée(s)")
        
        for idx, m in enumerate(methods):
            with st.container(border=True):
                col_info, col_actions = st.columns([3, 1])
                
                with col_info:
                    st.markdown(f"### {idx+1}. {m['name']}")
                    st.markdown(f"**Colonne de sortie :** `{m.get('output_column', '?')}`")
                    
                    if m.get("description"):
                        st.caption(m["description"])
                    
                    badges = [f"<span class='badge'>Ordre: {m.get('order',1)}</span>"]
                    if m.get("required_columns"):
                        badges.append(f"<span class='badge'>Colonnes requises: {', '.join(m['required_columns'])}</span>")
                    if m.get("param_schema"):
                        badges.append(f"<span class='badge'>{len(m['param_schema'])} paramètre(s)</span>")
                    st.markdown(" ".join(badges), unsafe_allow_html=True)
                
                with col_actions:
                    if st.button("✏️ Éditer", key=f"edit_{m['id']}", use_container_width=True):
                        st.session_state.current_view = "edit"
                        st.session_state.editing_method_id = m.get("id")
                        st.session_state.form_name = m.get("name","")
                        st.session_state.form_output_col = m.get("output_column","")
                        st.session_state.form_desc = m.get("description","")
                        st.session_state.form_code = m.get("code","")
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

                        st.rerun()
                    
                    col_up, col_down = st.columns(2)
                    with col_up:
                        if st.button("⬆️", key=f"up_{m['id']}", use_container_width=True, disabled=idx==0):
                            ids = [x["id"] for x in methods]
                            ids[idx-1], ids[idx] = ids[idx], ids[idx-1]
                            reorder_methods_for_gabarit(gabarit.name, gabarit.version, ids)
                            st.rerun()
                    with col_down:
                        if st.button("⬇️", key=f"down_{m['id']}", use_container_width=True, disabled=idx==len(methods)-1):
                            ids = [x["id"] for x in methods]
                            ids[idx+1], ids[idx] = ids[idx], ids[idx+1]
                            reorder_methods_for_gabarit(gabarit.name, gabarit.version, ids)
                            st.rerun()
                    
                    if st.button("▶️ Tester", key=f"test_{m['id']}", use_container_width=True):
                        st.session_state.current_view = "test"
                        st.session_state.testing_method = m
                        st.rerun()
                    
                    if st.button("🗑️ Supprimer", key=f"del_{m['id']}", use_container_width=True):
                        if st.session_state.get(f"confirm_del_{m['id']}", False):
                            delete_method_for_gabarit(gabarit.name, gabarit.version, m["id"])
                            st.session_state[f"confirm_del_{m['id']}"] = False
                            st.success(f"Méthode '{m['name']}' supprimée")
                            st.rerun()
                        else:
                            st.session_state[f"confirm_del_{m['id']}"] = True
                            st.warning("Cliquez à nouveau pour confirmer la suppression")
                            st.rerun()

elif st.session_state.current_view == "edit":
    # ÉDITION
    is_new = st.session_state.editing_method_id is None
    
    col_title, col_back = st.columns([4, 1])
    with col_title:
        st.subheader("➕ Nouvelle méthode" if is_new else f"✏️ Modifier : {st.session_state.form_name}")
    with col_back:
        if st.button("⬅️ Liste", use_container_width=True):
            st.session_state.current_view = "list"
            st.rerun()
    
    name = st.text_input("Nom de la méthode *", value=st.session_state.form_name, 
                       key="edit_name", placeholder="ex: moyenne_glissante_simple")
    outcol = st.text_input("Colonne de sortie *", value=st.session_state.form_output_col, 
                         key="edit_outcol", placeholder="ex: prix_moyen")
    
    desc = st.text_area("Description", value=st.session_state.form_desc, key="edit_desc", height=80)
    
    st.markdown("---")
    st.subheader("📝 Formule Python")
    
    code = st.text_area("Code (définissez value = ...)", value=st.session_state.form_code, 
                      key="edit_code", height=200,
                      placeholder="# value = df['valeur'] / df['volume']")
    
    inferred = _infer_required_columns(code)
    if inferred:
        st.markdown("**Colonnes détectées :** " + " ".join([f"`{c}`" for c in inferred]))


    # ================== PARAMS DYNAMIQUES ==================
    st.markdown("---")
    st.subheader("⚙️ Paramètres (optionnel)")

    # sample pour listes basées sur colonnes / python
    _editor_sample = _get_sample_for_editor(gabarit)
    _sample_cols = list(_editor_sample.columns) if isinstance(_editor_sample, pd.DataFrame) else []

    # initialisation state
    if not st.session_state.form_param_list:
        st.session_state.form_param_list = []

    # bouton d'ajout
    if st.button("➕ Ajouter un paramètre", key="add_param_btn", use_container_width=True):
        st.session_state.form_param_list.append({
            "name": "",
            "label": "",
            "type": "text",                  # text | number | boolean | select | select_multi
            "default": "",
            # champs spécifiques select/select_multi
            "options_mode": "manual",        # manual | column | python
            "options_text": "",              # CSV si manuel
            "source_column": "",             # si column
            "options_python": "",            # si python (doit produire 'options' = list)
        })
        st.rerun()

    # édition param par param (expander)
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

            # zone dynamique selon type
            if p["type"] in ("select","select_multi"):
                st.caption("Source des options")
                mode = st.radio(
                    "Mode", ["manual","column","python"],
                    index=["manual","column","python"].index(p.get("options_mode","manual")),
                    key=f"p_mode_{idx}", horizontal=True, label_visibility="collapsed"
                )
                p["options_mode"] = mode

                if mode == "manual":
                    p["options_text"] = st.text_area(
                        "Options (séparées par des virgules)",
                        value=p.get("options_text",""),
                        key=f"p_opts_text_{idx}",
                        height=80
                    )
                    # aperçu options
                    opts_preview = [s.strip() for s in (p.get("options_text") or "").split(",") if s.strip()]
                    if opts_preview:
                        st.caption(f"Prévisualisation : {len(opts_preview)} option(s)")

                elif mode == "column":
                    p["source_column"] = st.selectbox(
                        "Colonne source",
                        options=_sample_cols if _sample_cols else ["(aucune donnée par défaut disponible)"],
                        index=(_sample_cols.index(p.get("source_column")) if p.get("source_column") in _sample_cols else 0) if _sample_cols else 0,
                        key=f"p_source_col_{idx}",
                        disabled=not _sample_cols
                    )
                    if _sample_cols and p.get("source_column") in _sample_cols:
                        uniques = sorted(list(_editor_sample[p["source_column"]].dropna().astype(str).unique()))
                        st.caption(f"Prévisualisation : {min(len(uniques), 500)} valeur(s) unique(s)")

                elif mode == "python":
                    st.caption("Le code doit définir une variable `options` (list[str]) ou une fonction `get_options(df)`.")
                    p["options_python"] = st.text_area(
                        "Code Python des options", value=p.get("options_python",""),
                        key=f"p_opts_py_{idx}", height=140, placeholder="options = sorted(df['col'].astype(str).unique().tolist())[:200]"
                    )
                    if st.button("▶️ Tester le code", key=f"p_opts_py_test_{idx}"):
                        preview = _safe_eval_options_python(p.get("options_python",""), _editor_sample)
                        st.caption(f"Prévisualisation : {len(preview)} option(s)")
                        if preview[:10]:
                            st.write(preview[:10])

                # valeur par défaut en fonction des options disponibles
                st.markdown("**Valeur par défaut**")
                # reconstituer les options déterministes pour le widget
                widget_opts = []
                if mode == "manual":
                    widget_opts = [s.strip() for s in (p.get("options_text") or "").split(",") if s.strip()]
                elif mode == "column" and _sample_cols and p.get("source_column") in _sample_cols:
                    widget_opts = sorted(list(_editor_sample[p["source_column"]].dropna().astype(str).unique()))[:500]
                elif mode == "python" and p.get("options_python"):
                    widget_opts = _safe_eval_options_python(p.get("options_python",""), _editor_sample)

                if p["type"] == "select":
                    # simple select
                    cur = p.get("default", "")
                    p["default"] = st.selectbox(
                        "Défaut", options=widget_opts if widget_opts else [cur],
                        index=(widget_opts.index(cur) if cur in widget_opts else 0) if widget_opts else 0,
                        key=f"p_def_sel_{idx}"
                    )
                else:
                    # multi select
                    cur = p.get("default", [])
                    if not isinstance(cur, list): cur = [cur] if cur not in (None,"") else []
                    p["default"] = st.multiselect(
                        "Défaut (multi)", options=widget_opts, default=[x for x in cur if x in widget_opts],
                        key=f"p_def_ms_{idx}"
                    )

            elif p["type"] == "number":
                p["default"] = st.number_input("Valeur par défaut", value=float(p.get("default") or 0.0), key=f"p_def_num_{idx}")
            elif p["type"] == "boolean":
                p["default"] = st.checkbox("Valeur par défaut", value=bool(p.get("default") in (True, "true", "True", 1, "1", "on")), key=f"p_def_bool_{idx}")
            else:
                p["default"] = st.text_input("Valeur par défaut", value=str(p.get("default") or ""), key=f"p_def_text_{idx}")

            # bouton supprimer
            if st.button("🗑️ Supprimer ce paramètre", key=f"p_del_{idx}"):
                to_delete.append(idx)

    # suppression différée pour éviter conflits d'index
    if to_delete:
        for i in sorted(to_delete, reverse=True):
            del st.session_state.form_param_list[i]
        st.rerun()
    # ================== /PARAMS DYNAMIQUES ==================

    
    if st.button("💾 Enregistrer", type="primary", use_container_width=True):
        if not name.strip():
            st.error("Le nom de la méthode est requis")
        elif not outcol.strip():
            st.error("Le nom de la colonne de sortie est requis")
        else:
            try:
                current_order = None
                if st.session_state.editing_method_id:
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
                    required_columns=inferred,
                    code=code,
                    order=current_order,
                    method_id=st.session_state.editing_method_id
                )
                
                st.success(f"Méthode '{saved['name']}' enregistrée avec succès !")
                st.session_state.current_view = "list"
                st.rerun()
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement : {e}")

elif st.session_state.current_view == "test_all":
    # TEST TOUTES LES MÉTHODES
    col_title, col_back = st.columns([4, 1])
    with col_title:
        st.subheader("🧪 Test complet : toutes les méthodes")
    with col_back:
        if st.button("⬅️ Liste", use_container_width=True):
            st.session_state.current_view = "list"
            st.rerun()
    
    sample, s_err = _load_default_sample(gabarit)
    if s_err:
        st.error(s_err)
    elif sample is None or sample.empty:
        st.info("Aucun échantillon disponible")
    else:
        st.caption(f"📊 Échantillon : {sample.shape[0]} lignes × {sample.shape[1]} colonnes")
        
        with st.spinner("Application de toutes les méthodes..."):
            out_df, warns, errs = _apply_all_methods_on_sample(gabarit, sample)
        
        calc_cols = [m.get("output_column") for m in methods if m.get("output_column")]
        cols_to_show = [c for c in key_cols if c in out_df.columns] + [c for c in calc_cols if c in out_df.columns]
        
        if not warns and not errs:
            st.success("Toutes les méthodes ont été appliquées avec succès")
        if warns:
            for w in warns:
                st.warning(f"⚠️ {w}")
        if errs:
            for e in errs:
                st.error(f"❌ {e}")
        
        if cols_to_show:
            st.dataframe(out_df[cols_to_show], use_container_width=True, height=500)
        else:
            st.info("Aucune colonne calculée à afficher")

elif st.session_state.current_view == "test":
    # TEST UNE MÉTHODE
    m = st.session_state.testing_method
    if not m:
        st.info("Sélectionnez une méthode à tester")
    else:
        col_title, col_back = st.columns([4, 1])
        with col_title:
            st.subheader(f"🧪 Test : {m['name']} → `{m.get('output_column','?')}`")
        with col_back:
            if st.button("⬅️ Liste", use_container_width=True):
                st.session_state.current_view = "list"
                st.session_state.testing_method = None
                st.rerun()
        
        sample, err = _load_default_sample(gabarit)
        if err:
            st.error(err)
        elif sample is None or sample.empty:
            st.info("Aucun échantillon disponible")
        else:
            miss = [c for c in (m.get("required_columns") or []) if c not in sample.columns]
            if miss:
                st.error(f"Colonnes requises manquantes : {', '.join(miss)}")
            else:
                st.caption(f"📊 Échantillon : {sample.shape[0]} lignes × {sample.shape[1]} colonnes")
                
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
