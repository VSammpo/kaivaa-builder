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
    vals={}
    for spec in schema or []:
        nm, tp, dv = spec.get("name"), spec.get("type","text"), spec.get("default",None)
        if tp == "number":
            try: vals[nm] = float(dv) if dv not in (None,"") else 0.0
            except: vals[nm] = 0.0
        elif tp == "boolean":
            vals[nm] = str(dv).lower() in ("1","true","yes","y","on")
        elif tp == "select":
            opts = spec.get("options") or []
            vals[nm] = dv if dv in opts else (opts[0] if opts else "")
        elif tp == "select_from_column":
            col = spec.get("source_column")
            if col and col in df.columns:
                opts = sorted(list(df[col].dropna().astype(str).unique()))[:200]
                vals[nm] = str(dv) if str(dv) in opts else (opts[0] if opts else "")
            else:
                vals[nm] = str(dv) if dv is not None else ""
        else:
            vals[nm] = "" if dv is None else str(dv)
    return vals

def _param_schema_from_form() -> list[dict]:
    out=[]
    for p in st.session_state.form_param_list:
        name=(p.get("name") or "").strip()
        if not name:
            continue
        tp=(p.get("type") or "text").strip()
        item={"name":name,"type":tp,"default":p.get("default","")}
        lbl=(p.get("label") or "").strip()
        if lbl:
            item["label"]=lbl
        if tp=="select":
            opts_csv=(p.get("options") or "").strip()
            if opts_csv:
                item["options"]=[s.strip() for s in opts_csv.split(",") if s.strip()][:200]
        if tp=="select_from_column":
            col=(p.get("source_column") or "").strip()
            if col:
                item["source_column"]=col
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
                        st.session_state.form_param_list = [
                            {
                                "name": p.get("name",""), 
                                "label": p.get("label",""),
                                "type": p.get("type","text"), 
                                "default": p.get("default",""),
                                "options": ",".join(p.get("options",[])) if isinstance(p.get("options"),list) else (p.get("options","") or ""),
                                "source_column": p.get("source_column",""),
                            } for p in (m.get("param_schema") or [])
                        ]
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
                    param_schema=[],
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
                            out = apply_method(sample, m, {})
                        st.success("Méthode appliquée avec succès")
                        st.dataframe(out, use_container_width=True, height=500)
                    except MethodExecutionError as e:
                        st.error(f"Erreur d'exécution : {e}")
                    except Exception as e:
                        st.error(f"Erreur inattendue : {e}")