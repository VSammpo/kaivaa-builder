# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from collections import deque

# ==== Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.services.gabarit_registry import (
    list_gabarits, get_gabarit, get_relations, load_registry, list_methods_for_gabarit
)
# --- Helpers de reconciliation pour multiselects (colonnes / méthodes) ---
import difflib

from backend.services.transformation_service import (
    list_transformations,
    get_transformation,
    get_transformation_output_columns,
)


def _safe_reconcile_defaults(defaults: list[str] | None, options: list[str]) -> tuple[list[str], dict[str, str], list[str]]:
    """
    Retourne:
      - cleaned: liste des valeurs "default" garanties présentes dans options (après mapping).
      - renamed_map: dict {old_name -> new_name} pour les éléments vraisemblablement renommés.
      - removed: liste des valeurs définitivement retirées (ni exact match ni proche).
    Logique:
      1) on garde les exact match
      2) pour celles manquantes, on tente une correspondance "proche" (cutoff=0.86)
      3) sinon on les met en 'removed'
    """
    defaults = list(defaults or [])
    cleaned: list[str] = []
    renamed: dict[str, str] = {}
    removed: list[str] = []

    # normalisation simple pour éviter la casse/espaces parasites
    def norm(s: str) -> str:
        return (s or "").strip()

    opt_set = set(options)
    for d in defaults:
        d0 = norm(d)
        if d0 in opt_set:
            cleaned.append(d0)
        else:
            # tentative de matching "renommage"
            match = difflib.get_close_matches(d0, options, n=1, cutoff=0.86)
            if match:
                target = match[0]
                renamed[d0] = target
                if target not in cleaned:
                    cleaned.append(target)
            else:
                removed.append(d0)

    # garantir unicité dans l'ordre
    seen = set()
    cleaned = [x for x in cleaned if not (x in seen or seen.add(x))]
    return cleaned, renamed, removed



st.set_page_config(page_title="Injection des données", page_icon="📑", layout="wide")

# ==== NAVBAR (5 boutons) ====
def render_template_subnav(active: str, template_id: int | None):
    cols = st.columns([1,1,1,1,1])
    with cols[0]:
        if st.button("← Retour bibliothèque", use_container_width=True):
            if "selected_template" in st.session_state:
                del st.session_state.selected_template
            if "selected_template_detail" in st.session_state:
                del st.session_state.selected_template_detail
            st.switch_page("pages/3_📚_Bibliotheque.py")
    with cols[1]:
        if st.button("🗂️ Détail du template", type=("primary" if active=="detail" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template_detail = template_id
            st.switch_page("pages/_3a_📊_Detail_Livrable.py")
    with cols[2]:
        if st.button("⚙️ Paramètres généraux", type=("primary" if active=="general" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_3b_➕_Form_Template.py")
    with cols[3]:
        st.button("📑 Injection des données", type="primary" if active=="inject" else "secondary", use_container_width=True)
    with cols[4]:
        if st.button("🧾 Ajustement de la table", type=("primary" if active=="adjust" else "secondary"), use_container_width=True):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_3d_🧾_Ajustement_Table.py")
    st.divider()

# ==== Guard sélection template ====
if 'selected_template' not in st.session_state and 'selected_template_detail' in st.session_state:
    st.session_state.selected_template = st.session_state.selected_template_detail

if 'selected_template' not in st.session_state or not st.session_state.selected_template:
    st.error("Aucun template sélectionné.")
    if st.button("← Retour bibliothèque"):
        st.switch_page("pages/3_📚_Bibliotheque.py")
    st.stop()

template_id = st.session_state.selected_template

# ==== Charger PRIMITIFS (évite Detached) ====
with DatabaseService.get_session() as db:
    ts = TemplateService(db)
    tpl = ts.get_template(template_id)
    cfg = ts.get_config(template_id)  # JSON dict
    usages = (cfg.get("gabarit_usages") or []) if isinstance(cfg, dict) else []
    tpl_name = tpl.name
    tpl_version = tpl.version

# ==== Navbar + Titre ====
render_template_subnav("inject", template_id)
st.title(f"📑 Injection des données — {tpl_name} (v{tpl_version})")
# Indicateur visuel du mode + switch création/édition
_edit_key = st.session_state.get("_inject_edit_target")
if isinstance(_edit_key, dict):
    _lbl = f"{_edit_key.get('gname','?')} (v{_edit_key.get('gver','v1')}) → {(_edit_key.get('sheet') or 'Data')}/{(_edit_key.get('table') or 'Table')}"
    cL, cR = st.columns([3,1])
    with cL:
        st.info(f"✏️ **Mode ÉDITION** de l’usage : {_lbl}")
    with cR:
        if st.button("➕ Créer un nouvel usage", use_container_width=True, key="switch_to_create"):
            st.session_state._inject_edit_target = None
            # on reset l'état d'édition d'enrichissements éventuel
            if "tpl_enrich_rows" in st.session_state:
                del st.session_state["tpl_enrich_rows"]
            if "_inject_loaded_for" in st.session_state:
                del st.session_state["_inject_loaded_for"]
            st.rerun()
else:
    st.success("➕ **Mode CRÉATION** d’un nouvel usage")


if st.session_state.get("_inject_saved"):
    st.success("✅ Usage enregistré")
    del st.session_state["_inject_saved"]

# -----------------------------------------------------------------------------------
# Liste des usages existants (lecture + suppression + édit)
# -----------------------------------------------------------------------------------
st.subheader("Usages existants")
if not usages:
    st.info("Aucun usage configuré pour ce template.")
else:
    for u in usages:
        # clé en cours d'édition ?
        _edit_key = st.session_state.get("_inject_edit_target")
        tgt = u.get("excel_target") or {}
        sheet_u, table_u = (tgt.get("sheet") or ""), (tgt.get("table") or "")

        is_editing = (
            isinstance(_edit_key, dict)
            and _edit_key.get("gname") == u.get("gabarit_name")
            and (_edit_key.get("gver") or "v1") == (u.get("gabarit_version") or "v1")
            and (_edit_key.get("sheet") or "") == sheet_u
            and (_edit_key.get("table") or "") == table_u
        )

        # Carte (un seul container) — si en édition, on ajoute un fin liseré coloré en haut
        box = st.container()
        st.markdown("<div style='border:1px solid #eee;border-radius:8px;padding:10px;margin-bottom:6px;'>", unsafe_allow_html=True)

        with box:
            if is_editing:
                st.markdown("<div style='height:0;border-top:3px solid #3b82f6;margin:-6px 0 10px 0;'></div>", unsafe_allow_html=True)

            c1, c2, c3, c4, c5 = st.columns([3,3,2,2,1])
            with c1:
                # Titre de carte = NOM DE L’USAGE (feuille/table)
                st.markdown(f"**Usage : {sheet_u}/{table_u}**")
                st.caption(f"Gabarit : {u.get('gabarit_name')} (v{u.get('gabarit_version','v1')})")
            with c2:
                st.caption(f"Colonnes gardées : {len(u.get('columns_enabled') or [])}")
                st.caption(f"Méthodes : {', '.join(u.get('methods') or []) or '—'}")
                st.caption(f"Enrichissements : {len(u.get('enrichments') or [])}")
            with c3:
                has_overlay = bool((u.get('overlay_python') or '').strip())
                st.caption(f"Overlay : {'Oui' if has_overlay else '—'}")
            with c4:
                if st.button(
                    "✏️ Éditer",
                    key=f"edit_usage_{u.get('gabarit_name')}|{u.get('gabarit_version','v1')}|{sheet_u}|{table_u}",
                    use_container_width=True
                ):

                    st.session_state._inject_edit_target = {
                        "gname": u.get("gabarit_name"),
                        "gver": u.get("gabarit_version","v1"),
                        "sheet": sheet_u,
                        "table": table_u,
                    }
                    st.rerun()
            with c5:
                del_key = (u.get("gabarit_name"), u.get("gabarit_version","v1"), sheet_u, table_u)
                if st.button("🗑️",
                             key=f"del_usage_{del_key[0]}|{del_key[1]}|{del_key[2]}|{del_key[3]}",

                             use_container_width=True,
                             help="Supprimer cet usage (clé = gabarit+feuille+table)"):
                    try:
                        with DatabaseService.get_session() as db2:
                            ts2 = TemplateService(db2)
                            cfg2 = ts2.get_config(template_id)
                            allu = cfg2.get("gabarit_usages", []) or []
                            gname, gver, sheet_, table_ = del_key
                            allu = [
                                x for x in allu
                                if not (
                                    x.get("gabarit_name")==gname
                                    and (x.get("gabarit_version") or "v1")==gver
                                    and ((x.get("excel_target") or {}).get("sheet","") or "")==sheet_
                                    and ((x.get("excel_target") or {}).get("table","") or "")==table_
                                )
                            ]
                            cfg2["gabarit_usages"] = allu
                            ts2.update_config(template_id, cfg2)
                        # si on supprime celui en édition, on repasse en création
                        if is_editing:
                            st.session_state._inject_edit_target = None
                        st.success("Usage supprimé")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Suppression impossible : {e}")
        st.markdown("</div>", unsafe_allow_html=True)


st.markdown("---")

# -----------------------------------------------------------------------------------
# FORMULAIRE AJOUTER / ÉDITER UN USAGE (seule colonne)
# -----------------------------------------------------------------------------------
st.subheader("➕ Ajouter / ✏️ Éditer un usage")

# Helpers
def compute_reachable_targets(start_name: str, start_version: str, max_depth: int = 4):
    """BFS sur les relations pour lister toutes les tables atteignables. Retourne {(name,ver): path}."""
    reached = {}
    visited = set([(start_name, start_version)])
    q = deque([((start_name, start_version), [])])
    depth = {(start_name, start_version): 0}
    while q:
        (cur_name, cur_ver), path = q.popleft()
        if depth[(cur_name, cur_ver)] >= max_depth:
            continue
        rels = get_relations(cur_name, cur_ver) or []
        for r in rels:
            to_name = r.get("to_gabarit")
            to_ver = r.get("to_version", "v1")
            if not to_name:
                continue
            step = [cur_name, r.get("left_key"), to_name, r.get("right_key")]
            new_path = path + [step]
            if (to_name, to_ver) not in visited:
                visited.add((to_name, to_ver))
                depth[(to_name, to_ver)] = depth[(cur_name, cur_ver)] + 1
                q.append(((to_name, to_ver), new_path))
            reached.setdefault((to_name, to_ver), new_path)
    reached.pop((start_name, start_version), None)
    return reached

def _list_methods_names_from_registry(gabarit_name: str, gabarit_version: str) -> list[str]:
    try:
        reg = load_registry()
        meta = (reg.get(gabarit_name) or {}).get("versions", {}).get(gabarit_version) \
               or (reg.get(gabarit_name) or {}).get("versions", {}).get("v1") or {}
        methods = meta.get("methods") or {}
        if isinstance(methods, dict):
            return sorted(list(methods.keys()))
        elif isinstance(methods, list):
            out = []
            for m in methods:
                if isinstance(m, dict) and m.get("name"):
                    out.append(str(m["name"]))
            return sorted(out)
        return []
    except Exception:
        return []

def _build_enrichments_payload(start_g, rows: list[dict]) -> list[dict]:
    payload = []
    path_map = compute_reachable_targets(start_g.name, start_g.version, max_depth=6)
    for r in rows:
        if not r.get("target"):
            continue
        tgt = tuple(r["target"])
        p = path_map.get(tgt)
        if not p:
            continue
        payload.append({
            "join": r.get("join","left"),
            "path": p,
            "columns": r.get("columns", [])
        })
    return payload

# Sélection gabarit
gab_list = list_gabarits()
if not gab_list:
    st.error("Aucun gabarit disponible dans le registre.")
    st.stop()

# Mode EDIT ciblé ?
edit_key = st.session_state.get("_inject_edit_target")

# ========== ÉTAPE 1 : CHOIX DU MODE (Gabarit vs Transformation) ==========
st.markdown("### 🧩 Source de la table")

# Détecter le mode existant (si on édite)
_existing_source_kind = "gabarit"
_existing_tname = ""
_existing_tver = "v1"

if edit_key:
    # Charger l'usage existant pour détecter le mode
    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        temp_existing = ts.get_gabarit_usage_by_target(
            template_id,
            edit_key.get("gname", gab_list[0].name),
            edit_key.get("gver", gab_list[0].version),
            edit_key.get("sheet", "Data"),
            edit_key.get("table", gab_list[0].name)
        ) or {}
        _existing_source_kind = (temp_existing.get("source_kind") or "gabarit").strip()
        _existing_tname = (temp_existing.get("transformation_name") or "").strip()
        _existing_tver = (temp_existing.get("transformation_version") or "v1").strip()

_source_options = ["Gabarit brut", "Transformation réutilisable"]
_source_index = 1 if (_existing_source_kind == "transformation" or bool(_existing_tname)) else 0

source_mode = st.radio(
    "Type de source",
    options=_source_options,
    index=_source_index,
    horizontal=True,
    key=f"src_mode_{template_id}_global",
)

_is_transfo_mode = (source_mode == "Transformation réutilisable")

# ========== ÉTAPE 2a : MODE TRANSFORMATION ==========
if _is_transfo_mode:
    all_tf = list_transformations() or []
    
    def _tf_label(t):
        gb = (t.get("gabarit_base") or {})
        return f"{t.get('name')} (v{t.get('version','v1')}) – base: {gb.get('name','?')}"
    
    tf_labels = [_tf_label(t) for t in all_tf]
    
    default_t_idx = 0
    if _existing_tname:
        for i, t in enumerate(all_tf):
            if (t.get("name") == _existing_tname) and (t.get("version","v1") == _existing_tver):
                default_t_idx = i
                break
    
    if not all_tf:
        st.warning("Aucune transformation disponible.")
        st.stop()
    
    sel = st.selectbox(
        "Transformation",
        tf_labels,
        index=default_t_idx,
        key=f"transfo_sel_{template_id}_global"
    )
    selected_transfo = all_tf[tf_labels.index(sel)]
    
    # 🔑 Récupérer le gabarit de BASE de la transformation
    transfo_gab_base = selected_transfo.get("gabarit_base", {})
    g_name = transfo_gab_base.get("name", "")
    g_version = transfo_gab_base.get("version", "v1")
    
    g = get_gabarit(g_name, g_version)
    if not g:
        st.error(f"Gabarit de base '{g_name}' de la transformation introuvable")
        st.stop()
    
    st.info(f"📦 Gabarit de base : **{g.name}** (v{g.version})")
    
    # Excel targets
    default_sheet = edit_key.get("sheet", "Data") if edit_key else "Data"
    default_table = edit_key.get("table", g.name) if edit_key else g.name
    
    # Charger usage existant
    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        existing = ts.get_gabarit_usage_by_target(template_id, g.name, g.version, default_sheet, default_table) or {}

# ========== ÉTAPE 2b : MODE GABARIT BRUT ==========
else:
    selected_transfo = None
    
    # Sélection du gabarit
    labels = [f"{gg.name} (v{gg.version})" for gg in gab_list]
    
    idx = 0
    if edit_key:
        try:
            idx = next(i for i, gg in enumerate(gab_list) if (gg.name, gg.version) == (edit_key["gname"], edit_key["gver"]))
        except StopIteration:
            idx = 0
    
    gab_choice = st.selectbox("Gabarit de départ", labels, index=idx if labels else 0)
    g = gab_list[labels.index(gab_choice)]
    
    # Excel targets
    if edit_key and (edit_key["gname"], edit_key["gver"]) == (g.name, g.version):
        default_sheet = edit_key.get("sheet", "Data")
        default_table = edit_key.get("table", g.name)
    else:
        default_sheet = "Data"
        default_table = g.name
    
    # Charger usage existant
    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        existing = ts.get_gabarit_usage_by_target(template_id, g.name, g.version, default_sheet, default_table) or {}

    
    # Préparer les colonnes de base pour le multiselect
    base_cols = [c.name for c in g.columns]
    existing_cols = existing.get("columns_enabled", []) if existing else base_cols[:]
    _clean_cols, _renamed_cols, _removed_cols = _safe_reconcile_defaults(existing_cols, base_cols)
    
    if _renamed_cols:
        st.caption("🪄 Renommages de colonnes appliqués : " + ", ".join([f"{k} → {v}" for k, v in _renamed_cols.items()]))
    if _removed_cols:
        st.caption("⚠️ Colonnes introuvables (retirées) : " + ", ".join(_removed_cols))

# ========== FIN DE LA SÉLECTION ==========
# À ce stade, on a toujours:
# - g: le gabarit (soit sélectionné, soit depuis la transformation)
# - selected_transfo: None ou la transformation
# - existing: l'usage existant
# - _is_transfo_mode: True/False

# === Détection du mode (gabarit brut vs transformation) — VERSION NETTOYÉE ===# === Détection du mode (gabarit brut vs transformation) — VERSION NETTOYÉE ===

# === UI Colonnes (un seul chemin, sans doublon) ===
if _is_transfo_mode and selected_transfo:
    out_cols = get_transformation_output_columns(
        selected_transfo.get("name",""),
        selected_transfo.get("version","v1")
    ) or []

    default_cols = (existing.get("final_order") or []) if existing else out_cols[:]
    default_cols, ren_map, removed = _safe_reconcile_defaults(default_cols, out_cols)

    if ren_map:
        st.caption("🪄 Renommages (transformation) : " + ", ".join([f"{k} → {v}" for k, v in ren_map.items()]))
    if removed:
        st.caption("⚠️ Colonnes introuvables (retirées) : " + ", ".join(removed))

    enabled = st.multiselect(
        "Colonnes de sortie (transformation) à conserver (l'ordre de sélection sera repris)",
        options=out_cols,
        default=default_cols,
        key=f"ms_cols_transfo_{template_id}_{g.name}_{g.version}_{default_sheet}_{default_table}",
    )
else:
    enabled = st.multiselect(
        "Colonnes à conserver (si vide → toutes les colonnes du gabarit)",
        options=base_cols,
        default=_clean_cols,
        key=f"ms_cols_{g.name}_{g.version}",
    )


# Enrichissements (masqués entièrement si transformation)
if _is_transfo_mode:
    st.markdown("### 🔗 Enrichissements")
    st.caption("Pilotés par la transformation sélectionnée — non éditables ici.")
else:
    st.markdown("### 🔗 Enrichissements")

    reachable = compute_reachable_targets(g.name, g.version, max_depth=4)
    target_labels = sorted([f"{nm} (v{ver})" for (nm,ver) in reachable.keys()])
    label_to_tuple = { f"{nm} (v{ver})": (nm,ver) for (nm,ver) in reachable.keys() }

    _loaded_key = (template_id, g.name, g.version, (default_sheet or "").strip(), (default_table or "").strip())

    if ("tpl_enrich_rows" not in st.session_state) or (st.session_state.get("_inject_loaded_for") != _loaded_key):
        rows = []
        if existing and existing.get("enrichments"):
            for e in existing["enrichments"]:
                path = e.get("path") or []
                if path:
                    last = path[-1]  # [from, left_key, to, right_key]
                    rows.append({
                        "join": e.get("join", "left"),
                        "target": (last[2], "v1"),
                        "columns": e.get("columns", [])
                    })
                else:
                    rows.append({
                        "join": e.get("join", "left"),
                        "target": None,
                        "columns": e.get("columns", [])
                    })
        st.session_state.tpl_enrich_rows = rows
        st.session_state._inject_loaded_for = _loaded_key

    if st.button("➕ Ajouter un enrichissement", use_container_width=True):
        st.session_state.tpl_enrich_rows.append({"join":"left", "target": None, "columns": []})
        st.rerun()

    to_delete = []
    for i, row in enumerate(st.session_state.tpl_enrich_rows):
        with st.expander(f"Enrichissement #{i+1}", expanded=True):
            c0, c1 = st.columns([1,3])
            with c0:
                row["join"] = st.selectbox(
                    "Type de jointure", ["left","inner"],
                    index=(0 if row.get("join","left")=="left" else 1),
                    key=f"join_{i}"
                )
            with c1:
                st.caption("Table cible (atteignable via les relations définies dans le gabarit)")
                cur_label = None
                if row.get("target"):
                    nm, ver = row["target"]
                    cur_label = f"{nm} (v{ver})" if (nm,ver) in reachable else None
                sel = st.selectbox(
                    "Table à enrichir",
                    options=["(choisir)"] + target_labels,
                    index=(target_labels.index(cur_label)+1 if cur_label in target_labels else 0),
                    key=f"target_{i}"
                )
                if sel != "(choisir)":
                    row["target"] = label_to_tuple[sel]
                    tgt_g = get_gabarit(*row["target"])
                    tgt_cols = [c.name for c in (tgt_g.columns or [])]

                    existing_e = row.get("columns", []) or []
                    _clean_e, _renamed_e, _removed_e = _safe_reconcile_defaults(existing_e, tgt_cols)

                    if _renamed_e:
                        st.caption("🪄 Renommages (enrichissement) : " + ", ".join([f"{k} → {v}" for k, v in _renamed_e.items()]))
                    if _removed_e:
                        st.caption("⚠️ Colonnes introuvables (enrichissement) : " + ", ".join(_removed_e))

                    row["columns"] = st.multiselect(
                        "Colonnes à rapatrier",
                        options=tgt_cols,
                        default=_clean_e,
                        key=f"cols_{i}"
                    )
                else:
                    row["target"] = None
                    row["columns"] = []

            if st.button("🗑️ Supprimer", key=f"del_enrich_{i}"):
                to_delete.append(i)

    if to_delete:
        for i in sorted(to_delete, reverse=True):
            del st.session_state.tpl_enrich_rows[i]
        st.rerun()

# -----------------------------------------------------------------------------------
# Étape 3 : ⚙️ Méthodes (colonnes calculées) à inclure
# -----------------------------------------------------------------------------------
if _is_transfo_mode:
    st.markdown("### ⚙️ Méthodes à inclure")
    st.caption("Pilotées par la transformation sélectionnée — non éditables ici.")
    methods_selected = []
else:
    st.markdown("### ⚙️ Méthodes à inclure")

    def _method_names(gname: str, gver: str) -> list[str]:
        allm = list_methods_for_gabarit(gname, gver) or []
        if isinstance(allm, dict):
            return sorted(list(allm.keys()))
        names = []
        for m in allm:
            if isinstance(m, dict) and m.get("name"):
                names.append(m["name"])
        return sorted(names)

    all_methods = _method_names(g.name, g.version)
    # 🔑 CORRECTION : Charger les méthodes du gabarit de base + des gabarits enrichis
    all_methods = _method_names(g.name, g.version)
    all_methods_with_source = {f"{g.name}: {m}": m for m in all_methods}  # Format: "SELL-IN: ma_methode"
    
    # Ajouter les méthodes des gabarits enrichis
    if st.session_state.get("tpl_enrich_rows"):
        for enrich_row in st.session_state.get("tpl_enrich_rows", []):
            target = enrich_row.get("target")
            if target:
                target_name, target_ver = target
                target_methods = _method_names(target_name, target_ver)
                for tm in target_methods:
                    all_methods_with_source[f"{target_name}: {tm}"] = tm

    existing_methods = existing.get("methods", []) if existing else []
    
    # Mapper les anciennes méthodes (format court) vers le nouveau format (avec source)
    existing_methods_mapped = []
    for em in existing_methods:
        # Chercher dans all_methods_with_source
        found = False
        for display_name, short_name in all_methods_with_source.items():
            if short_name == em:
                existing_methods_mapped.append(display_name)
                found = True
                break
        if not found:
            # Méthode introuvable, on garde quand même
            existing_methods_mapped.append(em)
    
    options_display = sorted(list(all_methods_with_source.keys()))
    _clean_m, _renamed_m, _removed_m = _safe_reconcile_defaults(existing_methods_mapped, options_display)

    if _renamed_m:
        st.caption("🪄 Renommages de méthodes appliqués : " + ", ".join([f"{k} → {v}" for k, v in _renamed_m.items()]))
    if _removed_m:
        st.caption("⚠️ Méthodes introuvables (retirées) : " + ", ".join(_removed_m))

    methods_selected_display = st.multiselect(
        "Méthodes disponibles (gabarit de base + enrichissements)",
        options=options_display,
        default=_clean_m,
        key=f"ms_methods_{g.name}_{g.version}_{default_sheet}_{default_table}",
        help="Ces colonnes calculées seront disponibles. Les méthodes des tables enrichies sont préfixées par le nom du gabarit."
    )
    
    # Convertir les sélections affichées (avec source) en format court pour la sauvegarde
    # Format sélectionné: "Dim_date_FMCG: Numéro_semaine_table"
    # Format à sauvegarder: "Numéro_semaine_table"
    methods_selected = [all_methods_with_source.get(m, m) for m in methods_selected_display]




# Cible Excel
st.markdown("### 🎯 Cible Excel")
cS, cT = st.columns(2)
with cS:
    sheet = st.text_input("Feuille Excel", value=(existing.get("excel_target",{}).get("sheet", default_sheet)))
with cT:
    table = st.text_input("Table Excel", value=(existing.get("excel_target",{}).get("table", default_table)))


st.markdown("---")

# ENREGISTRER (en préservant l'ajustement final existant)
if st.button("💾 Enregistrer l’usage", type="primary", use_container_width=True):
    try:
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            old_for_key = ts.get_gabarit_usage_by_target(
                template_id, g.name, g.version, sheet.strip(), table.strip()
            ) or {}

        keep_final_order = old_for_key.get("final_order")
        keep_final_excl  = old_for_key.get("final_excludes")
        keep_final_ren   = old_for_key.get("final_renames")
        keep_final_sort  = old_for_key.get("final_sort")

        with DatabaseService.get_session() as db:
            ts = TemplateService(db)

            if source_mode == "Transformation réutilisable" and selected_transfo:
                ts.upsert_gabarit_usage(
                    template_id=template_id,
                    gabarit_name=g.name,
                    gabarit_version=g.version,
                    excel_sheet=sheet.strip(),
                    excel_table=table.strip(),

                    # côté template on neutralise ces champs (la transfo porte le pipeline)
                    columns_enabled=[],
                    methods=[],
                    enrichments=[],

                    # la sélection 'enabled' pilote le sous-ensemble/ordre final
                    final_order=list(enabled or []),
                    final_excludes=list(keep_final_excl or []),

                    # source + référence transformation
                    source_kind="transformation",
                    transformation_name=selected_transfo.get("name",""),
                    transformation_version=selected_transfo.get("version","v1"),
                )

                # préserver final_renames / final_sort si absents après upsert
                cfg2 = ts.get_config(template_id) or {}
                usages2 = list(cfg2.get("gabarit_usages") or [])
                for uu in usages2:
                    tgt2 = uu.get("excel_target") or {}
                    if (
                        uu.get("gabarit_name") == g.name
                        and (uu.get("gabarit_version") or "v1") == g.version
                        and (tgt2.get("sheet") or "") == sheet.strip()
                        and (tgt2.get("table") or "") == table.strip()
                    ):
                        if keep_final_ren is not None and "final_renames" not in uu:
                            uu["final_renames"] = keep_final_ren
                        if keep_final_sort is not None and "final_sort" not in uu:
                            uu["final_sort"] = keep_final_sort
                        break
                cfg2["gabarit_usages"] = usages2
                ts.update_config(template_id, cfg2)

            else:
                # Mode gabarit brut (comportement actuel)
                enrich_payload = _build_enrichments_payload(g, st.session_state.get("tpl_enrich_rows", []))
                # ✅ CORRECTION : Utiliser la variable methods_selected déjà convertie (ligne 592)
                # (pas besoin de relire depuis session_state car déjà calculée)
                ts.upsert_gabarit_usage(
                    template_id=template_id,
                    gabarit_name=g.name,
                    gabarit_version=g.version,
                    excel_sheet=sheet.strip(),
                    excel_table=table.strip(),
                    columns_enabled=enabled,
                    methods=methods_selected,
                    enrichments=enrich_payload,
                    final_order=keep_final_order,
                    final_excludes=keep_final_excl,
                    # source_kind par défaut = "gabarit" (côté service)
                )

        st.session_state["_inject_saved"] = True
        st.rerun()
    except Exception as e:
        st.error(f"Erreur : {e}")