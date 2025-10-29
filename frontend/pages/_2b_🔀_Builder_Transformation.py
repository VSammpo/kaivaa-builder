"""
Page de construction/édition des transformations
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from collections import deque
from code_editor import code_editor

# Bootstrap
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.transformation_service import (
    get_transformation,
    update_transformation
)
from backend.services.gabarit_registry import (
    get_gabarit,
    list_methods_for_gabarit,
    get_relations,
    get_default_preview
)
from backend.services.table_builder_service import build_table_from_usage
from backend.services.parameter_service import ParameterService

st.set_page_config(
    page_title="Builder Transformation",
    page_icon="🔀",
    layout="wide"
)

# ========= HELPERS ==========
def compute_reachable_targets(start_name: str, start_version: str, max_depth: int = 4):
    """BFS sur les relations pour lister toutes les tables atteignables."""
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

def _fmt_version_label(ver: str) -> str:
    """Retourne 'v1' si on reçoit '1', et laisse tel quel si on reçoit déjà 'v1'."""
    v = str(ver or "").strip()
    return v if v.lower().startswith("v") else f"v{v}"


def _build_enrichments_payload(start_name, start_version, rows: list[dict]) -> list[dict]:
    """Construit le payload des enrichissements."""
    payload = []
    path_map = compute_reachable_targets(start_name, start_version, max_depth=6)
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

def _sample_value(col: str, gname: str, gver: str) -> str:
    """Exemple rapide depuis le preview de base."""
    try:
        prev = get_default_preview(gname, gver) or {}
        rows, cols = prev.get("rows") or [], prev.get("columns") or []
        if rows and cols:
            dfp = pd.DataFrame(rows, columns=cols)
            if col in dfp.columns and not dfp.empty:
                return str(dfp[col].iloc[0])[:60]
    except Exception:
        pass
    return "—"

def _muted_html(s: str) -> str:
    return f"<span style='color:#9aa0a6'>{s}</span>"

# ========= VÉRIFICATION TRANSFORMATION ==========
if "selected_transformation" not in st.session_state:
    st.error("Aucune transformation sélectionnée")
    if st.button("Retour à la liste"):
        st.switch_page("pages/_3d_🔀_Transformations.py")
    st.stop()

name, version = st.session_state.selected_transformation
transformation = get_transformation(name, version)

if not transformation:
    st.error(f"Transformation '{name}' v{version} introuvable")
    st.stop()

# Gabarit de base
gabarit_base = transformation.get("gabarit_base", {})
gab_name = gabarit_base.get("name", "")
gab_version = gabarit_base.get("version", "v1")

gabarit = get_gabarit(gab_name, gab_version)
if not gabarit:
    st.error(f"Gabarit '{gab_name}' {gab_version} introuvable")
    st.stop()

# ========= TITRE ET NAV ==========
st.title(f"🔀 Configuration : {name}")
st.caption(f"Version : {version}")

# ========= INITIALISATION SESSION ==========
if "transfo_enrich_rows" not in st.session_state:
    # Charger depuis la transformation existante
    rows = []
    for e in transformation.get("enrichments", []):
        path = e.get("path") or []
        if path:
            last = path[-1]
            rows.append({
                "join": e.get("join", "left"),
                "target": (last[2], "v1"),
                "columns": e.get("columns", [])
            })
    st.session_state.transfo_enrich_rows = rows

# Code Python buffer
code_key = f"transfo_code_{name}_{version}"
if code_key not in st.session_state:
    st.session_state[code_key] = transformation.get("overlay_python", "")

# État ajustement
if "transfo_adjust_state" not in st.session_state:
    st.session_state.transfo_adjust_state = {
        "final_order": transformation.get("final_order", []),
        "final_excludes": set(transformation.get("final_excludes", [])),
        "final_renames": transformation.get("final_renames", {})
    }

# ========= TABS PRINCIPAUX ==========
tab_cols, tab_enrich, tab_methods, tab_script, tab_adjust, tab_preview = st.tabs([
    "📊 Colonnes",
    "🔗 Enrichissements", 
    "⚙️ Méthodes",
    "🐍 Script Python",
    "🧾 Ajustement",
    "👀 Prévisualisation"
])

# ==================== TAB COLONNES ====================
with tab_cols:
    st.subheader("📊 Colonnes de base")
    st.info(f"Gabarit source : **{gab_name}** ({gab_version})")
    
    all_columns = [c.name for c in gabarit.columns]
    # Par défaut, toutes les colonnes sont sélectionnées
    selected_cols = transformation.get("columns_enabled") or all_columns[:]
    
    enabled = st.multiselect(
        "Colonnes à conserver",
        options=all_columns,
        default=selected_cols,
        help="Par défaut, toutes les colonnes sont conservées"
    )
    
    # Afficher les colonnes sélectionnées
    if enabled:
        st.success(f"✅ {len(enabled)} colonnes sélectionnées")
        
        # Affichage en grille
        cols_grid = st.columns(min(len(enabled), 4))
        for idx, col_name in enumerate(enabled[:12]):
            with cols_grid[idx % len(cols_grid)]:
                st.caption(f"• {col_name}")
        if len(enabled) > 12:
            st.caption(f"... et {len(enabled) - 12} autres")
    else:
        st.warning("⚠️ Aucune colonne sélectionnée")

# ==================== TAB ENRICHISSEMENTS ====================
with tab_enrich:
    st.subheader("🔗 Enrichissements")
    st.caption("Enrichissez vos données avec des tables de référence via des jointures")
    
    # Calcul des cibles atteignables
    reachable = compute_reachable_targets(gab_name, gab_version, max_depth=4)
    target_labels = sorted([f"{nm} ({_fmt_version_label(ver)})" for (nm,ver) in reachable.keys()])
    label_to_tuple = { f"{nm} ({_fmt_version_label(ver)})": (nm,ver) for (nm,ver) in reachable.keys() }

    
    # Bouton ajouter
    if st.button("➕ Ajouter un enrichissement", use_container_width=True, type="primary"):
        st.session_state.transfo_enrich_rows.append({
            "join": "left", 
            "target": None, 
            "columns": []
        })
        st.rerun()
    
    # Affichage des enrichissements
    to_delete = []
    for i, row in enumerate(st.session_state.transfo_enrich_rows):
        with st.expander(f"Enrichissement #{i+1}", expanded=True):
            col1, col2 = st.columns([1, 3])
            
            with col1:
                row["join"] = st.selectbox(
                    "Type de jointure",
                    ["left", "inner"],
                    index=0 if row.get("join","left")=="left" else 1,
                    key=f"transfo_join_{i}"
                )
            
            with col2:
                st.caption("Table cible (atteignable via les relations)")
                
                # Sélection de la cible
                cur_label = None
                if row.get("target"):
                    nm, ver = row["target"]
                    cur_label = f"{nm} ({_fmt_version_label(ver)})" if (nm,ver) in reachable else None
                
                sel = st.selectbox(
                    "Table à enrichir",
                    options=["(choisir)"] + target_labels,
                    index=(target_labels.index(cur_label)+1 if cur_label in target_labels else 0),
                    key=f"transfo_target_{i}"
                )
                
                if sel != "(choisir)":
                    row["target"] = label_to_tuple[sel]
                    tgt_g = get_gabarit(*row["target"])
                    if tgt_g:
                        tgt_cols = [c.name for c in (tgt_g.columns or [])]
                        
                        row["columns"] = st.multiselect(
                            "Colonnes à rapatrier",
                            options=tgt_cols,
                            default=row.get("columns", []),
                            key=f"transfo_cols_{i}"
                        )
                else:
                    row["target"] = None
                    row["columns"] = []
            
            if st.button("🗑️ Supprimer", key=f"del_transfo_enrich_{i}"):
                to_delete.append(i)
    
    # Suppression
    if to_delete:
        for i in sorted(to_delete, reverse=True):
            del st.session_state.transfo_enrich_rows[i]
        st.rerun()

# ==================== TAB MÉTHODES ====================
with tab_methods:
    st.subheader("⚙️ Méthodes (colonnes calculées)")
    
    available_methods = list_methods_for_gabarit(gab_name, gab_version) or []
    
    if available_methods:
        method_names = [m.get("name", "") for m in available_methods]
        selected_methods = transformation.get("methods", [])
        
        methods_selected = st.multiselect(
            "Méthodes à appliquer",
            options=method_names,
            default=selected_methods,
            help="Les méthodes seront appliquées dans l'ordre"
        )
        
        # Afficher détails des méthodes
        if methods_selected:
            st.success(f"✅ {len(methods_selected)} méthode(s) sélectionnée(s)")
            
            for method_name in methods_selected:
                method = next((m for m in available_methods if m.get("name") == method_name), None)
                if method:
                    with st.expander(f"⚙️ {method_name}"):
                        st.markdown(f"**Description :** {method.get('description', '')}")
                        st.markdown(f"**Colonne créée :** `{method.get('output_column', '')}`")
                        req_cols = method.get('required_columns', [])
                        if req_cols:
                            st.markdown(f"**Colonnes requises :** {', '.join(f'`{c}`' for c in req_cols)}")
    else:
        st.info("Aucune méthode définie pour ce gabarit")
        methods_selected = []

# ==================== TAB SCRIPT PYTHON ====================
with tab_script:
    st.subheader("🐍 Script Python de transformation")
    st.caption("💡 Variables disponibles : `df` (DataFrame), `pd` (pandas)")
    st.caption("⚠️ Le script doit réassigner `df` (ex : `df = df[df['Col'] == 'Value']`)")
    
    # Code editor avec persistance
    custom_buttons = [{
        "name": "Copier", "feather": "Copy", "hasText": True,
        "commands": ["copyAll"], "style": {"top": "0.46rem", "right": "0.4rem"}
    }]
    
    with st.form(f"transfo_script_form_{name}_{version}", clear_on_submit=False, border=True):
        editor_result = code_editor(
            st.session_state[code_key],
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
                "enableBasicAutocompletion": True
            },
            key=f"transfo_code_editor_{name}_{version}",
            response_mode=["submit", "blur"]
        )
        
        # Extraction du code
        if editor_result:
            _new = None
            if isinstance(editor_result, dict):
                _new = editor_result.get("text") or editor_result.get("content") or editor_result.get("code")
            elif isinstance(editor_result, str):
                _new = editor_result
            if isinstance(_new, str):
                st.session_state[code_key] = _new
        
        st.caption(f"📄 {len(st.session_state[code_key])} caractères")
        st.markdown("---")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            do_preview = st.form_submit_button("🧪 Prévisualiser", use_container_width=True)
        with col2:
            do_test = st.form_submit_button("✅ Valider", use_container_width=True)
        with col3:
            do_save_script = st.form_submit_button("💾 Enregistrer script", type="primary", use_container_width=True)
    
    # Actions après le form
    if do_preview or do_test:
        # Construire la config temporaire
        temp_config = {
            "gabarit_base": {"name": gab_name, "version": gab_version},
            "columns_enabled": enabled if 'enabled' in locals() else transformation.get("columns_enabled", []),
            "enrichments": _build_enrichments_payload(gab_name, gab_version, st.session_state.transfo_enrich_rows),
            "methods": methods_selected if 'methods_selected' in locals() else [],
            "overlay_python": st.session_state[code_key]
        }
        
        with st.spinner("Exécution..."):
            df, error = build_table_from_usage(temp_config, full=True, log_kpis=True)

        
        if error:
            st.error(f"❌ Erreur :\n```\n{error}\n```")
        elif df is None or df.empty:
            st.warning("Aucun résultat")
        else:
            st.success(f"✅ {len(df)} lignes × {len(df.columns)} colonnes")
            st.dataframe(df.head(20), use_container_width=True, hide_index=True)
    
    if do_save_script:
        st.success("✅ Script enregistré")
        st.toast("Script sauvegardé en mémoire")

# ==================== TAB AJUSTEMENT ====================
with tab_adjust:
    st.subheader("🧾 Ajustement de la table finale")
    st.caption("Réorganisez, renommez ou excluez des colonnes")
    
    # Calculer les colonnes disponibles
    temp_config = {
        "gabarit_base": {"name": gab_name, "version": gab_version},
        "columns_enabled": enabled if 'enabled' in locals() else transformation.get("columns_enabled", []),
        "enrichments": _build_enrichments_payload(gab_name, gab_version, st.session_state.transfo_enrich_rows),
        "methods": methods_selected if 'methods_selected' in locals() else [],
        "overlay_python": st.session_state[code_key]
    }
    
    with st.spinner("Calcul des colonnes disponibles..."):
        df_test, _ = build_table_from_usage(temp_config, full=False, log_kpis=False)
        if df_test is not None:
            available_cols = list(df_test.columns)
        else:
            available_cols = enabled if 'enabled' in locals() else []
    
    st.success(f"✅ {len(available_cols)} colonnes disponibles")
    
    # État local
    state = st.session_state.transfo_adjust_state
    final_order = state["final_order"] or available_cols[:]
    final_excludes = state["final_excludes"]
    final_renames = state["final_renames"]
    
    # Synchronisation avec les colonnes effectives
    final_order = [c for c in final_order if c in available_cols] + \
                  [c for c in available_cols if c not in final_order]
    final_excludes = set([c for c in final_excludes if c in available_cols])
    
    # En-tête du tableau
    st.markdown("### Colonnes injectées")
    hdr0, hdr1, hdr2, hdr3, hdr4 = st.columns([1, 3, 2, 2, 2])
    with hdr0: st.markdown("**Inclure**")
    with hdr1: st.markdown("**Colonne**")
    with hdr2: st.markdown("**Exemple**")
    with hdr3: st.markdown("**Nouveau nom**")
    with hdr4: st.markdown("**Actions**")
    
    # Positions visibles pour l'ordre
    visible_positions = [i for i, c in enumerate(final_order) if c not in final_excludes]
    move_up_idx = move_down_idx = None
    toggled = False
    
    # Affichage des colonnes
    for idx, col in enumerate(final_order):
        included = col not in final_excludes
        pos_visible = visible_positions.index(idx) if included and idx in visible_positions else None
        
        c0, c1, c2, c3, c4 = st.columns([1, 3, 2, 2, 2])
        
        # Checkbox inclure/exclure
        with c0:
            new_state = st.checkbox(
                label=f"Inclure {col}",
                value=included,
                key=f"transfo_incl_{col}",
                label_visibility="collapsed"
            )
            if new_state != included:
                if new_state:
                    final_excludes.discard(col)
                else:
                    final_excludes.add(col)
                toggled = True
        
        # Nom de la colonne
        with c1:
            if included:
                st.markdown(f"**{col}**")
            else:
                st.markdown(_muted_html(col), unsafe_allow_html=True)
        
        # Exemple
        with c2:
            txt = _sample_value(col, gab_name, gab_version)
            st.markdown(_muted_html(txt) if not included else txt, unsafe_allow_html=True)
        
        # Renommage
        with c3:
            proposed = st.text_input(
                "Nouveau nom",
                value=final_renames.get(col, ""),
                placeholder=col,
                key=f"transfo_rename_{col}",
                label_visibility="collapsed",
                disabled=not included
            )
            if proposed.strip() and proposed.strip() != col:
                final_renames[col] = proposed.strip()
            elif col in final_renames:
                del final_renames[col]
        
        # Actions ordre
        with c4:
            b1, b2 = st.columns(2)
            with b1:
                if st.button("⬆️", key=f"transfo_up_{idx}", 
                            disabled=(not included or pos_visible is None or pos_visible == 0)):
                    move_up_idx = pos_visible
            with b2:
                if st.button("⬇️", key=f"transfo_dn_{idx}",
                            disabled=(not included or pos_visible is None or pos_visible == len(visible_positions)-1)):
                    move_down_idx = pos_visible
        
        st.markdown("<hr style='margin:5px 0'>", unsafe_allow_html=True)
    
    # Gérer les changements
    if toggled:
        state["final_excludes"] = final_excludes
        st.rerun()
    
    if move_up_idx is not None:
        cur = visible_positions[move_up_idx]
        prev = visible_positions[move_up_idx - 1]
        final_order[prev], final_order[cur] = final_order[cur], final_order[prev]
        state["final_order"] = final_order
        st.rerun()
    
    if move_down_idx is not None:
        cur = visible_positions[move_down_idx]
        nxt = visible_positions[move_down_idx + 1]
        final_order[nxt], final_order[cur] = final_order[cur], final_order[nxt]
        state["final_order"] = final_order
        st.rerun()
    
    # Mise à jour de l'état
    state["final_order"] = final_order
    state["final_excludes"] = final_excludes
    state["final_renames"] = final_renames

# ==================== TAB PREVIEW ====================
with tab_preview:
    st.subheader("👀 Prévisualisation finale")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        mode = st.radio("Mode", ["Preview (20 lignes)", "Complet"], index=0)
        preview_mode = (mode == "Preview (20 lignes)")
    with col2:
        if st.button("🔄 Générer la preview", type="primary", use_container_width=True):
            state = st.session_state.transfo_adjust_state
            
            full_config = {
                "gabarit_base": {"name": gab_name, "version": gab_version},
                "columns_enabled": enabled if 'enabled' in locals() else transformation.get("columns_enabled", []),
                "enrichments": _build_enrichments_payload(gab_name, gab_version, st.session_state.transfo_enrich_rows),
                "methods": methods_selected if 'methods_selected' in locals() else [],
                "overlay_python": st.session_state[code_key],
                "final_order": state["final_order"],
                "final_excludes": list(state["final_excludes"]),
                "final_renames": state["final_renames"]
            }
            
            with st.spinner("Génération en cours..."):
                # 👉 Toujours construire la table complète, puis tronquer l'affichage si Preview
                df_full, error = build_table_from_usage(
                    full_config,
                    full=True,      # <— toujours full
                    log_kpis=True
                )
            
            if error:
                st.error(f"❌ Erreur : {error}")
                st.session_state["_builder_preview"] = None
            elif df_full is None or df_full.empty:
                st.warning("Aucun résultat")
                st.session_state["_builder_preview"] = None
            else:
                st.session_state["_builder_preview"] = {
                    "df_full": df_full,
                    "preview_mode": preview_mode
                }
                n_display = min(20, len(df_full)) if preview_mode else len(df_full)
                st.success(f"✅ {n_display} lignes × {len(df_full.columns)} colonnes")
    
    # Affichage / Stats
    if st.session_state.get("_builder_preview"):
        df_full = st.session_state["_builder_preview"]["df_full"]
        preview_mode = st.session_state["_builder_preview"]["preview_mode"]
        df_display = df_full.head(20) if preview_mode else df_full
        
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        with st.expander("📊 Statistiques"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Lignes", f"{len(df_display):,}")
            with col2:
                st.metric("Colonnes", len(df_display.columns))
            with col3:
                memory_mb = df_display.memory_usage(deep=True).sum() / 1024 / 1024
                st.metric("Mémoire", f"{memory_mb:.2f} MB")

# ==================== BARRE D'ACTIONS ====================
st.markdown("---")

col1, col2, col3 = st.columns([2, 2, 1])

with col1:
    if st.button("💾 Sauvegarder tout", type="primary", use_container_width=True):
        # Construire la config finale
        state = st.session_state.transfo_adjust_state
        
        final_transformation = {
            "name": name,
            "version": version,
            "description": transformation.get("description", ""),
            "gabarit_base": {"name": gab_name, "version": gab_version},
            "columns_enabled": enabled if 'enabled' in locals() else transformation.get("columns_enabled", []),
            "enrichments": _build_enrichments_payload(gab_name, gab_version, st.session_state.transfo_enrich_rows),
            "methods": methods_selected if 'methods_selected' in locals() else [],
            "overlay_python": st.session_state[code_key],
            "final_order": state["final_order"],
            "final_excludes": list(state["final_excludes"]),
            "final_renames": state["final_renames"]
        }
        
        # Sauvegarder
        update_transformation(name, version, final_transformation)
        st.success("✅ Transformation sauvegardée !")
        st.balloons()

with col2:
    if st.button("🔄 Réinitialiser", use_container_width=True):
        # Réinitialiser tous les états
        if "transfo_enrich_rows" in st.session_state:
            del st.session_state.transfo_enrich_rows
        if code_key in st.session_state:
            del st.session_state[code_key]
        if "transfo_adjust_state" in st.session_state:
            del st.session_state.transfo_adjust_state
        st.rerun()

with col3:
    if st.button("📋 Retour", use_container_width=True):
        st.switch_page("pages/_2a_🔀_Detail_Transformation.py")