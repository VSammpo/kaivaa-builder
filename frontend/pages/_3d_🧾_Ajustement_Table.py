# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from typing import Optional, Tuple
from loguru import logger

# ---------------------------------------------------------------------
# Bootstrap import path
# ---------------------------------------------------------------------
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.services.gabarit_registry import get_gabarit, get_default_preview
from code_editor import code_editor
from backend.services.parameter_service import ParameterService
# ✅ IMPORT UNIQUE depuis le service centralisé
from backend.services.table_builder_service import build_table_from_usage

# ============================ HELPERS SPÉCIFIQUES À L'UI ============================

def _check_script_sync(usage: dict, persist_key: str) -> tuple[bool, str]:
    """Vérifie si le script en session diffère de celui en DB."""
    db_script = (usage.get("overlay_python") or "").strip()
    session_script = st.session_state.get(persist_key, "").strip()
    
    if db_script != session_script:
        return False, "⚠️ Le script en mémoire diffère de la version sauvegardée en base"
    return True, ""

def _force_recalc_columns(usage: dict, template_id: int) -> list[str]:
    """
    ✅ CORRECTION : Pour les scripts complexes, utiliser TOUJOURS FULL mode.
    Le PREVIEW (20 lignes) ne peut pas exécuter correctement les agrégations.
    """
    # 1. Recharger l'usage FRAIS depuis la DB
    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        fresh_usage = ts.get_gabarit_usage_by_target(
            template_id,
            usage.get("gabarit_name"),
            usage.get("gabarit_version", "v1"),
            (usage.get("excel_target", {}).get("sheet") or "").strip(),
            (usage.get("excel_target", {}).get("table") or "").strip()
        )
    
    if not fresh_usage:
        fresh_usage = usage
    
    # 2. Charger les paramètres avec leurs valeurs par défaut
    params_dict = {}
    try:
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            template_config = ts.load_template_config(template_id)
            params_dict = {p.name: ParameterService.get_default_value(p) for p in template_config.parameters}
    except Exception:
        pass
    
    # ✅ 3. CORRECTION : Vérifier si le script est complexe (> 1000 caractères)
    script = (fresh_usage.get("overlay_python") or "").strip()
    is_complex_script = len(script) > 1000
    
    if is_complex_script:
        logger.info(f"[_force_recalc_columns] 🔬 Script complexe détecté ({len(script)} car.) → FULL mode obligatoire")
    
    # ✅ 4. Si script simple, essayer PREVIEW d'abord (rapide)
    if not is_complex_script:
        try:
            df, err = build_table_from_usage(fresh_usage, full=False, log_kpis=False, params=params_dict)
            
            # Vérifier que le résultat est cohérent (au moins 3 colonnes attendues)
            if df is not None and not df.empty and len(df.columns) >= 3:
                logger.info(f"[_force_recalc_columns] ✅ PREVIEW OK : {len(df.columns)} colonnes")
                return list(df.columns)
            else:
                logger.warning(f"[_force_recalc_columns] ⚠️ PREVIEW incohérent ({len(df.columns) if df is not None else 0} col.) → tentative FULL")
        except Exception as e:
            logger.warning(f"[_force_recalc_columns] ⚠️ Erreur PREVIEW : {e} → tentative FULL")
    
    # ✅ 5. FULL mode (scripts complexes OU échec preview)
    try:
        logger.info(f"[_force_recalc_columns] 🔄 Exécution FULL mode...")
        df, err = build_table_from_usage(fresh_usage, full=True, log_kpis=False, params=params_dict)
        
        if df is not None and not df.empty:
            logger.info(f"[_force_recalc_columns] ✅ FULL OK : {len(df.columns)} colonnes")
            return list(df.columns)
        
        if err:
            logger.error(f"[_force_recalc_columns] ❌ Erreur construction : {err}")
    except Exception as e:
        logger.error(f"[_force_recalc_columns] ❌ Exception FULL : {e}")
        import traceback
        logger.debug(traceback.format_exc())
    
    # ✅ 6. Fallback : colonnes de base du gabarit
    try:
        g = get_gabarit(fresh_usage.get("gabarit_name"), fresh_usage.get("gabarit_version", "v1"))
        if g:
            base_cols = [c.name for c in (g.columns or [])]
            logger.warning(f"[_force_recalc_columns] ⚠️ Fallback colonnes de base : {len(base_cols)} colonnes")
            return base_cols
    except Exception:
        pass
    
    logger.error(f"[_force_recalc_columns] ❌ Impossible de calculer les colonnes")
    return []

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


# ============================ FIN DES HELPERS ============================

st.set_page_config(page_title="Ajustement de la table", page_icon="🧾", layout="wide")


# ============================ Navbar (5 boutons) ============================

def render_template_subnav(active: str, template_id: int | None):
    cols = st.columns([1, 1, 1, 1, 1])
    with cols[0]:
        if st.button("← Retour bibliothèque", use_container_width=True, key=f"nav_back_{active}"):
            if "selected_template" in st.session_state:
                del st.session_state.selected_template
            if "selected_template_detail" in st.session_state:
                del st.session_state.selected_template_detail
            st.switch_page("pages/3_📚_Bibliotheque.py")
    with cols[1]:
        if st.button("🗂️ Détail du template",
                     type=("primary" if active == "detail" else "secondary"),
                     use_container_width=True, key=f"nav_detail_{active}"):
            if template_id:
                st.session_state.selected_template_detail = template_id
            st.switch_page("pages/_3a_📊_Detail_Livrable.py")
    with cols[2]:
        if st.button("⚙️ Paramètres généraux",
                     type=("primary" if active == "general" else "secondary"),
                     use_container_width=True, key=f"nav_general_{active}"):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_3b_➕_Form_Template.py")
    with cols[3]:
        if st.button("📑 Injection des données",
                    type=("primary" if active == "inject" else "secondary"),
                    use_container_width=True, key=f"nav_inject_{active}"):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_3c_📑_Tables_Template.py")

    with cols[4]:
        st.button("🧾 Ajustement de la table",
                  type="primary" if active == "adjust" else "secondary",
                  use_container_width=True, key=f"nav_adjust_{active}")
    st.divider()


# ============================ Guard & navbar ============================

if "selected_template" not in st.session_state and "selected_template_detail" in st.session_state:
    st.session_state.selected_template = st.session_state.selected_template_detail

if "selected_template" not in st.session_state or not st.session_state.selected_template:
    st.error("Aucun template sélectionné.")
    if st.button("← Retour bibliothèque"):
        st.switch_page("pages/3_📚_Bibliotheque.py")
    st.stop()

template_id = st.session_state.selected_template
render_template_subnav("adjust", template_id)


# ============================ Charger primitives ============================

with DatabaseService.get_session() as db:
    ts = TemplateService(db)
    tpl = ts.get_template(template_id)
    cfg = ts.get_config(template_id)
    usages = cfg.get("gabarit_usages", []) if isinstance(cfg, dict) else []
    tpl_name, tpl_version = tpl.name, tpl.version

st.title(f"🧾 Ajustement de la table — {tpl_name} (v{tpl_version})")

if not usages:
    st.info("Aucune table configurée (via « 🔐 Injection des données »).")
    st.stop()


# ============================ Sélection de la table ============================

choices, index_map = [], []
for u in usages:
    tgt = u.get("excel_target") or {}
    sheet_u, table_u = tgt.get("sheet", ""), tgt.get("table", "")
    if sheet_u or table_u:
        label = f"{sheet_u} / {table_u} — {u.get('gabarit_name')} (v{u.get('gabarit_version','v1')})"
        choices.append(label)
        index_map.append(u)

sel = st.selectbox("Sélectionnez la table à ajuster", options=choices, index=0)
usage = index_map[choices.index(sel)]
gname, gver = usage.get("gabarit_name"), usage.get("gabarit_version", "v1")
sheet = (usage.get("excel_target") or {}).get("sheet", "")
table = (usage.get("excel_target") or {}).get("table", "")

tab_script, tab_adjust, tab_preview = st.tabs(["🧪 Script Python", "🛠️ Ajustement", "👀 Prévisualisation"])


# ============================ Onglet 1 — Script Python ============================
with tab_script:
    st.caption(f"Feuille : **{sheet}** • Table : **{table}**")
    
    # ✅ AFFICHAGE DES PARAMÈTRES DISPONIBLES
    st.markdown("### 📋 Paramètres disponibles")

    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        template_config = ts.load_template_config(template_id)
        params = template_config.parameters

    if params:
        st.markdown("**Chaque paramètre est utilisable de deux façons :**")
        st.caption("• Accès direct (recommandé) : `NomParam`  • Accès dict : `params['NomParam']`")

        cols = st.columns(3)
        for idx, param in enumerate(params):
            with cols[idx % 3]:
                with st.container(border=True):
                    st.markdown(f"**{param.name}**")
                    st.caption(f"Type : {param.type}")

                    if param.default:
                        st.caption(f"Défaut : `{param.default}`")

        st.markdown("---")
        st.info("💡 Exemples : `df = df[df['Marque'] == Sous_Marque]` **ou** `df = df[df['Marque'] == params['Sous_Marque']]`")
    else:
        st.info("Aucun paramètre défini pour ce template")

    st.markdown("---")
    
    # ✅ VÉRIFICATION DE LA SYNCHRONISATION
    persist_key = f"code_persist_{template_id}_{gname}_{gver}_{sheet}_{table}"
    
    if persist_key not in st.session_state:
        st.session_state[persist_key] = (usage.get("overlay_python") or "").strip()
    
    is_synced, warning = _check_script_sync(usage, persist_key)
    if not is_synced:
        st.warning(warning + " — Cliquez sur '💾 Enregistrer' pour persister les changements.")
    
    # ✅ CODE EDITOR AVEC PERSISTANCE
    st.markdown("### 🔧 Script Python de transformation")
    st.caption("💡 Variables : `df` (DataFrame), `pd` (pandas), `params` (dict), et chaque paramètre accessible par son nom (ex : `Secteur`).")
    st.caption("⚠️ Le script doit réassigner `df` (ex : `df = df[df['Col'] == Secteur]`).")

    custom_buttons = [{
        "name": "Copier", "feather": "Copy", "hasText": True,
        "commands": ["copyAll"], "style": {"top": "0.46rem", "right": "0.4rem"}
    }]

    with st.form(f"overlay_form_{persist_key}", clear_on_submit=False, border=True):
        editor_result = code_editor(
            st.session_state[persist_key],
            lang="python", height=300, theme="contrast", shortcuts="vscode",
            focus=False, buttons=custom_buttons, allow_reset=True,
            options={"wrap": True, "showLineNumbers": True, "highlightActiveLine": True,
                    "enableLiveAutocompletion": True, "enableBasicAutocompletion": True},
            key=f"overlay_editor_{persist_key}",
            response_mode=["submit", "blur"],
        )

        # Extraction robuste
        if editor_result:
            _new = None
            if isinstance(editor_result, dict):
                _new = editor_result.get("text") or editor_result.get("content") or editor_result.get("code")
            elif isinstance(editor_result, str):
                _new = editor_result
            if isinstance(_new, str):
                st.session_state[persist_key] = _new

        st.caption(f"📄 {len(st.session_state[persist_key])} caractères")
        st.markdown("---")

        c1, c2, c3 = st.columns([1, 1, 1], gap="large")
        with c1:
            do_preview = st.form_submit_button("🧪 Prévisualiser le script", use_container_width=True)
        with c2:
            do_save = st.form_submit_button("💾 Enregistrer le script", type="primary", use_container_width=True)
        with c3:
            do_recalc = st.form_submit_button("🔄 Recalculer colonnes", use_container_width=True)

    code = st.session_state[persist_key]

    # === Actions après le form ===
    if do_save:
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            cfg2 = ts.get_config(template_id)
            usages2 = cfg2.get("gabarit_usages", []) or []
            for uu in usages2:
                tgt2 = uu.get("excel_target") or {}
                if (
                    uu.get("gabarit_name") == gname
                    and (uu.get("gabarit_version") or "v1") == gver
                    and (tgt2.get("sheet") or "") == sheet
                    and (tgt2.get("table") or "") == table
                ):
                    uu["overlay_python"] = code
                    break
            cfg2["gabarit_usages"] = usages2
            ts.update_config(template_id, cfg2)

        st.success("✅ Script enregistré")
        st.rerun()

    if do_recalc:
        with st.spinner("Recalcul des colonnes disponibles..."):
            new_cols = _force_recalc_columns(usage, template_id)
            st.success(f"✅ {len(new_cols)} colonnes disponibles après exécution du pipeline complet")
            with st.expander("📋 Colonnes disponibles", expanded=True):
                for col in new_cols:
                    st.caption(f"• {col}")

    if do_preview:
        params_dict = {}
        for param in template_config.parameters:
            params_dict[param.name] = ParameterService.get_default_value(param)

        usage_test = dict(usage)
        usage_test["overlay_python"] = code

        # ✅ UTILISER LE SERVICE CENTRALISÉ
        df_after_script, error = build_table_from_usage(usage_test, full=True, log_kpis=True, params=params_dict)

        if error:
            st.error(f"❌ Erreur :\n```\n{error}\n```")
        elif df_after_script is None or df_after_script.empty:
            st.warning("Aucun résultat")
        else:
            st.success(f"✅ {len(df_after_script)} lignes, {len(df_after_script.columns)} colonnes")
            st.dataframe(df_after_script.head(20), use_container_width=True, hide_index=True)
            
            with st.expander("📋 Colonnes disponibles après script", expanded=False):
                for col in df_after_script.columns:
                    st.caption(f"• {col}")


# ============================ Onglet 2 — AJUSTEMENT ============================
with tab_adjust:
    # ✅ CORRECTION MAJEURE : RECHARGER L'USAGE DEPUIS LA DB (source de vérité)
    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        fresh_usage = ts.get_gabarit_usage_by_target(
            template_id,
            usage.get("gabarit_name"),
            usage.get("gabarit_version", "v1"),
            (usage.get("excel_target", {}).get("sheet") or "").strip(),
            (usage.get("excel_target", {}).get("table") or "").strip()
        )
    
    # Si pas trouvé, utiliser l'usage courant
    if not fresh_usage:
        fresh_usage = usage
    
    # ✅ Utiliser le FRESH usage pour calculer les colonnes disponibles
    # (les colonnes clés sont maintenant automatiquement incluses par le backend)
    with st.spinner("🔄 Calcul des colonnes disponibles (inclut script + enrichissements + méthodes + clés)..."):
        default_cols = _force_recalc_columns(fresh_usage, template_id)


    
    if not default_cols:
        st.error("❌ Impossible de calculer les colonnes. Vérifiez la configuration dans l'onglet 'Script Python'.")
        st.info("💡 Les colonnes peuvent être absentes si :")
        st.caption("• Le script Python contient une erreur")
        st.caption("• Les enrichissements pointent vers des tables inexistantes")
        st.caption("• Les méthodes utilisent des colonnes manquantes")
        st.stop()
    
    st.success(f"✅ {len(default_cols)} colonnes disponibles")

    # --- état local lié à la table sélectionnée ---
    usage_key = f"{gname}|{gver}|{sheet}|{table}"
    if st.session_state.get("_adj_key") != usage_key:
        st.session_state["_adj_key"] = usage_key
        # ✅ Utiliser fresh_usage au lieu de usage
        st.session_state["adj_final_order"] = list(fresh_usage.get("final_order") or default_cols[:])
        st.session_state["adj_final_excludes"] = set(fresh_usage.get("final_excludes") or [])
        st.session_state["adj_final_renames"] = dict(fresh_usage.get("final_renames") or {})

    # synchronisation avec la structure effective
    final_order = [c for c in st.session_state["adj_final_order"] if c in default_cols] + \
                  [c for c in default_cols if c not in st.session_state["adj_final_order"]]
    final_excludes = set([c for c in st.session_state["adj_final_excludes"] if c in default_cols])
    final_renames = dict(st.session_state["adj_final_renames"])

    # ✅ Utiliser fresh_usage pour les métadonnées
    g = get_gabarit(gname, gver)
    type_map = {c.name: (c.type or "text") for c in (g.columns or [])}
    
    # ✅ CORRECTION : Inclure TOUJOURS les colonnes clés (is_key=True) même si non sélectionnées
    key_cols = [c.name for c in (g.columns or []) if getattr(c, 'is_key', False)]
    enabled_cols = fresh_usage.get("columns_enabled") or [c.name for c in (g.columns or [])]
    
    # Combiner colonnes activées + colonnes clés (déduplication en gardant l'ordre)
    all_base_cols = list(dict.fromkeys(enabled_cols + key_cols))
    source_map = {c: "gabarit [CLÉ]" if c in key_cols else "gabarit" for c in all_base_cols}

    for e in (fresh_usage.get("enrichments") or []):
        if not e.get("path"):
            continue
        last = e["path"][-1]
        tgt_name = last[2]
        tgt_g = get_gabarit(tgt_name, "v1")
        for c in (e.get("columns") or []):
            source_map[c] = f"enrich:{tgt_name}"
            if c not in type_map:
                type_map[c] = next((col.type for col in (tgt_g.columns or []) if col.name == c), "text")

    for m in (fresh_usage.get("methods") or []):
        source_map[m] = f"method:{m}"
        type_map.setdefault(m, "unknown")

    # --- entête ---
    
    # 🔑 Afficher un avertissement si colonnes clés détectées
    if key_cols:
        st.info(f"🔑 **Colonnes clés protégées** (toujours présentes) : {", ".join(key_cols)}")

    st.markdown("### Colonnes injectées")
    st.caption(
        "Cochez/décochez pour inclure/exclure une colonne dans la **sortie finale** "
        "(les sélections amont ne sont pas modifiées)."
    )

    hdr0, hdr1, hdr2, hdr3, hdr4, hdr5 = st.columns([1.2, 3, 2, 2, 2, 2], gap="small")
    with hdr0: st.markdown("**Inclure**")
    with hdr1: st.markdown("**Colonne**")
    with hdr2: st.markdown("**Source**")
    with hdr3: st.markdown("**Type**")
    with hdr4: st.markdown("**Exemple**")
    with hdr5: st.markdown("**Actions**")

    # indices visibles (pour gérer ↑↓ sur les colonnes incluses)
    visible_positions = [i for i, c in enumerate(final_order) if c not in final_excludes]
    move_up_idx = move_down_idx = None
    toggled = False

    for idx, col in enumerate(final_order):
        included = col not in final_excludes
        pos_visible = visible_positions.index(idx) if included and idx in visible_positions else None

        c0, c1, c2, c3, c4, c5 = st.columns([1.2, 3, 2, 2, 2, 2], gap="small")

        # Inclure / Exclure
        with c0:
            new_state = st.checkbox(
                label=f"Inclure {col}",
                value=included,
                key=f"incl_{usage_key}_{col}",
                label_visibility="collapsed",
            )
            if new_state != included:
                if new_state:
                    final_excludes.discard(col)
                else:
                    final_excludes.add(col)
                toggled = True

        # Renommage (grisé si exclu)
        with c1:
            proposed = st.text_input(
                "Nouveau nom",
                value=final_renames.get(col, ""),
                placeholder=col,
                key=f"rename_{usage_key}_{col}",
                label_visibility="collapsed",
                disabled=not new_state,
            )
            if proposed.strip():
                final_renames[col] = proposed.strip()
            else:
                final_renames.pop(col, None)

        # Source / Type / Exemple (gris si exclu)
        with c2:
            txt = source_map.get(col, "—")
            st.markdown(_muted_html(txt) if not new_state else txt, unsafe_allow_html=True)
        with c3:
            txt = type_map.get(col, "text")
            st.markdown(_muted_html(txt) if not new_state else txt, unsafe_allow_html=True)
        with c4:
            txt = _sample_value(col, gname, gver)
            st.markdown(_muted_html(txt) if not new_state else txt, unsafe_allow_html=True)

        # Actions : ordre ↑↓ (désactivées si exclu)
        with c5:
            b1, b2, _ = st.columns([1, 1, 1], gap="small")
            with b1:
                if st.button("⬆️", key=f"adj_up_{usage_key}_{idx}", use_container_width=True,
                             disabled=(not new_state or pos_visible is None or pos_visible == 0)):
                    move_up_idx = pos_visible
            with b2:
                if st.button("⬇️", key=f"adj_dn_{usage_key}_{idx}", use_container_width=True,
                             disabled=(not new_state or pos_visible is None or pos_visible == len(visible_positions)-1)):
                    move_down_idx = pos_visible

        st.markdown(
            "<div style='border-bottom:1px dashed #e6e8eb; margin:6px 0 10px 0;'></div>",
            unsafe_allow_html=True,
        )

    # si un toggle a eu lieu -> mémoriser et relancer pour recalculer visible_positions
    if toggled:
        st.session_state["adj_final_excludes"] = set(final_excludes)
        st.session_state["adj_final_renames"] = dict(final_renames)
        st.rerun()

    # gestion ↑↓ (sur les visibles seulement)
    if move_up_idx is not None:
        cur = visible_positions[move_up_idx]
        prev = visible_positions[move_up_idx - 1]
        final_order[prev], final_order[cur] = final_order[cur], final_order[prev]
        st.session_state["adj_final_order"] = final_order
        st.rerun()

    if move_down_idx is not None:
        cur = visible_positions[move_down_idx]
        nxt = visible_positions[move_down_idx + 1]
        final_order[nxt], final_order[cur] = final_order[cur], final_order[nxt]
        st.session_state["adj_final_order"] = final_order
        st.rerun()

    # --- actions globales ---
    colA, colB = st.columns([1, 1])
    with colA:
        if st.button("💾 Enregistrer l'ajustement", type="primary", use_container_width=True, key=f"btn_save_adj_{usage_key}"):
            with DatabaseService.get_session() as db:
                ts = TemplateService(db)
                ts.update_usage_final_view(
                    template_id=template_id,
                    gabarit_name=gname,
                    gabarit_version=gver,
                    final_order=final_order,
                    final_excludes=list(final_excludes),
                    final_renames=final_renames,
                )
            st.session_state["adj_final_order"] = final_order[:]
            st.session_state["adj_final_excludes"] = set(final_excludes)
            st.session_state["adj_final_renames"] = dict(final_renames)
            st.success("✅ Ajustement enregistré")

    with colB:
        if st.button("🔄 Réinitialiser", use_container_width=True, key=f"btn_reset_adj_{usage_key}"):
            eff = _force_recalc_columns(usage, template_id)
            with DatabaseService.get_session() as db:
                ts = TemplateService(db)
                ts.update_usage_final_view(
                    template_id=template_id,
                    gabarit_name=gname,
                    gabarit_version=gver,
                    final_order=eff,
                    final_excludes=[],
                    final_renames={},
                )
            st.session_state["adj_final_order"] = eff[:]
            st.session_state["adj_final_excludes"] = set()
            st.session_state["adj_final_renames"] = {}
            st.info("Réinialisé")
            st.rerun()


# ============================ Onglet 3 — PRÉVISUALISATION ============================
with tab_preview:
    st.caption(f"Feuille : **{sheet}** • Table : **{table}**")
    
    # ✅ AFFICHER L'ÉTAT DE LA CONFIGURATION
    persist_key = f"code_persist_{template_id}_{gname}_{gver}_{sheet}_{table}"
    is_synced, _ = _check_script_sync(usage, persist_key)
    
    with st.expander("ℹ️ État de la configuration", expanded=False):
        st.caption(f"Script synchronisé : {'✅ Oui' if is_synced else '❌ Non (version en mémoire)'}")
        st.caption(f"Enrichissements : {len(usage.get('enrichments', []))}")
        st.caption(f"Méthodes : {', '.join(usage.get('methods', [])) or '—'}")
        st.caption(f"Renommages : {len(usage.get('final_renames', {}))}")
        st.caption(f"Exclusions : {len(usage.get('final_excludes', []))}")
    
    st.markdown("---")
    
    # ✅ RECHARGER TOUJOURS DEPUIS LA DB (source de vérité)
    with DatabaseService.get_session() as db:
        ts = TemplateService(db)
        fresh_usage = ts.get_gabarit_usage_by_target(
            template_id,
            usage.get("gabarit_name"),
            usage.get("gabarit_version", "v1"),
            (usage.get("excel_target", {}).get("sheet") or "").strip(),
            (usage.get("excel_target", {}).get("table") or "").strip()
        )
        
        if not fresh_usage:
            fresh_usage = usage
        
        template_config = ts.load_template_config(template_id)
    
    # ✅ PARAMÈTRES AVEC VALEURS PAR DÉFAUT
    params_dict = {p.name: ParameterService.get_default_value(p) for p in template_config.parameters}
    
    if params_dict:
        st.caption("📋 Paramètres utilisés (valeurs par défaut) : " + ", ".join(f"{k}={v}" for k, v in params_dict.items()))
    
    # ✅ AVERTIR SI VERSION EN MÉMOIRE ≠ DB
    if not is_synced:
        st.warning("⚠️ Attention : cette prévisualisation utilise la version SAUVEGARDÉE du script (pas celle en mémoire). Cliquez sur '💾 Enregistrer' dans l'onglet 'Script Python' pour persister vos changements.")
    
    # ✅ FORCER LE RECALCUL DES COLONNES EFFECTIVES
    with st.spinner("Calcul des colonnes disponibles..."):
        effective_cols = _force_recalc_columns(fresh_usage, template_id)
    
    # ✅ MISE À JOUR DE L'ORDRE FINAL
    current_order = fresh_usage.get("final_order") or []
    missing_cols = [c for c in effective_cols if c not in current_order]
    if missing_cols:
        fresh_usage["final_order"] = current_order + missing_cols
        st.info(f"ℹ️ {len(missing_cols)} nouvelle(s) colonne(s) détectée(s) (ajoutées à la fin)")
    
    # ✅ S'ASSURER QUE LES RENOMMAGES SONT PRÉSENTS
    if "final_renames" not in fresh_usage or not fresh_usage["final_renames"]:
        if st.session_state.get("adj_final_renames"):
            fresh_usage["final_renames"] = dict(st.session_state["adj_final_renames"])
    
    # ✅ CORRECTION MAJEURE : UTILISER LE SERVICE CENTRALISÉ (EXACTEMENT COMME L'INJECTION RÉELLE)
    st.markdown("### 🔄 Exécution du pipeline complet")
    with st.spinner("Génération du tableau final..."):
        df_prev, err = build_table_from_usage(fresh_usage, full=True, log_kpis=True, params=params_dict)
    
    if err:
        st.error(f"❌ Erreur pipeline :\n\n```\n{err}\n```")
        st.info("💡 Vérifiez :")
        st.caption("• Que le script Python ne contient pas d'erreurs")
        st.caption("• Que les enrichissements sont correctement configurés")
        st.caption("• Que les méthodes sont valides")
    elif df_prev is None or df_prev.empty:
        st.info("Pipeline exécuté mais aucun résultat affichable.")
    else:
        st.success(f"✅ Résultat FINAL — {len(df_prev)} lignes × {len(df_prev.columns)} colonnes")
        
        # ✅ AFFICHER LES TRANSFORMATIONS APPLIQUÉES
        col1, col2 = st.columns(2)
        with col1:
            if fresh_usage.get("final_renames"):
                with st.expander("🏷️ Renommages appliqués", expanded=False):
                    for old, new in fresh_usage["final_renames"].items():
                        st.caption(f"• {old} → {new}")
        
        with col2:
            if fresh_usage.get("final_excludes"):
                with st.expander("🚫 Colonnes exclues", expanded=False):
                    for col in fresh_usage["final_excludes"]:
                        st.caption(f"• {col}")
        
        # ✅ AFFICHER LE TABLEAU
        st.dataframe(df_prev.head(20), use_container_width=True, hide_index=True)
        
        # ✅ COLONNES FINALES
        with st.expander("📋 Colonnes du tableau final", expanded=False):
            for idx, col in enumerate(df_prev.columns, 1):
                st.caption(f"{idx}. {col}")