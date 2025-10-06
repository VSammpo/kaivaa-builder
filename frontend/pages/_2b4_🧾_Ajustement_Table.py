# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.services.gabarit_registry import (
    get_gabarit, get_default_preview,
)

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
            st.switch_page("pages/2_📚_Bibliotheque.py")
    with cols[1]:
        if st.button("🗂️ Détail du template",
                     type=("primary" if active == "detail" else "secondary"),
                     use_container_width=True, key=f"nav_detail_{active}"):
            if template_id:
                st.session_state.selected_template_detail = template_id
            st.switch_page("pages/_2a_📊_Detail_Livrable.py")
    with cols[2]:
        if st.button("⚙️ Paramètres généraux",
                     type=("primary" if active == "general" else "secondary"),
                     use_container_width=True, key=f"nav_general_{active}"):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_2b_➕_Form_Template.py")
    with cols[3]:
        if st.button("📑 Injection des données",
                     type=("primary" if active == "inject" else "secondary"),
                     use_container_width=True, key=f"nav_inject_{active}"):
            if template_id:
                st.session_state.selected_template = template_id
            st.switch_page("pages/_2b3_📑_Tables_Template.py")
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
        st.switch_page("pages/2_📚_Bibliotheque.py")
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
    st.info("Aucune table configurée (via « 📑 Injection des données »).")
    st.stop()


# ============================ Sélection de la table ============================

choices, index_map = [], []
for u in usages:
    tgt = u.get("excel_target") or {}
    sheet, table = tgt.get("sheet", ""), tgt.get("table", "")
    if sheet or table:
        label = f"{sheet} / {table} — {u.get('gabarit_name')} (v{u.get('gabarit_version','v1')})"
        choices.append(label)
        index_map.append(u)

sel = st.selectbox("Sélectionnez la table à ajuster", options=choices, index=0)
usage = index_map[choices.index(sel)]
gname, gver = usage.get("gabarit_name"), usage.get("gabarit_version", "v1")
sheet = (usage.get("excel_target") or {}).get("sheet", "")
table = (usage.get("excel_target") or {}).get("table", "")

tab_adjust, tab_preview = st.tabs(["🛠️ Ajustement", "👀 Prévisualisation"])


# ============================ Fonctions utilitaires ============================

def _resolve_default_cols(u: dict) -> list[str]:
    """Colonnes par défaut (base + enrich + sorties de méthodes)."""
    from backend.services.gabarit_registry import list_methods_for_gabarit

    g = get_gabarit(gname, gver)
    base = [c.name for c in (g.columns or [])]
    cols = u.get("columns_enabled") or base[:]

    # enrichissements : colonnes rapatriées
    for e in (u.get("enrichments") or []):
        for c in (e.get("columns") or []):
            if c not in cols:
                cols.append(c)

    # sorties de méthodes
    selm = set(u.get("methods") or [])
    if selm:
        allm = list_methods_for_gabarit(gname, gver) or []
        it = (allm.values() if isinstance(allm, dict) else allm)
        for m in it:
            if isinstance(m, dict) and m.get("name") in selm:
                outc = (m.get("output_column") or "").strip()
                if outc and outc not in cols:
                    cols.append(outc)

    return list(dict.fromkeys(cols))


def _sample_value(col: str) -> str:
    try:
        prev = get_default_preview(gname, gver) or {}
        rows, cols = prev.get("rows") or [], prev.get("columns") or []
        if rows and cols:
            df = pd.DataFrame(rows, columns=cols)
            if col in df.columns and not df.empty:
                return str(df[col].iloc[0])[:60]
    except Exception:
        pass
    return "—"


# ============================ Onglet AJUSTEMENT ============================

with tab_adjust:
    default_cols = _resolve_default_cols(usage)

    # État local (lié à la table sélectionnée)
    usage_key = f"{gname}|{gver}|{sheet}|{table}"
    if st.session_state.get("_adj_key") != usage_key:
        st.session_state["_adj_key"] = usage_key
        st.session_state["adj_final_order"] = list(usage.get("final_order") or default_cols[:])
        st.session_state["adj_final_excludes"] = set(usage.get("final_excludes") or [])

    # Synchronisation avec la structure par défaut
    final_order = [c for c in st.session_state["adj_final_order"] if c in default_cols] + \
                  [c for c in default_cols if c not in st.session_state["adj_final_order"]]
    final_excludes = set([c for c in st.session_state["adj_final_excludes"] if c in default_cols])

    # Métadonnées (source/type)
    g = get_gabarit(gname, gver)
    type_map = {c.name: (c.type or "text") for c in g.columns}
    source_map = {c: "gabarit" for c in (usage.get("columns_enabled") or [c.name for c in g.columns])}

    for e in (usage.get("enrichments") or []):
        if not e.get("path"):
            continue
        last = e["path"][-1]
        tgt_name = last[2]
        tgt_g = get_gabarit(tgt_name, "v1")
        for c in e.get("columns", []):
            source_map[c] = f"enrich:{tgt_name}"
            if c not in type_map:
                type_map[c] = next((col.type for col in (tgt_g.columns or []) if col.name == c), "text")

    for m in (usage.get("methods") or []):
        source_map[m] = f"method:{m}"
        type_map.setdefault(m, "unknown")

    # Tableau compact “à la Airtable”
    st.markdown("### Colonnes injectées")
    st.caption("Cliquez sur ⬆️/⬇️ pour modifier l'ordre, 🗑️ pour retirer la colonne de la **sortie finale** (sans impacter les sélections amont).")

    hdr1, hdr2, hdr3, hdr4, hdr5 = st.columns([3, 2, 2, 2, 2], gap="small")
    with hdr1:
        st.markdown("**Colonne**")
    with hdr2:
        st.markdown("**Source**")
    with hdr3:
        st.markdown("**Type**")
    with hdr4:
        st.markdown("**Exemple**")
    with hdr5:
        st.markdown("**Actions**")

    # positions visibles (pour des flèches correctes)
    visible_positions = [i for i, c in enumerate(final_order) if c not in final_excludes]

    move_up_idx = move_down_idx = drop_idx = None

    for vidx, pos in enumerate(visible_positions):
        col = final_order[pos]
        c1, c2, c3, c4, c5 = st.columns([3, 2, 2, 2, 2], gap="small")

        with c1:
            st.write(col)
        with c2:
            st.caption(source_map.get(col, "—"))
        with c3:
            st.caption(type_map.get(col, "text"))
        with c4:
            st.caption(_sample_value(col))
        with c5:
            b1, b2, b3 = st.columns([1, 1, 1], gap="small")
            with b1:
                if st.button("⬆️", key=f"adj_up_{usage_key}_{vidx}", use_container_width=True, disabled=(vidx == 0)):
                    move_up_idx = vidx
            with b2:
                if st.button("⬇️", key=f"adj_dn_{usage_key}_{vidx}", use_container_width=True, disabled=(vidx == len(visible_positions) - 1)):
                    move_down_idx = vidx
            with b3:
                if st.button("🗑️", key=f"adj_rm_{usage_key}_{vidx}", use_container_width=True):
                    drop_idx = vidx

        # séparateur discret
        st.markdown("<div style='border-bottom:1px dashed #e6e8eb; margin:6px 0 10px 0;'></div>", unsafe_allow_html=True)

    # Appliquer les actions
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

    if drop_idx is not None:
        col_to_drop = final_order[visible_positions[drop_idx]]
        final_excludes.add(col_to_drop)
        st.session_state["adj_final_excludes"] = final_excludes
        st.rerun()

    # Actions de page
    colA, colB = st.columns([1, 1])
    with colA:
        if st.button("💾 Enregistrer l’ajustement", type="primary", use_container_width=True, key=f"btn_save_adj_{usage_key}"):
            with DatabaseService.get_session() as db:
                ts = TemplateService(db)
                ts.update_usage_final_view(
                    template_id=template_id,
                    gabarit_name=gname,
                    gabarit_version=gver,
                    final_order=final_order,
                    final_excludes=list(final_excludes),
                )
            st.success("✅ Ajustement enregistré")
            st.rerun()
    with colB:
        if st.button("🔁 Réinitialiser", use_container_width=True, key=f"btn_reset_adj_{usage_key}"):
            with DatabaseService.get_session() as db:
                ts = TemplateService(db)
                ts.update_usage_final_view(
                    template_id=template_id,
                    gabarit_name=gname,
                    gabarit_version=gver,
                    final_order=default_cols,
                    final_excludes=[],
                )
            st.info("Réinitialisé")
            st.session_state["adj_final_order"] = default_cols[:]
            st.session_state["adj_final_excludes"] = set()
            st.rerun()


# ============================ Onglet PRÉVISUALISATION ============================

with tab_preview:
    st.caption(f"Feuille : **{sheet}** • Table : **{table}**")
    st.markdown("---")

    def _load_preview_df_for_gabarit(name: str, ver: str):
        try:
            prev = get_default_preview(name, ver) or {}
            rows, cols = prev.get("rows") or [], prev.get("columns") or []
            if rows and cols:
                return pd.DataFrame(rows, columns=cols)
        except Exception:
            pass
        return None

    # 1) DF principal (gabarit de départ)
    df_main = _load_preview_df_for_gabarit(gname, gver)

    # 2) appliquer enrichissements (joins successifs) si possible
    complete = df_main is not None
    df = df_main.copy() if df_main is not None else None

    if df is not None:
        for e in (usage.get("enrichments") or []):
            path = e.get("path") or []
            if not path:
                continue
            last = path[-1]
            tgt_name, tgt_ver = last[2], "v1"
            df_tgt = _load_preview_df_for_gabarit(tgt_name, tgt_ver)
            if df_tgt is None:
                complete = False
                break
            # join sur la clé du premier saut
            left_key = path[0][1]
            right_key = path[0][3]
            if left_key in df.columns and right_key in df_tgt.columns:
                df = df.merge(
                    df_tgt[[right_key] + [c for c in e.get("columns", []) if c in df_tgt.columns]].drop_duplicates(),
                    left_on=left_key, right_on=right_key, how="left"
                )
                if right_key in df.columns:
                    df.drop(columns=[right_key], inplace=True)
            else:
                complete = False
                break

    if not complete or df is None:
        st.info("ℹ️ Toutes les gabarits utilisés pour construire la table doivent avoir une **donnée par défaut** pour permettre la prévisualisation.")
    else:
        # 3) Appliquer l'ordre/exclusions pour la vue finale
        final_cols_cfg = usage.get("final_order") or df.columns.tolist()
        final_excl_cfg = set(usage.get("final_excludes") or [])
        final_cols = [c for c in final_cols_cfg if c in df.columns and c not in final_excl_cfg] + \
                     [c for c in df.columns if c not in final_cols_cfg and c not in final_excl_cfg]

        st.dataframe(df[final_cols].head(15), use_container_width=True, hide_index=True)
