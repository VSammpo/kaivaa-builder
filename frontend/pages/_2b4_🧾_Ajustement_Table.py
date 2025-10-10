# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
import traceback
from typing import Dict, Any, Optional, Tuple

# ---------------------------------------------------------------------
# Bootstrap import path
# ---------------------------------------------------------------------
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.services.gabarit_registry import (
    get_gabarit,
    get_default_preview,
    list_methods_for_gabarit,  # utilisé pour fallback colonnes
)

st.set_page_config(page_title="Ajustement de la table", page_icon="🧾", layout="wide")

from code_editor import code_editor
from backend.services.parameter_service import ParameterService
from backend.models.template_config import ParameterConfig
from backend.services.table_builder_service import _debug_merge

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


# ============================ Utilitaires communs ============================

def _load_df_for_gabarit(name: str, ver: str, full: bool = False) -> tuple[pd.DataFrame | None, bool]:
    """
    Retourne (df, is_preview).

    Comportement voulu :
      - full=True  : on TENTE d'abord le FULL (registry.dataset -> dataset_service).
                     S'il est indisponible, on NE RETOMBE PAS en preview.
                     => on renvoie (None, False) pour que l'appelant affiche une erreur claire.
      - full=False : on prend la preview persistée si dispo,
                     sinon on lit le FULL et on renvoie head(20) en mode preview.

    NB : le pré-typage/alignement est fait PLUS TARD (dans _compose_with_enrichments) si nécessaire.
    """
    # ----- chemin FULL -----
    if full:
        # 1) via registry facade
        try:
            from backend.services.gabarit_registry import get_default_dataframe
            df_full = get_default_dataframe(name, ver)
            if isinstance(df_full, pd.DataFrame) and not df_full.empty:
                return df_full, False  # FULL ok
        except Exception:
            pass
        # 2) via dataset_service direct
        try:
            from backend.services.dataset_service import get_default_dataframe_for_gabarit
            df_full = get_default_dataframe_for_gabarit(name, ver)
            if isinstance(df_full, pd.DataFrame) and not df_full.empty:
                return df_full, False  # FULL ok
        except Exception:
            pass

        # FULL demandé mais indisponible -> PAS DE FALLBACK preview ici
        return None, False

    # ----- chemin PREVIEW (rapide) -----
    # preview persisté (registry)
    try:
        from backend.services.gabarit_registry import get_default_preview
        prev = get_default_preview(name, ver) or {}
        rows, cols = prev.get("rows") or [], prev.get("columns") or []
        if rows and cols:
            return pd.DataFrame(rows, columns=cols), True
    except Exception:
        pass

    # sinon : lire le FULL et ne renvoyer qu'un échantillon -> preview
    try:
        from backend.services.gabarit_registry import get_default_dataframe
        df_full = get_default_dataframe(name, ver)
        if isinstance(df_full, pd.DataFrame) and not df_full.empty:
            return df_full.head(20).copy(), True
    except Exception:
        pass
    try:
        from backend.services.dataset_service import get_default_dataframe_for_gabarit
        df_full = get_default_dataframe_for_gabarit(name, ver)
        if isinstance(df_full, pd.DataFrame) and not df_full.empty:
            return df_full.head(20).copy(), True
    except Exception:
        pass

    return None, True

def _compose_with_enrichments(
    u: dict,
    *,
    full: bool,
    bring_all_last: bool = False,  # ramener toutes les colonnes du dernier saut (utile pour tester un script)
    log: bool = False,             # afficher des KPIs (tailles, mode FULL/PREVIEW, taux de match)
) -> tuple[pd.DataFrame | None, bool, str | None]:
    """
    Construit la table base + enrichissements (N sauts) à partir des données par défaut.
    Renvoie (df, complete, err).
      - df: DataFrame composé (ou None si impossible)
      - complete: True si toutes les tables ont été chargées en FULL quand full=True (sinon False)
      - err: message d'erreur explicite si pertinent
    Règles:
      - On normalise les clés avant chaque merge.
      - On transporte les clés intermédiaires (left_key du saut suivant).
      - On évite de ramener des colonnes déjà présentes à gauche.
      - On ne re-sélectionne jamais right_key dans les colonnes rapatriées.
    """
    # Charger la base
    df, is_preview = _load_df_for_gabarit(gname, gver, full=full)
    if df is None:
        if full:
            return None, False, f"FULL demandé mais aucune source n'est définie pour {gname} (v{gver})."
        return None, False, f"Aucune donnée par défaut pour {gname} (v{gver})."

    if log:
        _log_kpi("📦 Source de base", {
            "gabarit": f"{gname} (v{gver})",
            "lignes": len(df),
            "colonnes": len(df.columns),
            "mode": "PREVIEW" if is_preview else "FULL",
        })

    # Si on demandait full et qu'on a une preview, on n'est pas "complete"
    complete = (not is_preview) if full else True

    # --- on mémorise les colonnes réellement ajoutées par les enrichissements
    added_enriched_cols: set[str] = set()

    # Parcours des enrichissements
    for e in (u.get("enrichments") or []):
        path = e.get("path") or []
        if not path:
            continue

        # Enchaîner les sauts
        for i, step in enumerate(path):
            frm, left_key, to, right_key = step  # [from, left_key, to, right_key]

            df_to, is_prev_to = _load_df_for_gabarit(to, "v1", full=full)
            if df_to is None:
                if full:
                    return df, False, f"FULL demandé mais aucune source n'est définie pour {to} (v1)."
                return df, False, f"Aucune donnée par défaut pour {to} (v1)."

            if log:
                _log_kpi(f"🔗 Table d'enrichissement #{i+1} → {to}", {
                    "lignes": len(df_to),
                    "colonnes": len(df_to.columns),
                    "mode": "PREVIEW" if is_prev_to else "FULL",
                    "left_key": left_key,
                    "right_key": right_key,
                })

            # Colonnes à rapatrier pour CE saut (sans dupliquer la clé droite)
            cols_to_fetch: set[str] = set()

            if i + 1 < len(path):
                # on aura besoin de la left_key du saut suivant (présente dans 'to')
                next_left_key = path[i + 1][1]
                if next_left_key and next_left_key in df_to.columns:
                    cols_to_fetch.add(next_left_key)
            else:
                # Dernier saut
                if bring_all_last:
                    # ramener toutes les colonnes de la cible (sauf la clé droite)
                    for c in df_to.columns:
                        if c != right_key:
                            cols_to_fetch.add(c)
                else:
                    # rigoureux : seulement les colonnes sélectionnées pour l'enrichissement
                    for c in (e.get("columns") or []):
                        if c and c in df_to.columns:
                            cols_to_fetch.add(c)

            # NE PAS inclure la clé droite dans les colonnes à rapatrier
            cols_to_fetch.discard(right_key)

            # Éviter de ramener des colonnes déjà présentes dans la table de gauche
            cols_to_add = [c for c in cols_to_fetch if c not in df.columns]

            # Vérifier la présence des clés
            if left_key not in df.columns or right_key not in df_to.columns:
                return df, False, f"Clé manquante pour joindre {frm} → {to} ({left_key} / {right_key})."

            # Normaliser les clés avant merge (zéro-padding SIREN/SIRET, retrait des séparateurs, etc.)
            df[left_key] = _normalize_key(df[left_key], left_key)
            df_to[right_key] = _normalize_key(df_to[right_key], right_key)

            # Sous-ensemble de la droite : [right_key] + colonnes utiles (sans doublons)
            right_cols = [right_key] + cols_to_add
            seen = set()
            right_cols = [c for c in right_cols if not (c in seen or seen.add(c))]
            right_subset = df_to[right_cols].drop_duplicates()

            if log:
                _log_kpi(f"🧮 Jointure {frm} → {to}", {
                    "left rows": len(df),
                    "right rows (subset)": len(right_subset),
                    "cols rapatriées": ", ".join(cols_to_add) if cols_to_add else "—",
                })

            # MERGE
            df = _debug_merge(
                df,
                right_subset,
                how="left",
                left_on=left_key,
                right_on=right_key,
                tag=f"{frm}->{to}"
            )

            # On ne garde pas la clé de droite après la jointure
            if right_key in df.columns:
                df.drop(columns=[right_key], inplace=True)


            # mémoriser les colonnes effectivement ajoutées par les enrichissements
            added_enriched_cols.update([c for c in cols_to_add if c in df.columns])

            # KPI après merge (compte seulement sur les colonnes effectivement présentes)
            if log:
                present_after = [c for c in cols_to_add if c in df.columns]
                after_non_na = df[present_after].notna().any(axis=1).sum() if present_after else 0
                total = len(df)
                rate = f"{(after_non_na/total*100):.1f}%" if total else "0%"
                _log_kpi("📊 Résultat jointure", {
                    "lignes totales": total,
                    "lignes avec enrichissement (≥1 col)": after_non_na,
                    "taux de match (approx)": rate,
                })

            # si on demandait FULL mais que cette table d'enrichissement est en preview → complete = False
            if full and is_prev_to:
                complete = False

    # === Filtrage final des colonnes enrichies ===
    # On garde toutes les colonnes de base + uniquement les colonnes d'enrichissement sélectionnées dans l'UI.
    selected_enriched = set()
    for e in (u.get("enrichments") or []):
        selected_enriched.update([c for c in (e.get("columns") or []) if c])

    # Colonnes enrichies qu'on retire (car non sélectionnées)
    to_drop = [c for c in added_enriched_cols if c not in selected_enriched and c in df.columns]
    if to_drop:
        df.drop(columns=to_drop, inplace=True, errors="ignore")
        if log:
            _log_kpi("🧹 Nettoyage colonnes enrichies non sélectionnées", {
                "drop": ", ".join(to_drop)
            })

    return df, complete, None


# -- Helper d'application des méthodes sur un DataFrame --
def _apply_methods(df: pd.DataFrame, only: list[str] | None = None) -> pd.DataFrame:
    """
    Applique les méthodes du gabarit courant dans l'ordre, éventuellement filtrées par 'only' (liste de noms).
    Requiert backend.services.gabarit_registry.list_methods_for_gabarit et method_executor.apply_method.
    """
    try:
        from backend.services.gabarit_registry import list_methods_for_gabarit
    except Exception:
        return df

    try:
        from backend.services.method_executor import apply_method as _apply_method
    except Exception:
        # Fallback : pas d'exécution si le moteur n'est pas dispo
        return df

    gname = u.get("gabarit_name")
    gver  = u.get("gabarit_version", "v1")
    methods = list_methods_for_gabarit(gname, gver) or []
    methods = sorted(methods, key=lambda m: m.get("order", 1))
    names_filter = set([n.strip() for n in (only or []) if n and str(n).strip()])

    cur = df.copy()
    for m in methods:
        if names_filter and m.get("name") not in names_filter:
            continue
        # paramètres par défaut (si schema présent)
        schema = m.get("param_schema") or []
        pvals = {}
        for spec in schema:
            nm = spec.get("name")
            dv = spec.get("default")
            if nm:
                pvals[nm] = dv
        try:
            cur = _apply_method(cur, m, pvals)
        except Exception:
            # on continue en cas d'échec d'une méthode spécifique
            pass
    return cur

def _compose_until_overlay_and_methods(u: dict, *, full: bool) -> tuple[pd.DataFrame | None, str | None]:
    """
    Étape de prévisualisation 'Script Python' :
    Base + Enrichissements + Méthodes → Script
    (pas d’ordre/exclusions/renommages).
    """
    # 1) base + enrichissements
    df, complete, err = _compose_with_enrichments(u, full=full, bring_all_last=True, log=True)
    if df is None:
        return None, err or "Aucune donnée de départ."

    # 2) méthodes (sélectionnées)
    only_selected = list(u.get("methods") or [])
    try:
        df = _apply_methods(df, only=only_selected)  # type: ignore[call-arg]
    except TypeError:
        df = _apply_methods(df)

    # 3) script utilisateur (overlay)
    code = (u.get("overlay_python") or "").strip()
    if code:
        df2, err2 = _apply_overlay(df, code)
        if err2 is None and isinstance(df2, pd.DataFrame):
            df = df2
        else:
            return None, err2

    return df, None


def _resolve_effective_columns_for_adjustment(u: dict) -> list[str]:
    """
    Colonnes proposées dans l’onglet 'Ajustement' :
    on exécute Base + Enrich. + Méthodes → Script (en mode PREVIEW si possible),
    puis on renvoie df.columns.
    """
    # On essaie en PREVIEW pour être léger ; si ça échoue on bascule en FULL.
    df, err = _compose_until_overlay_and_methods(u, full=False)
    if df is None:
        df, err = _compose_until_overlay_and_methods(u, full=True)
    if df is None:
        # fallback minimal si vraiment rien ne marche
        try:
            from backend.services.template_service import TemplateService
            ts = TemplateService()
            return ts.resolve_usage_expected_columns(u.get("template_id"), u.get("gabarit_name"), u.get("gabarit_version", "v1"))
        except Exception:
            return []
    return list(df.columns)

def _get_current_usage_for_preview(template_id: int, gname: str, gver: str, usage_fallback: dict) -> dict:
    """
    Recharge le 'usage' persistant (DB) pour la preview finale.
    Puis superpose les éventuels ajustements en session (si non sauvés).
    """
    try:
        from backend.services.template_service import TemplateService
        from backend.services.database_service import DatabaseService
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            fresh = ts.get_gabarit_usage(template_id, gname, gver) or {}
    except Exception:
        fresh = {}

    # Si rien en DB, on retombe sur l'usage courant de la page
    u = dict(usage_fallback)
    u.update(fresh)  # la DB prime sur l'ancien 'usage' si présent

    # Superposer les ajustements encore en session (non sauvegardés)
    ord_ss = st.session_state.get("adj_final_order")
    if ord_ss:
        u["final_order"] = list(ord_ss)

    exc_ss = st.session_state.get("adj_final_excludes")
    if exc_ss is not None:
        u["final_excludes"] = list(exc_ss)

    ren_ss = st.session_state.get("adj_final_renames")
    if ren_ss is not None:
        u["final_renames"] = dict(ren_ss)

    # Idem si tu as ajouté un tri final dans l’UI plus tard :
    if "final_sort" in st.session_state:
        u["final_sort"] = st.session_state["final_sort"]

    return u

def _apply_overlay(df: pd.DataFrame, code: str, params: dict | None = None) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Exécute le script Python overlay sur le DataFrame.

    Variables exposées au script :
      - df : DataFrame (copie)
      - pd : pandas
      - params : dict des paramètres
      - chaque paramètre exporté comme variable si son nom est un identifiant Python (ex: Secteur)
    """
    if not code or not code.strip():
        return df, None

    try:
        local_vars = {"pd": pd, "df": df.copy(), "params": params or {}}

        # ➕ rendre chaque paramètre accessible directement (ex: Secteur)
        for k, v in (params or {}).items():
            if isinstance(k, str) and k not in {"pd", "df", "params"} and k.isidentifier():
                local_vars[k] = v

        exec(code, {}, local_vars)
        result = local_vars.get("df", None)
        if isinstance(result, pd.DataFrame):
            return result, None
        return None, "Le script n'a pas produit de DataFrame 'df'."
    except Exception as ex:
        import traceback
        return None, f"{type(ex).__name__}: {ex}\n{traceback.format_exc()}"


def _resolve_effective_columns_for_adjustment(u: dict) -> list[str]:
    """
    Colonnes à proposer dans l'onglet Ajustement :
    - Si overlay_python est exécutable et que les données par défaut sont complètes → colonnes du df overlayé.
    - Sinon, colonnes 'par défaut' (base + enrich + sorties de méthodes).
    """
    code = (u.get("overlay_python") or "").strip()
    df, complete, _ = _compose_with_enrichments(u, full=False)
    if code and complete and isinstance(df, pd.DataFrame):
        df = _apply_methods(df)
        df2, err = _apply_overlay(df, code)
        if err is None and isinstance(df2, pd.DataFrame):
            return list(df2.columns)

    # fallback (base + enrich + sorties méthodes, selon registry)
    g = get_gabarit(gname, gver)
    base = [c.name for c in (g.columns or [])]
    cols = u.get("columns_enabled") or base[:]
    for e in (u.get("enrichments") or []):
        for c in (e.get("columns") or []):
            if c not in cols:
                cols.append(c)
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

def _compose_full_pipeline(u: dict, *, full: bool, params: dict | None = None) -> tuple[pd.DataFrame | None, str | None]:
    """
    Pipeline final pour la PRÉVISUALISATION :
    Base + Enrichissements + Méthodes → Script → Renommages → TRI → Ordre/Exclusions.
    """
    # 1) base + enrichissements
    df, complete, err = _compose_with_enrichments(u, full=full, bring_all_last=True, log=True)
    if df is None:
        return None, err or "Aucune donnée de départ."

    # 2) méthodes sélectionnées
    only_selected = list(u.get("methods") or [])
    try:
        df = _apply_methods(df, only=only_selected)  # type: ignore[call-arg]
    except TypeError:
        df = _apply_methods(df)

    # 3) script utilisateur (overlay)
    code = (u.get("overlay_python") or "").strip()
    if code:
        df2, err2 = _apply_overlay(df, code, params=params)
        if err2 is None and isinstance(df2, pd.DataFrame):
            df = df2
        else:
            return None, err2

    # 4) renommages de colonnes
    ren: dict[str, str] = u.get("final_renames") or {}
    if ren:
        safe_map = {k: v for k, v in ren.items() if k in df.columns and v and v != k}
        if safe_map:
            df = df.rename(columns=safe_map)

    # 5) TRI DES LIGNES (optionnel)
    # u["final_sort"] peut être une liste de dicts: [{"col":"Nom", "asc": True}, ...]
    sort_rules = u.get("final_sort") or []
    if isinstance(sort_rules, list) and sort_rules:
        by: list[str] = []
        ascending: list[bool] = []
        for r in sort_rules:
            if not isinstance(r, dict):
                continue
            col = (r.get("col") or "").strip()
            if not col:
                continue
            # si la colonne a été renommée plus haut, on travaille sur son NOM ACTUEL
            col_now = ren.get(col, col)
            if col_now in df.columns:
                by.append(col_now)
                ascending.append(bool(r.get("asc", True)))
        if by:
            # tri stable pour préserver l'ordre relatif si égalité
            df = df.sort_values(by=by, ascending=ascending, kind="mergesort", ignore_index=True)

    # 5-bis) Sécuriser les noms de colonnes (supprimer les doublons de noms)
    # Après renames + tri, il peut rester des noms identiques (ex. 'EBE' provenant de 2 sources).
    # On garde la première occurrence et on supprime les suivantes pour éviter l'erreur PyArrow.
    if df.columns.duplicated().any():
        dup_names = list(df.columns[df.columns.duplicated(keep=False)])
        _log_kpi("⚠️ Noms de colonnes dupliqués détectés (suppression des doublons, keep=first)", {
            "doublons": ", ".join(map(str, dup_names[:30])) + (" …" if len(dup_names) > 30 else "")
        })
        df = df.loc[:, ~df.columns.duplicated(keep="first")]


    # 6) Ordre / exclusions (après renommages) — VERSION STRICTE
    src_order = u.get("final_order") or []
    src_excl  = set(u.get("final_excludes") or [])

    # mapper l’ordre via les renommages
    mapped_order = []
    seen = set()
    for c in src_order:
        cc = ren.get(c, c)
        if cc in df.columns and cc not in seen:
            mapped_order.append(cc)
            seen.add(cc)

    # exclusions : exclure anciens noms et nouveaux noms
    excl_names = set()
    for c in src_excl:
        excl_names.add(c)
        rc = ren.get(c)
        if rc:
            excl_names.add(rc)

    # STRICT : n'afficher que les colonnes sélectionnées, dans l’ordre défini
    if mapped_order:
        final_cols = [c for c in mapped_order if c in df.columns and c not in excl_names]
    else:
        # si aucun ordre n’est défini, on garde tout (comportement “neutre”)
        final_cols = [c for c in df.columns if c not in excl_names]

    _log_kpi("📦 Sortie pipeline (finale)", {
        "lignes": len(df),
        "colonnes": len(final_cols),
        "complete(full)": complete
    })
    return df[final_cols], None


def _sample_value(col: str) -> str:
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



def _normalize_key(series: pd.Series, colname: str) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(r"[^0-9A-Za-z]", "", regex=True)
    low = (colname or "").lower()
    if any(k in low for k in ["siren", "siret"]):
        digits = s.str.replace(r"[^0-9]", "", regex=True)
        if "siret" in low:
            s = digits.str.zfill(14)
        else:
            s = digits.str.zfill(9)
    return s


def _log_kpi(title: str, kv: dict):
    with st.expander(title, expanded=False):
        for k, v in kv.items():
            st.caption(f"• {k}: {v}")



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
                    # Affiche "Type" + la colonne source si le paramètre liste provient d'une colonne
                    _src_hint = ""
                    try:
                        _mode = getattr(param, "options_mode", "none")
                        # compat : type peut être "select" ou "liste" ; mode peut être "from_column" ou "column"
                        if getattr(param, "type", "") in ("select", "liste") and _mode in ("from_column", "column"):
                            _src = getattr(param, "options_source", None) or {}
                            _col = _src.get("column") or getattr(param, "source_column", None)
                            if _col:
                                _src_hint = f" | Colonne source : `df['{_col}']`"
                    except Exception:
                        pass
                    st.caption(f"Type : {param.type}{_src_hint}")

                    if param.default:
                        st.caption(f"Défaut : `{param.default}`")
                    # options + méta (si calculables)
                    if getattr(param, "type", None) == "select" and getattr(param, "options_mode", "none") != "none":
                        # 1) Afficher la colonne source si l'option est construite depuis une colonne
                        if getattr(param, "options_mode", None) == "column":
                            src_col = getattr(param, "source_column", None)
                            if src_col:
                                st.caption(f"📎 Construit depuis la colonne : `{src_col}`")

                        # 2) Afficher une aperçu des options (quelle que soit l'origine)
                        try:
                            options = ParameterService.resolve_parameter_options(param)
                            if options:
                                st.caption(f"Options : {', '.join(map(str, options[:3]))}{'...' if len(options) > 3 else ''}")
                        except Exception:
                            # on n'échoue pas l'affichage si la résolution d'options plante
                            pass


        st.markdown("---")
        st.info("💡 Exemples : `df = df[df['Marque'] == Sous_Marque]` **ou** `df = df[df['Marque'] == params['Sous_Marque']]`")
    else:
        st.info("Aucun paramètre défini pour ce template")

    
    st.markdown("---")
    
    # ✅ CODE EDITOR AVEC PERSISTANCE TOTALE (avec persistance + form)
    st.markdown("### 🔧 Script Python de transformation")
    st.caption("💡 Variables : `df` (DataFrame), `pd` (pandas), `params` (dict), et chaque paramètre accessible par son nom (ex : `Secteur`).")
    st.caption("⚠️ Le script doit réassigner `df` (ex : `df = df[df['Col'] == Secteur]`).")

    # ✅ Clé unique par table (persistance)
    persist_key = f"code_persist_{template_id}_{gname}_{gver}_{sheet}_{table}"

    # ✅ Initialisation depuis DB si première fois
    if persist_key not in st.session_state:
        st.session_state[persist_key] = (usage.get("overlay_python") or "").strip()

    # ✅ Form => synchro garantie (pas d’effacement au 1er clic)
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
            response_mode=["submit", "blur"],  # <-- capture AVANT le rerun
        )

        # Extraction robuste -> met à jour le buffer AVANT de traiter les clics
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

        c1, c2 = st.columns([1, 1], gap="large")
        with c1:
            do_preview = st.form_submit_button("🧪 Prévisualiser le script", use_container_width=True)
        with c2:
            do_save = st.form_submit_button("💾 Enregistrer le script", type="primary", use_container_width=True)

    code = st.session_state[persist_key]

    # === Actions après le form ===
    if do_save:
        with DatabaseService.get_session() as db:
            ts = TemplateService(db)
            cfg2 = ts.get_config(template_id)
            usages2 = cfg2.get("gabarit_usages", []) or []
            # mise à jour de l'usage ciblé
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

    if do_preview:
        # Params avec valeurs par défaut
        params_dict = {}
        for param in template_config.parameters:
            params_dict[param.name] = ParameterService.get_default_value(param)

        # Usage modifié temporairement
        usage_test = dict(usage)
        usage_test["overlay_python"] = code

        df_final, error = _compose_full_pipeline(usage_test, full=True, params=params_dict)

        if error:
            st.error(f"❌ Erreur :\n```\n{error}\n```")
        elif df_final is None or df_final.empty:
            st.warning("Aucun résultat")
        else:
            st.success(f"✅ {len(df_final)} lignes, {len(df_final.columns)} colonnes")
            st.dataframe(df_final.head(20), use_container_width=True, hide_index=True)


    with st.container():
        st.markdown("**Colonnes disponibles**")
        cols_eff = _resolve_effective_columns_for_adjustment(usage)
        if cols_eff:
            for col in cols_eff[:10]:
                st.caption(f"• {col}")
            if len(cols_eff) > 10:
                st.caption(f"... et {len(cols_eff)-10} autres")
        else:
            st.caption("—")


# ============================ Onglet 2 — AJUSTEMENT ============================
with tab_adjust:
    # Colonnes “effectives” proposées à l’ajustement (base + enrich + sorties méthodes,
    # ou résultat overlay si exécutable sur preview)
    default_cols = _resolve_effective_columns_for_adjustment(usage)

    # --- état local lié à la table sélectionnée ---
    usage_key = f"{gname}|{gver}|{sheet}|{table}"
    if st.session_state.get("_adj_key") != usage_key:
        st.session_state["_adj_key"] = usage_key
        st.session_state["adj_final_order"] = list(usage.get("final_order") or default_cols[:])
        st.session_state["adj_final_excludes"] = set(usage.get("final_excludes") or [])
        st.session_state["adj_final_renames"] = dict(usage.get("final_renames") or {})

    # synchronisation avec la structure effective
    final_order = [c for c in st.session_state["adj_final_order"] if c in default_cols] + \
                  [c for c in default_cols if c not in st.session_state["adj_final_order"]]
    final_excludes = set([c for c in st.session_state["adj_final_excludes"] if c in default_cols])
    final_renames = dict(st.session_state["adj_final_renames"])

    # --- métadonnées (source/type) pour affichage ---
    g = get_gabarit(gname, gver)
    type_map = {c.name: (c.type or "text") for c in (g.columns or [])}
    source_map = {c: "gabarit" for c in (usage.get("columns_enabled") or [c.name for c in (g.columns or [])])}

    for e in (usage.get("enrichments") or []):
        if not e.get("path"):
            continue
        last = e["path"][-1]
        tgt_name = last[2]
        tgt_g = get_gabarit(tgt_name, "v1")
        for c in (e.get("columns") or []):
            source_map[c] = f"enrich:{tgt_name}"
            if c not in type_map:
                type_map[c] = next((col.type for col in (tgt_g.columns or []) if col.name == c), "text")

    for m in (usage.get("methods") or []):
        source_map[m] = f"method:{m}"
        type_map.setdefault(m, "unknown")

    # --- entête ---
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

    def _muted_html(s: str) -> str:
        return f"<span style='color:#9aa0a6'>{s}</span>"

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
            txt = _sample_value(col)
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
        if st.button("💾 Enregistrer l’ajustement", type="primary", use_container_width=True, key=f"btn_save_adj_{usage_key}"):
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
        if st.button("🔁 Réinitialiser", use_container_width=True, key=f"btn_reset_adj_{usage_key}"):
            eff = _resolve_effective_columns_for_adjustment(usage)
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
            st.info("Réinitialisé")
            st.rerun()


# ============================ Onglet 3 — PRÉVISUALISATION ============================

with tab_preview:
    st.caption(f"Feuille : **{sheet}** • Table : **{table}**")
    st.markdown("---")

    params_dict = {p.name: ParameterService.get_default_value(p) for p in template_config.parameters}
    if params_dict:
        st.caption("Paramètres (valeurs par défaut) : " + ", ".join(params_dict.keys()))



    # Prévisualisation basée sur le pipeline "rapide" (full=False)
    usage_preview = _get_current_usage_for_preview(template_id, gname, gver, usage)

    df_prev, err = _compose_full_pipeline(usage_preview, full=True, params=params_dict)

    if err:
        st.error(f"Erreur pipeline :\n\n{err}")
    elif df_prev is None or df_prev.empty:
        st.info("Pipeline exécuté mais aucun résultat affichable.")
    else:
        st.success("Résultat FINAL — 20 premières lignes :")
        st.dataframe(df_prev.head(20), use_container_width=True, hide_index=True)
