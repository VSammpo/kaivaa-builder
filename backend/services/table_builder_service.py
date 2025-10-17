# backend/services/table_builder_service.py
"""
Service centralisé de construction de tables à partir des usages de gabarits.
Utilisé à la fois pour la prévisualisation et l'injection dans les livrables.
"""
from __future__ import annotations
import pandas as pd
from typing import Optional, Tuple
from loguru import logger
from collections import deque
import difflib
import re
import unicodedata

# Imports des services nécessaires
from backend.services.gabarit_registry import (
    get_gabarit, 
    get_relations, 
    get_default_dataframe,
    get_default_preview,
    list_methods_for_gabarit
)
from backend.services.dataset_service import get_default_dataframe_for_gabarit

# === DEBUG ENRICHISSEMENTS ===
DEBUG_ENRICH = True
def _normalize_to_usage(transfo_or_usage: dict) -> dict:
    """
    Convertit une transformation OU un usage en format usage unifié.
    ⚠️ IMPORTANT : on conserve les clés techniques (ex: "_source_df") transmises par l'appelant.
    Politique de MERGE quand une transformation est référencée dans l'usage :
      - On charge la transformation (gabarit_base, columns_enabled, enrichments, methods, overlay, final_*).
      - On applique une surcouche éventuelle de l'usage :
          overlay_python = transfo.overlay + "\n\n" + usage.overlay (si usage.overlay non vide)
          final_order    = si usage.final_order est fourni → sous-ensemble/ordre de la transfo
          final_excludes = union (transfo ∪ usage)
          final_renames  = transfo puis update avec usage
          final_sort     = usage.final_sort si fourni, sinon transfo.final_sort
      - columns_enabled / methods / enrichments : on garde ceux de la transformation (on ignore ceux fournis côté usage).
    """
    from backend.services.transformation_service import get_transformation

    if not transfo_or_usage or not isinstance(transfo_or_usage, dict):
        return {}

    u = dict(transfo_or_usage)

    # --- Helper pour recopier les clés techniques (ex: "_source_df") ---
    def _carry_tech_keys(src: dict, dst: dict):
        for k, v in src.items():
            # on recopie toutes les clés "techniques" (commençant par "_")
            # et, plus largement, toute clé inconnue non couverte par le schéma standard
            if k.startswith("_"):
                dst[k] = v

    # Cas A — Usage SANS transformation : le retourner tel quel (mais normalisé minimalement)
    transfo_name = (u.get("transformation_name") or "").strip()
    if not transfo_name:
        # Si c'est une transformation "pleine" (clé 'gabarit_base'), normaliser en usage
        if "gabarit_base" in u and "gabarit_name" not in u:
            gabarit_base = u.get("gabarit_base", {})
            base = {
                "gabarit_name": gabarit_base.get("name", ""),
                "gabarit_version": gabarit_base.get("version", "v1"),
                "columns_enabled": list(u.get("columns_enabled") or []),
                "enrichments": list(u.get("enrichments") or []),
                "methods": list(u.get("methods") or []),
                "overlay_python": (u.get("overlay_python") or "").strip(),
                "final_order": list(u.get("final_order") or []),
                "final_excludes": list(u.get("final_excludes") or []),
                "final_renames": dict(u.get("final_renames") or {}),
                "final_sort": list(u.get("final_sort") or []),
                "excel_target": dict(u.get("excel_target") or {}),
            }
            _carry_tech_keys(u, base)
            return base

        # Sinon, usage "brut" : recopier les clés techniques et renvoyer
        base = dict(u)
        _carry_tech_keys(u, base)
        return base

    # Cas B — Usage AVEC transformation : fusionner transformation + surcouche usage
    tver = (u.get("transformation_version") or "v1").strip()
    T = get_transformation(transfo_name, tver) or {}

    base = {
        "gabarit_name": (T.get("gabarit_base") or {}).get("name", ""),
        "gabarit_version": (T.get("gabarit_base") or {}).get("version", "v1"),
        "columns_enabled": list(T.get("columns_enabled") or []),
        "enrichments": list(T.get("enrichments") or []),
        "methods": list(T.get("methods") or []),
        "overlay_python": (T.get("overlay_python") or "").strip(),
        "final_order": list(T.get("final_order") or []),
        "final_excludes": list(T.get("final_excludes") or []),
        "final_renames": dict(T.get("final_renames") or {}),
        "final_sort": list(T.get("final_sort") or []),
        "excel_target": dict(u.get("excel_target") or {}),
    }

    # Surcouche usage : overlay
    u_overlay = (u.get("overlay_python") or "").strip()
    if u_overlay:
        base["overlay_python"] = (base["overlay_python"] + "\n\n" + u_overlay) if base["overlay_python"] else u_overlay

    # Surcouche usage : final_excludes (union)
    u_excl = [str(c).strip() for c in (u.get("final_excludes") or []) if str(c).strip()]
    if u_excl:
        base["final_excludes"] = list(dict.fromkeys(list(base.get("final_excludes") or []) + u_excl))

    # Surcouche usage : final_renames (update)
    u_ren = dict(u.get("final_renames") or {})
    if u_ren:
        ren = dict(base.get("final_renames") or {})
        for k, v in u_ren.items():
            ks = str(k).strip()
            vs = str(v).strip()
            if ks and vs:
                ren[ks] = vs
        base["final_renames"] = ren

    # Surcouche usage : final_sort (priorité usage si fourni)
    if u.get("final_sort") is not None:
        base["final_sort"] = list(u.get("final_sort") or [])

    # Surcouche usage : final_order (sous-ensemble/ordre)
    u_order = list(u.get("final_order") or [])
    if u_order:
        t_order = list(base.get("final_order") or [])
        if t_order:
            t_set = set(t_order)
            base["final_order"] = [c for c in u_order if c in t_set]
        else:
            base["final_order"] = [c for c in u_order if c]

    # 🔴 CRITIQUE : conserver les clés techniques (dont "_source_df")
    _carry_tech_keys(u, base)

    return base


def build_table_from_transformation(
    transformation_name: str,
    transformation_version: str = "v1",
    *,
    full: bool = True,
    log_kpis: bool = False,
    params: dict | None = None
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Construit une table à partir d'une transformation nommée.
    Wrapper convenience autour de build_table_from_usage.
    """
    from backend.services.transformation_service import get_transformation
    
    logger.info(f"Construction depuis transformation : {transformation_name} v{transformation_version}")
    
    # Charger la transformation
    transformation = get_transformation(transformation_name, transformation_version)
    if not transformation:
        return None, f"Transformation '{transformation_name}' v{transformation_version} introuvable"
    
    # Utiliser build_table_from_usage qui va normaliser
    return build_table_from_usage(
        transformation,
        full=full,
        log_kpis=log_kpis,
        params=params
    )

def _debug_merge(left_df, right_df, *, how: str, left_on: str, right_on: str, tag: str = ""):
    """Remplace un pd.merge pour tracer ce qui se passe lors des enrichissements."""
    import logging
    logger = logging.getLogger("kaivaa.enrich")
    try:
        l_nonnull = left_df[left_on].notna().sum() if left_on in left_df.columns else 0
        r_nonnull = right_df[right_on].notna().sum() if right_on in right_df.columns else 0
        l_dtype = str(left_df[left_on].dtype) if left_on in left_df.columns else "?"
        r_dtype = str(right_df[right_on].dtype) if right_on in right_df.columns else "?"

        if left_on in left_df.columns:
            l_sample = list(map(str, left_df[left_on].dropna().astype(str).head(5).unique()))
        else:
            l_sample = []
        if right_on in right_df.columns:
            r_sample = list(map(str, right_df[right_on].dropna().astype(str).head(5).unique()))
        else:
            r_sample = []

        if DEBUG_ENRICH:
            logger.info(
                f"[ENRICH{(':'+tag) if tag else ''}] how={how} "
                f"left({len(left_df)} rows) on={left_on}[{l_dtype}] nnz={l_nonnull} sample={l_sample} "
                f"right({len(right_df)} rows) on={right_on}[{r_dtype}] nnz={r_nonnull} sample={r_sample}"
            )
        out = left_df.merge(
            right_df,
            how=how,
            left_on=left_on,
            right_on=right_on,
            suffixes=("_L", "_R"),
        )
        if DEBUG_ENRICH:
            logger.info(f"[ENRICH{(':'+tag) if tag else ''}] result rows={len(out)} cols={len(out.columns)}")
        return out
    except Exception as e:
        if DEBUG_ENRICH:
            logger.exception(f"[ENRICH{(':'+tag) if tag else ''}] merge failed: {e}")
        raise


def _normalize_key(series: pd.Series, colname: str) -> pd.Series:
    """Normalise les clés pour les jointures (SIREN/SIRET, etc.)"""
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


def _load_df_for_gabarit(name: str, ver: str, full: bool = False) -> Tuple[Optional[pd.DataFrame], bool]:
    """
    Charge le DataFrame d'un gabarit (full ou preview).
    Retourne (df, is_preview).
    """
    if full:
        # Tenter le FULL
        try:
            df_full = get_default_dataframe(name, ver)
            if isinstance(df_full, pd.DataFrame) and not df_full.empty:
                return df_full, False
        except Exception:
            pass
        try:
            df_full = get_default_dataframe_for_gabarit(name, ver)
            if isinstance(df_full, pd.DataFrame) and not df_full.empty:
                return df_full, False
        except Exception:
            pass
        return None, False
    
    # Preview
    try:
        prev = get_default_preview(name, ver) or {}
        rows, cols = prev.get("rows") or [], prev.get("columns") or []
        if rows and cols:
            return pd.DataFrame(rows, columns=cols), True
    except Exception:
        pass
    
    # Fallback : lire FULL et ne renvoyer qu'un échantillon
    try:
        df_full = get_default_dataframe(name, ver)
        if isinstance(df_full, pd.DataFrame) and not df_full.empty:
            return df_full.head(20).copy(), True
    except Exception:
        pass
    
    return None, True


def _apply_methods(df: pd.DataFrame, gabarit_name: str, gabarit_version: str, 
                   only: list[str] | None = None) -> pd.DataFrame:
    """Applique les méthodes du gabarit dans l'ordre"""
    try:
        from backend.services.method_executor import apply_method
    except Exception:
        return df
    
    methods = list_methods_for_gabarit(gabarit_name, gabarit_version) or []
    methods = sorted(methods, key=lambda m: m.get("order", 1))
    names_filter = set([n.strip() for n in (only or []) if n and str(n).strip()])
    
    cur = df.copy()
    for m in methods:
        if names_filter and m.get("name") not in names_filter:
            continue
        schema = m.get("param_schema") or []
        pvals = {}
        for spec in schema:
            nm = spec.get("name")
            dv = spec.get("default")
            if nm:
                pvals[nm] = dv
        try:
            cur = apply_method(cur, m, pvals)
        except Exception as e:
            logger.warning(f"Erreur lors de l'application de la méthode {m.get('name')}: {e}")
    
    return cur



def _apply_overlay(df: pd.DataFrame, code: str, params: dict | None = None):
    """
    Exécute un script d'usage (overlay) sur un DataFrame.
    - Injecte pd, df (copie), params dans un namespace unique (globals == locals)
    - Expose chaque paramètre sous plusieurs alias (k, k.lower(), k.upper(), Capitalized)
    - Le script doit laisser le résultat final dans 'df'
    """
    if not code or not code.strip():
        return df, None
    try:
        env: dict = {
            "pd": pd,
            "df": df.copy(),
            "params": params or {},
            "__builtins__": __builtins__,
        }
        # Alias de paramètres utilisables dans le script
        for k, v in (params or {}).items():
            if not isinstance(k, str):
                continue
            if k in {"pd", "df", "params"}:
                continue
            aliases = {k, k.lower(), k.upper(), (k[:1].upper() + k[1:])}
            for alias in aliases:
                if isinstance(alias, str) and alias.isidentifier():
                    env.setdefault(alias, v)

        exec(code, env, env)  # 👈 un seul namespace
        result = env.get("df", None)
        if isinstance(result, pd.DataFrame):
            return result, None
        return None, "Le script n'a pas produit de DataFrame 'df'."
    except Exception as ex:
        import traceback
        return None, f"{type(ex).__name__}: {ex}\n{traceback.format_exc()}"


def build_table_from_usage(
    usage: dict,
    *,
    full: bool = True,
    log_kpis: bool = False,
    params: dict | None = None
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Construit le DataFrame final à partir d'un usage de gabarit.
    ORDRE STRICT : Base → Enrichissements → Méthodes → Script → Renommages → Tri → Ordre/Exclusions
    
    Args:
        usage: Configuration de l'usage
        full: Charger données complètes (True) ou preview (False)
        log_kpis: Afficher les KPIs de construction
        params: Paramètres du template pour injection dans le script
    """
    
    # ✅ NOUVEAU : Normaliser transformation → usage
    usage = _normalize_to_usage(usage)
    if not usage:
        return None, "Usage vide ou invalide après normalisation"
    
    gabarit_name = usage.get("gabarit_name", "")
    gabarit_version = usage.get("gabarit_version", "v1")
    
    if log_kpis:
        logger.info(f"[PIPELINE] 🚀 DÉBUT pour {gabarit_name} (v{gabarit_version})")
        logger.info(f"[PIPELINE]   • Mode: {'FULL' if full else 'PREVIEW'}")
        logger.info(f"[PIPELINE]   • Params: {list(params.keys()) if params else []}")
    
    # ============================================================================
    # ÉTAPE 1 : CHARGEMENT BASE
    # ============================================================================
    if "_source_df" in usage:
        df = usage["_source_df"].copy()
        is_preview = False
        complete = True
        if log_kpis:
            logger.info(f"[PIPELINE] 📦 ÉTAPE 1 : Base fournie - {len(df)} lignes × {len(df.columns)} colonnes")
    else:
        df, is_preview = _load_df_for_gabarit(gabarit_name, gabarit_version, full=full)
        if df is None:
            if full:
                return None, f"FULL demandé mais aucune source n'est définie pour {gabarit_name} (v{gabarit_version})."
            return None, f"Aucune donnée par défaut pour {gabarit_name} ({gabarit_version})."

        
        complete = (not is_preview) if full else True
        
        if log_kpis:
            logger.info(f"[PIPELINE] 📦 ÉTAPE 1 : Base chargée - {len(df)} lignes × {len(df.columns)} colonnes")
            logger.info(f"[PIPELINE]   • Mode: {'PREVIEW' if is_preview else 'FULL'}")
            logger.info(f"[PIPELINE]   • Colonnes: {list(df.columns)[:10]}{'...' if len(df.columns) > 10 else ''}")
    
    # Mémoriser les colonnes ajoutées par enrichissements
    cols_before_enrich = set(df.columns)
    added_enriched_cols: set[str] = set()
    
    # ============================================================================
    # ÉTAPE 2 : ENRICHISSEMENTS
    # ============================================================================
    enrichments = usage.get("enrichments") or []
    if enrichments and log_kpis:
        logger.info(f"[PIPELINE] 🔗 ÉTAPE 2 : {len(enrichments)} enrichissement(s)")
    
    for idx_enrich, e in enumerate(enrichments):
        path = e.get("path") or []
        if not path:
            continue
        
        if log_kpis:
            logger.info(f"[PIPELINE]   • Enrichissement #{idx_enrich+1} : {len(path)} saut(s)")
        
        for i, step in enumerate(path):
            frm, left_key, to, right_key = step
            
            df_to, is_prev_to = _load_df_for_gabarit(to, "v1", full=full)
            if df_to is None:
                if full:
                    return df, f"FULL demandé mais aucune source n'est définie pour {to} (v1)."
                return df, f"Aucune donnée par défaut pour {to} (v1)."
            
            # Colonnes à rapatrier
            cols_to_fetch: set[str] = set()
            
            if i + 1 < len(path):
                next_left_key = path[i + 1][1]
                if next_left_key and next_left_key in df_to.columns:
                    cols_to_fetch.add(next_left_key)
            else:
                for c in (e.get("columns") or []):
                    if c and c in df_to.columns:
                        cols_to_fetch.add(c)
            
            cols_to_fetch.discard(right_key)
            cols_to_add = [c for c in cols_to_fetch if c not in df.columns]
            
            # Vérifier les clés
            if left_key not in df.columns or right_key not in df_to.columns:
                return df, f"Clé manquante pour joindre {frm} → {to} ({left_key} / {right_key})."
            
            # Normaliser les clés
            df[left_key] = _normalize_key(df[left_key], left_key)
            df_to[right_key] = _normalize_key(df_to[right_key], right_key)
            
            # Sous-ensemble
            right_cols = [right_key] + cols_to_add
            seen = set()
            right_cols = [c for c in right_cols if not (c in seen or seen.add(c))]
            right_subset = df_to[right_cols].drop_duplicates()
            
            if log_kpis:
                logger.info(f"[PIPELINE]     ↳ Saut {i+1}/{len(path)} : {frm}[{left_key}] → {to}[{right_key}]")
                logger.info(f"[PIPELINE]       Colonnes ajoutées: {cols_to_add if cols_to_add else 'aucune'}")
            
            # MERGE
            df = _debug_merge(
                df,
                right_subset,
                how="left",
                left_on=left_key,
                right_on=right_key,
                tag=f"{frm}->{to}"
            )
            
            if right_key in df.columns:
                df.drop(columns=[right_key], inplace=True)
            
            added_enriched_cols.update([c for c in cols_to_add if c in df.columns])
            
            if full and is_prev_to:
                complete = False
    
    # Filtrage des colonnes enrichies non sélectionnées
    selected_enriched = set()
    for e in enrichments:
        selected_enriched.update([c for c in (e.get("columns") or []) if c])
    
    to_drop = [c for c in added_enriched_cols if c not in selected_enriched and c in df.columns]
    if to_drop:
        df.drop(columns=to_drop, inplace=True, errors="ignore")
        if log_kpis:
            logger.info(f"[PIPELINE] 🧹 Nettoyage : {len(to_drop)} colonne(s) enrichie(s) non sélectionnée(s) retirée(s)")
    
    if log_kpis and enrichments:
        cols_after_enrich = set(df.columns)
        new_from_enrich = cols_after_enrich - cols_before_enrich
        if new_from_enrich:
            logger.info(f"[PIPELINE] ✅ ÉTAPE 2 terminée : {len(new_from_enrich)} nouvelle(s) colonne(s) : {list(new_from_enrich)}")
    
    # ============================================================================
    # ÉTAPE 3 : MÉTHODES
    # ============================================================================
    only_selected = list(usage.get("methods") or [])
    cols_before_methods = set(df.columns)
    
    if only_selected:
        if log_kpis:
            logger.info(f"[PIPELINE] ⚙️ ÉTAPE 3 : Application de {len(only_selected)} méthode(s) : {only_selected}")
        
        df = _apply_methods(df, gabarit_name, gabarit_version, only=only_selected)
        
        if log_kpis:
            cols_after_methods = set(df.columns)
            new_from_methods = cols_after_methods - cols_before_methods
            if new_from_methods:
                logger.info(f"[PIPELINE] ✅ ÉTAPE 3 terminée : {len(new_from_methods)} nouvelle(s) colonne(s) : {list(new_from_methods)}")
    
    # ============================================================================
    # ÉTAPE 4 : SCRIPT PYTHON
    # ============================================================================
    code = (usage.get("overlay_python") or "").strip()
    cols_before_script = set(df.columns)
    
    if code:
        if log_kpis:
            logger.info(f"[PIPELINE] 🧪 ÉTAPE 4 : Script Python ({len(code)} caractères)")
            logger.info(f"[PIPELINE]   • Colonnes avant script: {len(df.columns)}")
            logger.info(f"[PIPELINE]   • Paramètres disponibles: {list(params.keys()) if params else []}")
        
        df2, err2 = _apply_overlay(df, code, params=params)
        if err2 is None and isinstance(df2, pd.DataFrame):
            df = df2
            if log_kpis:
                cols_after_script = set(df.columns)
                new_from_script = cols_after_script - cols_before_script
                removed_by_script = cols_before_script - cols_after_script
                logger.info(f"[PIPELINE] ✅ ÉTAPE 4 terminée : {len(df.columns)} colonnes")
                if new_from_script:
                    logger.info(f"[PIPELINE]   • Nouvelles: {list(new_from_script)}")
                if removed_by_script:
                    logger.info(f"[PIPELINE]   • Retirées: {list(removed_by_script)}")
        else:
            if log_kpis:
                logger.error(f"[PIPELINE] ❌ ERREUR ÉTAPE 4 : {err2}")
            return None, err2
    
    # ============================================================================
    # ÉTAPE 4-bis : SÉCURISATION DOUBLONS
    # ============================================================================
    if df.columns.duplicated().any():
        dup_names = list(df.columns[df.columns.duplicated(keep=False)])
        if log_kpis:
            logger.warning(f"[PIPELINE] ⚠️ Doublons de colonnes détectés : {dup_names[:10]}")
        df = df.loc[:, ~df.columns.duplicated(keep="first")]
    
    # ============================================================================
    # ÉTAPE 5 : RENOMMAGES
    # ============================================================================
    ren: dict[str, str] = usage.get("final_renames") or {}
    if ren:
        safe_map = {k: v for k, v in ren.items() if k in df.columns and v and v != k}
        if safe_map:
            if log_kpis:
                logger.info(f"[PIPELINE] 🏷️ ÉTAPE 5 : {len(safe_map)} renommage(s)")
                for old, new in list(safe_map.items())[:5]:
                    logger.info(f"[PIPELINE]   • {old} → {new}")
            df = df.rename(columns=safe_map)
    
    # ============================================================================
    # ÉTAPE 6 : TRI
    # ============================================================================
    sort_rules = usage.get("final_sort") or []
    if isinstance(sort_rules, list) and sort_rules:
        by: list[str] = []
        ascending: list[bool] = []
        for r in sort_rules:
            if not isinstance(r, dict):
                continue
            col = (r.get("col") or "").strip()
            if not col:
                continue
            col_now = ren.get(col, col)
            if col_now in df.columns:
                by.append(col_now)
                ascending.append(bool(r.get("asc", True)))
        
        if by:
            if log_kpis:
                logger.info(f"[PIPELINE] 📊 ÉTAPE 6 : Tri sur {len(by)} colonne(s)")
            df = df.sort_values(by=by, ascending=ascending, kind="mergesort", ignore_index=True)
    
    # ============================================================================
    # ÉTAPE 7 : ORDRE / EXCLUSIONS
    # ============================================================================
    src_order = usage.get("final_order") or df.columns.tolist()
    src_excl = set(usage.get("final_excludes") or [])
    
    # Mapper l'ordre via les renommages
    mapped_order = []
    seen = set()
    for c in src_order:
        cc = ren.get(c, c)
        if cc in df.columns and cc not in seen:
            mapped_order.append(cc)
            seen.add(cc)
    
    # Exclusions
    excl_names = set()
    for c in src_excl:
        excl_names.add(c)
        rc = ren.get(c)
        if rc:
            excl_names.add(rc)
    
    final_cols = [c for c in mapped_order if c in df.columns and c not in excl_names] + \
                 [c for c in df.columns if c not in mapped_order and c not in excl_names]
    
    if log_kpis:
        logger.info(f"[PIPELINE] 📦 ÉTAPE 7 : Ordre final")
        logger.info(f"[PIPELINE]   • Exclusions: {len(excl_names)} colonne(s)")
        logger.info(f"[PIPELINE]   • Colonnes finales: {len(final_cols)}")
        logger.info(f"[PIPELINE] ✅ FIN : {len(df)} lignes × {len(final_cols)} colonnes")
    
    return df[final_cols], None