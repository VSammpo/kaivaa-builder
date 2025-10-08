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


def _apply_overlay(df: pd.DataFrame, code: str) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """Exécute le script Python overlay sur le DataFrame"""
    if not code or not code.strip():
        return df, None
    
    try:
        local_vars = {"pd": pd, "df": df.copy()}
        exec(code, {}, local_vars)
        result = local_vars.get("df", None)
        if isinstance(result, pd.DataFrame):
            return result, None
        return None, "Le script n'a pas produit de DataFrame 'df'."
    except Exception as ex:
        import traceback
        tb = traceback.format_exc()
        return None, f"{type(ex).__name__}: {ex}\n{tb}"


def build_table_from_usage(
    usage: dict,
    *,
    full: bool = True,
    log_kpis: bool = False
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Construit le DataFrame final à partir d'un usage de gabarit.
    """
    
    gabarit_name = usage.get("gabarit_name", "")
    gabarit_version = usage.get("gabarit_version", "v1")
    
    # 1. Charger la base (ou utiliser celle fournie)
    if "_source_df" in usage:
        df = usage["_source_df"].copy()
        is_preview = False
        complete = True  # ✅ AJOUTER CETTE LIGNE
        if log_kpis:
            logger.info(f"📦 Base fournie : {len(df)} lignes, {len(df.columns)} colonnes")
    else:
        df, is_preview = _load_df_for_gabarit(gabarit_name, gabarit_version, full=full)
        if df is None:
            if full:
                return None, f"FULL demandé mais aucune source n'est définie pour {gabarit_name} (v{gabarit_version})."
            return None, f"Aucune donnée par défaut pour {gabarit_name} (v{gabarit_version})."
        
        complete = (not is_preview) if full else True  # ✅ Déjà présent
        
        if log_kpis:
            logger.info(f"📦 Base : {gabarit_name} (v{gabarit_version}) - {len(df)} lignes, {len(df.columns)} colonnes - Mode: {'PREVIEW' if is_preview else 'FULL'}")
        
    # Mémoriser les colonnes ajoutées par enrichissements
    added_enriched_cols: set[str] = set()
    
    # 2. Enrichissements (jointures multi-sauts)
    for idx_enrich, e in enumerate(usage.get("enrichments") or []):
        path = e.get("path") or []
        if not path:
            continue
        
        for i, step in enumerate(path):
            frm, left_key, to, right_key = step
            
            df_to, is_prev_to = _load_df_for_gabarit(to, "v1", full=full)
            if df_to is None:
                if full:
                    return df, f"FULL demandé mais aucune source n'est définie pour {to} (v1)."
                return df, f"Aucune donnée par défaut pour {to} (v1)."
            
            if log_kpis:
                logger.info(f"🔗 Enrichissement #{idx_enrich+1}.{i+1} → {to} - {len(df_to)} lignes - Mode: {'PREVIEW' if is_prev_to else 'FULL'}")
            
            # Colonnes à rapatrier
            cols_to_fetch: set[str] = set()
            
            if i + 1 < len(path):
                # Transporter la left_key du saut suivant
                next_left_key = path[i + 1][1]
                if next_left_key and next_left_key in df_to.columns:
                    cols_to_fetch.add(next_left_key)
            else:
                # Dernier saut : colonnes sélectionnées
                for c in (e.get("columns") or []):
                    if c and c in df_to.columns:
                        cols_to_fetch.add(c)
            
            # Ne pas inclure la clé droite
            cols_to_fetch.discard(right_key)
            
            # Éviter les colonnes déjà présentes
            cols_to_add = [c for c in cols_to_fetch if c not in df.columns]
            
            # Vérifier les clés
            if left_key not in df.columns or right_key not in df_to.columns:
                return df, f"Clé manquante pour joindre {frm} → {to} ({left_key} / {right_key})."
            
            # Normaliser les clés
            df[left_key] = _normalize_key(df[left_key], left_key)
            df_to[right_key] = _normalize_key(df_to[right_key], right_key)
            
            # Sous-ensemble de la droite
            right_cols = [right_key] + cols_to_add
            seen = set()
            right_cols = [c for c in right_cols if not (c in seen or seen.add(c))]
            right_subset = df_to[right_cols].drop_duplicates()
            
            # MERGE
            df = df.merge(
                right_subset,
                left_on=left_key,
                right_on=right_key,
                how="left",
            )
            
            # Supprimer la clé de droite
            if right_key in df.columns:
                df.drop(columns=[right_key], inplace=True)
            
            # Mémoriser les colonnes ajoutées
            added_enriched_cols.update([c for c in cols_to_add if c in df.columns])
            
            if full and is_prev_to:
                complete = False
    
    # Filtrage des colonnes enrichies non sélectionnées
    selected_enriched = set()
    for e in (usage.get("enrichments") or []):
        selected_enriched.update([c for c in (e.get("columns") or []) if c])
    
    to_drop = [c for c in added_enriched_cols if c not in selected_enriched and c in df.columns]
    if to_drop:
        df.drop(columns=to_drop, inplace=True, errors="ignore")
        if log_kpis:
            logger.info(f"🧹 Nettoyage : suppression de {len(to_drop)} colonnes enrichies non sélectionnées")
    
    # 3. Méthodes (colonnes calculées)
    only_selected = list(usage.get("methods") or [])
    if only_selected:
        df = _apply_methods(df, gabarit_name, gabarit_version, only=only_selected)
        if log_kpis:
            logger.info(f"⚙️ Méthodes appliquées : {', '.join(only_selected)}")
    
    # 4. Script Python overlay
    code = (usage.get("overlay_python") or "").strip()
    if code:
        df2, err2 = _apply_overlay(df, code)
        if err2 is None and isinstance(df2, pd.DataFrame):
            df = df2
            if log_kpis:
                logger.info(f"🧪 Script Python appliqué")
        else:
            return None, err2
    
    # 5. Renommages de colonnes
    ren: dict[str, str] = usage.get("final_renames") or {}
    if ren:
        safe_map = {k: v for k, v in ren.items() if k in df.columns and v and v != k}
        if safe_map:
            df = df.rename(columns=safe_map)
            if log_kpis:
                logger.info(f"📝 Renommages : {len(safe_map)} colonnes")
    
    # 6. Tri des lignes
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
            df = df.sort_values(by=by, ascending=ascending, kind="mergesort", ignore_index=True)
            if log_kpis:
                logger.info(f"📊 Tri appliqué sur {len(by)} colonne(s)")
    
    # 6-bis. Sécuriser les noms de colonnes (supprimer les doublons)
    if df.columns.duplicated().any():
        dup_names = list(df.columns[df.columns.duplicated(keep=False)])
        if log_kpis:
            logger.warning(f"⚠️ Noms de colonnes dupliqués détectés : {', '.join(map(str, dup_names[:5]))}")
        df = df.loc[:, ~df.columns.duplicated(keep="first")]
    
    # 7. Ordre final / exclusions
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
    
    # Exclusions : anciens ET nouveaux noms
    excl_names = set()
    for c in src_excl:
        excl_names.add(c)
        rc = ren.get(c)
        if rc:
            excl_names.add(rc)
    
    final_cols = [c for c in mapped_order if c in df.columns and c not in excl_names] + \
                 [c for c in df.columns if c not in mapped_order and c not in excl_names]
    
    if log_kpis:
        logger.info(f"📦 Sortie finale : {len(df)} lignes, {len(final_cols)} colonnes - Complet: {complete}")
    
    return df[final_cols], None