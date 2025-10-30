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

def _get_key_columns(gabarit_name: str, gabarit_version: str) -> list[str]:
    """Retourne les colonnes marquées is_key=True dans un gabarit."""
    try:
        g = get_gabarit(gabarit_name, gabarit_version)
        if g and g.columns:
            return [c.name for c in g.columns if getattr(c, 'is_key', False)]
    except Exception as e:
        logger.warning(f"Impossible de récupérer les colonnes clés pour {gabarit_name}: {e}")
    return []

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
        logger.debug(f"[MERGE {tag}] AVANT : left={len(left_df)} (key {left_on} non-null={l_nonnull}, dtype={l_dtype}), right={len(right_df)} (key {right_on} non-null={r_nonnull}, dtype={r_dtype}), how={how}")
    except Exception:
        pass

    merged = pd.merge(left_df, right_df, how=how, left_on=left_on, right_on=right_on)

    try:
        logger.debug(f"[MERGE {tag}] APRES : {len(merged)} lignes")
    except Exception:
        pass
    return merged


def _normalize_key(series: pd.Series, key_name: str) -> pd.Series:
    """
    Normalise une colonne-clé (strip, casse, NaN, etc.) de façon agressive.
    """
    def _norm(val):
        if pd.isna(val):
            return val
        s = str(val).strip().upper()
        return s if s else None

    return series.apply(_norm)


def _apply_methods(df: pd.DataFrame, gabarit_name: str, gabarit_version: str, only: list[str] | None = None, enrichments: list[dict] | None = None) -> pd.DataFrame:
    """
    Applique les méthodes (colonnes calculées) listées dans 'only'.
    Si only est None ou vide, on ne fait rien.
    
    🔧 CORRECTION : Charge aussi les méthodes des tables enrichies.
    """
    from backend.services.method_executor import apply_method
    
    if not only:
        return df
    
    # Récupérer toutes les méthodes du gabarit de base
    all_methods = list_methods_for_gabarit(gabarit_name, gabarit_version) or []
    
    # Si c'est un dict, convertir en liste
    if isinstance(all_methods, dict):
        all_methods = list(all_methods.values())
    
    # 🔧 CORRECTION : Ajouter les méthodes des tables enrichies
    if enrichments:
        enriched_gabarits = set()
        for e in enrichments:
            path = e.get("path", [])
            if path:
                # Récupérer le gabarit cible du dernier saut
                last_hop = path[-1]
                if len(last_hop) >= 3:
                    target_gabarit = last_hop[2]  # [from, left_key, to, right_key]
                    if target_gabarit not in enriched_gabarits:
                        enriched_gabarits.add(target_gabarit)
                        enriched_methods = list_methods_for_gabarit(target_gabarit, "v1") or []
                        if isinstance(enriched_methods, dict):
                            enriched_methods = list(enriched_methods.values())
                        all_methods.extend(enriched_methods)
    
    # Filtrer selon 'only'
    selected = [m for m in all_methods if isinstance(m, dict) and m.get("name") in only]
    
    df_out = df.copy()
    for method in selected:
        try:
            df_out = apply_method(df_out, method, params={})
        except Exception as e:
            logger.warning(f"Erreur méthode '{method.get('name')}': {e}")
    
    return df_out


def _apply_overlay(df: pd.DataFrame, code: str, params: dict | None = None) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Exécute un script Python overlay sur le DataFrame.
    Le script a accès à : df, pd, params
    Il doit modifier df in-place OU redéfinir df.
    """
    if not code or not code.strip():
        return df, None
    
    try:
        env = {
            "df": df.copy(),
            "pd": pd,
            "params": params or {},
            "__builtins__": __builtins__,
        }
        exec(code, env, env)
        new_df = env.get("df")
        
        if not isinstance(new_df, pd.DataFrame):
            return None, f"Le script doit produire un DataFrame dans 'df' (type actuel: {type(new_df).__name__})"
        
        return new_df, None
    
    except Exception as e:
        return None, f"Erreur script overlay: {e}"


def build_table_from_usage(
    usage: dict,
    *,
    full: bool = True,
    log_kpis: bool = False,
    params: dict | None = None
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Point d'entrée UNIQUE pour construire une table depuis un usage (ou transformation).
    
    Args:
        usage: Dict représentant l'usage (ou transformation)
        full: Si True, charge le FULL dataframe. Si False, utilise le preview (20 lignes).
        log_kpis: Active les logs détaillés du pipeline
        params: Paramètres de template (ex: {"sous_marque": "BOMBAY"})
    
    Returns:
        (DataFrame, erreur) où erreur=None si succès
    """
    
    # Normalisation usage/transformation
    usage = _normalize_to_usage(usage)
    
    gabarit_name = usage.get("gabarit_name", "")
    gabarit_version = usage.get("gabarit_version", "v1")
    
    if not gabarit_name:
        return None, "gabarit_name manquant dans l'usage"
    
    # Récupérer le gabarit
    gabarit = get_gabarit(gabarit_name, gabarit_version)
    if not gabarit:
        return None, f"Gabarit '{gabarit_name}' v{gabarit_version} introuvable"
    
    if log_kpis:
        logger.info(f"[PIPELINE] 🚀 DÉBUT construction table : {gabarit_name} v{gabarit_version}")
        logger.info(f"[PIPELINE]   Mode: {'FULL' if full else 'PREVIEW (20 lignes)'}")
    
    # ============================================================================
    # CHARGEMENT DES DONNÉES (preview ou full)
    # ============================================================================
    
    # Vérifier si un DataFrame custom est fourni
    custom_df = usage.get("_source_df")
    if isinstance(custom_df, pd.DataFrame):
        df = custom_df.copy()
        if log_kpis:
            logger.info(f"[PIPELINE] 📥 Source: DataFrame personnalisé ({len(df)} lignes)")
    elif full:
        # Mode FULL : charger toutes les données
        df = get_default_dataframe_for_gabarit(gabarit_name, gabarit_version)
        if df is None or df.empty:
            return None, f"Impossible de charger les données FULL pour {gabarit_name} v{gabarit_version}"
        if log_kpis:
            logger.info(f"[PIPELINE] 📥 Source: FULL dataframe ({len(df)} lignes)")
    else:
        # Mode PREVIEW : utiliser le preview (20 lignes)
        preview_data = get_default_preview(gabarit_name, gabarit_version)
        if not preview_data or not preview_data.get("rows"):
            return None, f"Aucun preview disponible pour {gabarit_name} v{gabarit_version}"
        
        rows = preview_data.get("rows", [])
        cols = preview_data.get("columns", [])
        df = pd.DataFrame(rows, columns=cols)
        
        if log_kpis:
            logger.info(f"[PIPELINE] 📥 Source: PREVIEW ({len(df)} lignes)")
    
    # ============================================================================
    # ÉTAPE 1 : COLONNES DE BASE + CLÉS OBLIGATOIRES
    # ============================================================================
    cols_enabled = list(usage.get("columns_enabled") or [])
    
    if not cols_enabled:
        cols_enabled = [c.name for c in (gabarit.columns or [])]
    
    # 🔑 CORRECTION CRITIQUE : Forcer l'inclusion des colonnes clés
    key_cols = _get_key_columns(gabarit_name, gabarit_version)
    cols_enabled_with_keys = list(dict.fromkeys(cols_enabled + key_cols))  # Déduplication en gardant l'ordre
    
    if log_kpis:
        logger.info(f"[PIPELINE] 📋 ÉTAPE 1 : Colonnes de base")
        logger.info(f"[PIPELINE]   • Colonnes sélectionnées: {len(cols_enabled)}")
        logger.info(f"[PIPELINE]   • Colonnes clés ajoutées: {key_cols}")
        logger.info(f"[PIPELINE]   • Total: {len(cols_enabled_with_keys)}")
    
    cols_in_df = [c for c in cols_enabled_with_keys if c in df.columns]
    df = df[cols_in_df]
    
    # ============================================================================
    # ÉTAPE 2 : ENRICHISSEMENTS
    # ============================================================================
    enrichments = usage.get("enrichments") or []
    cols_before_enrich = set(df.columns)
    added_enriched_cols = set()
    
    if enrichments and log_kpis:
        logger.info(f"[PIPELINE] 🔗 ÉTAPE 2 : {len(enrichments)} enrichissement(s)")
    
    complete = full
    for idx, e in enumerate(enrichments):
        join_type = e.get("join", "left")
        path = e.get("path", [])
        
        if not path:
            continue
        
        if log_kpis:
            logger.info(f"[PIPELINE]   Enrichissement #{idx+1} (type: {join_type})")
        
        # Parcourir le chemin
        for i, hop in enumerate(path):
            if len(hop) < 4:
                continue
            
            frm, left_key, to, right_key = hop[0], hop[1], hop[2], hop[3]
            
            # Charger la table cible
            is_prev_to = False  # Flag pour savoir si on utilise preview
            if full:
                df_to = get_default_dataframe_for_gabarit(to, "v1")
            else:
                prev_to = get_default_preview(to, "v1")
                is_prev_to = prev_to and prev_to.get("rows")
                if is_prev_to:
                    df_to = pd.DataFrame(prev_to["rows"], columns=prev_to["columns"])
                else:
                    df_to = None
            
            if df_to is None or df_to.empty:
                return df, f"Table cible '{to}' introuvable ou vide."
            
            # Colonnes à ajouter
            cols_to_add = [c for c in (e.get("columns") or []) if c and c in df_to.columns and c not in df.columns]
            
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
            # 🔑 CORRECTION : Si right_key existe déjà dans df, éviter le conflit
            if right_key in df.columns and right_key != left_key:
                # La clé existe déjà côté gauche (c'est une clé commune)
                # On ne veut pas la dupliquer, donc on la retire du sous-ensemble droit
                right_subset_cols = [c for c in right_subset.columns if c != right_key]
                if right_subset_cols:  # S'il reste des colonnes à ajouter
                    right_subset_for_merge = right_subset[right_subset_cols].copy()
                    # Ajouter temporairement la clé pour le merge
                    right_subset_for_merge[right_key] = right_subset[right_key]
                    
                    df = _debug_merge(
                        df,
                        right_subset_for_merge,
                        how="left",
                        left_on=left_key,
                        right_on=right_key,
                        tag=f"{frm}->{to}"
                    )
                    
                    # Supprimer la clé dupliquée si elle existe
                    if right_key in df.columns and left_key in df.columns and left_key != right_key:
                        df.drop(columns=[right_key], inplace=True, errors='ignore')
            else:
                # Cas standard : pas de conflit
                df = _debug_merge(
                    df,
                    right_subset,
                    how="left",
                    left_on=left_key,
                    right_on=right_key,
                    tag=f"{frm}->{to}"
                )
                
                if right_key in df.columns and left_key != right_key:
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
        
        df = _apply_methods(df, gabarit_name, gabarit_version, only=only_selected, enrichments=enrichments)
        
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
    # ÉTAPE 7 : ORDRE / EXCLUSIONS (+ PROTECTION CLÉS + MÉTHODES)
    # ============================================================================
    src_order = usage.get("final_order") or []
    src_excl = set(usage.get("final_excludes") or [])
    
    # 🔑 CORRECTION CRITIQUE : Les colonnes clés ne peuvent JAMAIS être exclues
    key_cols_original = _get_key_columns(gabarit_name, gabarit_version)
    key_cols_renamed = [ren.get(k, k) for k in key_cols_original]  # Prendre en compte les renommages
    
    # ⚙️ CORRECTION : Identifier les colonnes créées par les méthodes
    method_cols = []
    if only_selected:
        # Charger toutes les méthodes du gabarit de base + enrichis
        all_methods_objs = list_methods_for_gabarit(gabarit_name, gabarit_version) or []
        if isinstance(all_methods_objs, dict):
            all_methods_objs = list(all_methods_objs.values())
        
        # Ajouter méthodes des enrichissements
        for e in (enrichments or []):
            path = e.get("path", [])
            if path:
                last_hop = path[-1]
                if len(last_hop) >= 3:
                    target_gabarit = last_hop[2]
                    enriched_methods = list_methods_for_gabarit(target_gabarit, "v1") or []
                    if isinstance(enriched_methods, dict):
                        enriched_methods = list(enriched_methods.values())
                    all_methods_objs.extend(enriched_methods)
        
        # Extraire les output_column des méthodes sélectionnées
        for method_obj in all_methods_objs:
            if isinstance(method_obj, dict) and method_obj.get("name") in only_selected:
                output_col = method_obj.get("output_column")
                if output_col:
                    # Prendre en compte les renommages
                    output_col_renamed = ren.get(output_col, output_col)
                    if output_col_renamed in df.columns:
                        method_cols.append(output_col_renamed)
    
    # Si final_order est vide, utiliser toutes les colonnes actuelles
    if not src_order:
        src_order = list(df.columns)
    
    # Mapper l'ordre via les renommages
    mapped_order = []
    seen = set()
    for c in src_order:
        cc = ren.get(c, c)
        if cc in df.columns and cc not in seen:
            mapped_order.append(cc)
            seen.add(cc)
    
    # 🔑 Ajouter les colonnes clés manquantes dans mapped_order (AU DÉBUT)
    for key_col in key_cols_renamed:
        if key_col in df.columns and key_col not in mapped_order:
            mapped_order.insert(0, key_col)
            if log_kpis:
                logger.warning(f"[PIPELINE] 🔑 Colonne clé '{key_col}' absente de final_order → ajoutée automatiquement")
    
    # ⚙️ Ajouter les colonnes de méthodes manquantes dans mapped_order (APRÈS LES CLÉS)
    for method_col in method_cols:
        if method_col not in mapped_order:
            # Insérer après les clés mais avant le reste
            insert_pos = len(key_cols_renamed)
            mapped_order.insert(insert_pos, method_col)
            if log_kpis:
                logger.warning(f"[PIPELINE] ⚙️ Colonne de méthode '{method_col}' absente de final_order → ajoutée automatiquement")
    
    # Exclusions (SAUF les colonnes clés ET les colonnes de méthodes)
    excl_names = set()
    for c in src_excl:
        # ❌ Ne PAS exclure les colonnes clés
        if c in key_cols_original or c in key_cols_renamed:
            if log_kpis:
                logger.warning(f"[PIPELINE] ⚠️ Tentative d'exclusion de la colonne clé '{c}' → IGNORÉE")
            continue
        
        # ❌ Ne PAS exclure les colonnes de méthodes
        if c in method_cols:
            if log_kpis:
                logger.warning(f"[PIPELINE] ⚠️ Tentative d'exclusion de la colonne de méthode '{c}' → IGNORÉE")
            continue
        
        excl_names.add(c)
        rc = ren.get(c)
        if rc and rc not in key_cols_renamed and rc not in method_cols:
            excl_names.add(rc)
    
    # Construction de la liste finale
    final_cols = [c for c in mapped_order if c in df.columns and c not in excl_names]
    
    # Ajouter les colonnes non ordonnées (sauf exclusions)
    final_cols += [c for c in df.columns if c not in final_cols and c not in excl_names]
    
    if log_kpis:
        logger.info(f"[PIPELINE] 📦 ÉTAPE 7 : Ordre final")
        logger.info(f"[PIPELINE]   • Colonnes clés protégées: {key_cols_renamed}")
        logger.info(f"[PIPELINE]   • Colonnes de méthodes protégées: {method_cols}")
        logger.info(f"[PIPELINE]   • Exclusions appliquées: {len(excl_names)} colonne(s)")
        logger.info(f"[PIPELINE]   • Colonnes finales: {len(final_cols)}")
        logger.info(f"[PIPELINE] ✅ FIN : {len(df)} lignes × {len(final_cols)} colonnes")
    
    return df[final_cols], None