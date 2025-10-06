# backend/services/dataset_service.py
from __future__ import annotations
import pandas as pd
from loguru import logger
from typing import Dict, List
# --- AJOUT : loaders "full" pour source par défaut --------------------------------
from typing import Optional, Any
import os
from pathlib import Path

# --- AJOUT/CONFIRMATION : loaders "full" pour source par défaut -----------------
from typing import Optional
import os
import pandas as pd
from loguru import logger


# === helpers pour appliquer un script python défini dans la "source" =========
def _apply_source_python(df: pd.DataFrame, source: dict) -> pd.DataFrame:
    """
    Si la source contient une clé 'python', on exécute le script dans un
    namespace {'df': df, 'pd': pd}. Le script peut soit modifier df in-place,
    soit retourner un nouveau DataFrame via 'df = ...'.
    """
    code = (source or {}).get("python")
    if not code or not isinstance(code, str) or not code.strip():
        return df
    try:
        local_vars = {"df": df, "pd": pd}
        exec(code, {}, local_vars)
        new_df = local_vars.get("df")
        if isinstance(new_df, pd.DataFrame):
            return new_df
        return df
    except Exception as e:
        logger.exception(f"[dataset_service] Erreur script 'python' dans source: {e}")
        return df


def _load_dataframe_from_source(source: dict) -> Optional[pd.DataFrame]:
    """
    Lecture *complète* d'une source (csv, parquet, sql) + application éventuelle
    du script 'python' présent dans la config de la source.
    """
    if not source or not isinstance(source, dict):
        return None

    stype = (source.get("type") or "").strip().lower()
    try:
        if stype == "csv":
            path = source.get("path")
            if not path or not os.path.exists(path):
                logger.warning(f"[dataset_service] CSV introuvable: {path}")
                return None
            sep = source.get("sep", ",")
            enc = source.get("encoding", "utf-8-sig")
            df = pd.read_csv(path, sep=sep, encoding=enc)
            df = _apply_source_python(df, source)
            logger.info(f"[dataset_service] CSV lu: {path}  shape={df.shape}")
            return df

        if stype in ("parquet", "pq"):
            path = source.get("path")
            if not path or not os.path.exists(path):
                logger.warning(f"[dataset_service] Parquet introuvable: {path}")
                return None
            df = pd.read_parquet(path)
            df = _apply_source_python(df, source)
            logger.info(f"[dataset_service] Parquet lu: {path}  shape={df.shape}")
            return df

        if stype == "sql":
            conn = source.get("connection")
            query = source.get("query")
            if not query:
                return None
            if conn:
                from sqlalchemy import create_engine
                eng = create_engine(conn)
                df = pd.read_sql(query, eng)
                df = _apply_source_python(df, source)
                logger.info(f"[dataset_service] SQL lu  shape={df.shape}")
                return df
            return None

    except Exception as e:
        logger.exception(f"[dataset_service] Erreur lecture source: {e}")
        return None

    return None


def get_default_dataframe_for_gabarit(gabarit_name: str, gabarit_version: str = "v1",
                                      template_id: int | None = None) -> Optional[pd.DataFrame]:
    """
    Retourne le DataFrame 'par défaut' COMPLET pour un gabarit/version.
    1) Cherche une source déclarée dans le registre
    2) La charge (csv/parquet/sql)
    3) None si introuvable
    """
    try:
        from backend.services.gabarit_registry import get_default_source
        src = get_default_source(gabarit_name, gabarit_version)
        if src:
            return _load_dataframe_from_source(src)
    except Exception as e:
        logger.warning(f"[dataset_service] Pas de source 'full' pour {gabarit_name} v{gabarit_version}: {e}")
    return None


def load_csv(source: dict) -> pd.DataFrame:
    """
    source: {"type":"csv", "path": "...", "sep": ";", "encoding": "utf-8-sig"}
    """
    assert source.get("type") == "csv", "MVP: type 'csv' uniquement"
    path = source.get("path")
    sep = source.get("sep", ",")
    encoding = source.get("encoding", "utf-8-sig")
    df = pd.read_csv(path, sep=sep, encoding=encoding)
    logger.info(f"[CSV] lu: {path}  shape={df.shape}")
    return df

def prepare_for_usage(df: pd.DataFrame, columns_enabled: List[str]) -> pd.DataFrame:
    """
    - Réordonne les colonnes selon columns_enabled
    - Ajoute les colonnes manquantes (vides) si besoin
    - Laisse passer les colonnes en plus (elles seront ignorées à l'injection si non demandées)
    """
    df = df.copy()
    for col in columns_enabled:
        if col not in df.columns:
            df[col] = pd.NA
    # réordonner: colonnes demandées d'abord
    ordered = columns_enabled + [c for c in df.columns if c not in columns_enabled]
    return df[ordered]

# === Alignement DataFrame ↔ colonnes attendues (non bloquant) ================

from typing import List, Tuple, Dict, Any
import pandas as pd

def align_df_to_expected_columns(df: pd.DataFrame, expected_columns: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    - Ajoute les colonnes manquantes (valeur NA)
    - Conserve l'ordre: expected_columns d'abord, puis les colonnes extra
    - Ne lève pas d'exception: retourne (df_aligne, warnings)
    warnings = {"missing": [...], "extra": [...]}
    """
    expected = [c for c in (expected_columns or []) if isinstance(c, str) and c.strip()]
    cur_cols = list(df.columns)

    missing = [c for c in expected if c not in cur_cols]
    for c in missing:
        df[c] = pd.NA

    ordered = expected + [c for c in df.columns if c not in expected]
    aligned = df[ordered]

    extra = [c for c in cur_cols if c not in expected]
    warnings = {"missing": missing, "extra": extra}
    return aligned, warnings
