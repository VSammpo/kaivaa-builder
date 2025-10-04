# backend/services/dataset_service.py
from __future__ import annotations
import pandas as pd
from loguru import logger
from typing import Dict, List

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
