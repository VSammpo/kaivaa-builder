from __future__ import annotations
from pathlib import Path
import xlwings as xw
import pandas as pd
from loguru import logger
import re
import unicodedata
from typing import List, Dict, Tuple, Optional

# ------------------------------------------------------------
# Canon + align: si tu as déjà collé ma version "tolérante",
# garde-la ; sinon voici une version sûre.
# ------------------------------------------------------------
def _canon(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\ufeff", "").replace("\u200b", "")
    s = s.replace("_", " ")
    s = re.sub(r"[\u00A0\u202F\s]+", " ", s, flags=re.UNICODE)
    return s.strip().lower()

def align_df_to_expected_columns(
    df: pd.DataFrame,
    expected_columns: List[str] | None
) -> Tuple[pd.DataFrame, Dict]:
    if not expected_columns:
        return df.copy(), {"missing": [], "extra": [], "mapping": {}}

    expected = list(expected_columns)
    canon_to_dfcol: Dict[str, str] = {_canon(c): c for c in df.columns}
    mapping: Dict[str, Optional[str]] = {}
    missing: List[str] = []

    selected: List[str] = []
    for exp in expected:
        src = canon_to_dfcol.get(_canon(exp))
        mapping[exp] = src
        if src is None:
            missing.append(exp)
            selected.append(exp)  # placeholder (on créera la colonne)
        else:
            selected.append(src)

    extra = [orig for k, orig in canon_to_dfcol.items() if k not in {_canon(e) for e in expected}]

    df2 = df.copy()
    # Crée les colonnes manquantes
    for exp in missing:
        df2[exp] = pd.NA

    # Ordre final
    final_cols = [c if c in df2.columns else c for c in selected]
    df2 = df2[final_cols]

    return df2, {"missing": missing, "extra": extra, "mapping": mapping}

# ------------------------------------------------------------
# Helpers Excel
# ------------------------------------------------------------
# Excel enum pour la source d'un ListObject
XL_LISTOBJECT_SRC_EXTERNAL = 0  # xlSrcExternal
XL_LISTOBJECT_SRC_RANGE    = 1  # xlSrcRange
XL_LISTOBJECT_SRC_XML      = 2  # xlSrcXml

def _get_listobject_by_name(sheet, table_name: str):
    target = table_name.strip().lower()
    for t in sheet.api.ListObjects:
        if str(t.Name).strip().lower() == target:
            return t
    return None

def _listobject_col_count(lo) -> int:
    try:
        return int(lo.ListColumns.Count)
    except Exception:
        # Fallback (rare)
        try:
            hdr = lo.HeaderRowRange
            vals = hdr.Value
            if isinstance(vals, (list, tuple)):
                row = vals[0] if isinstance(vals[0], (list, tuple)) else vals
                return len(row)
        except Exception:
            pass
        return 0

def _listobject_headers(lo) -> List[str]:
    names = []
    try:
        # Plus robuste : passer par ListColumns(i).Name
        for i in range(1, int(lo.ListColumns.Count) + 1):
            names.append(str(lo.ListColumns.Item(i).Name))
        return names
    except Exception:
        # Fallback via HeaderRowRange (peut être vide si pas d'en-têtes "affichés")
        try:
            vals = lo.HeaderRowRange.Value
            if isinstance(vals, (list, tuple)) and len(vals) > 0:
                row = vals[0] if isinstance(vals[0], (list, tuple)) else vals
                return ["" if v is None else str(v) for v in row]
        except Exception:
            pass
        return names

# ------------------------------------------------------------
# Injection PRINCIPALE
# ------------------------------------------------------------
def inject_dataframe(
    xlsx_path: str | Path,
    sheet_name: str,
    table_name: str,
    df: pd.DataFrame,
    expected_columns: List[str] | None = None
) -> dict:
    xlsx_path = str(xlsx_path)
    
    # Alignement
    df_aligned, base_warn = align_df_to_expected_columns(df, expected_columns)
    
    app = xw.App(visible=False, add_book=False)
    wb = None
    try:
        wb = app.books.open(xlsx_path, update_links=False, read_only=False)
        sht = wb.sheets[sheet_name]
        
        lo = _get_listobject_by_name(sht, table_name)
        if lo is None:
            raise RuntimeError(f"Table '{table_name}' introuvable")
        
        headers = _listobject_headers(lo)
        
        # Préparer le DF dans l'ordre des headers
        warnings = {"ignored_df_columns": [], "added_empty_columns": []}
        canon_df = {_canon(c): c for c in df_aligned.columns}
        df_to_write_cols = []
        
        for h in headers:
            src = canon_df.get(_canon(h))
            if src is None:
                df_aligned[h] = pd.NA
                df_to_write_cols.append(h)
                warnings["added_empty_columns"].append(h)
            else:
                df_to_write_cols.append(src)
        
        ignored = [c for c in df_aligned.columns if _canon(c) not in {_canon(h) for h in headers}]
        if ignored:
            warnings["ignored_df_columns"] = ignored
        
        df_to_write = df_aligned[df_to_write_cols]
        n_rows = len(df_to_write)
        
        # Clear UNIQUEMENT les données (pas la structure)
        if lo.DataBodyRange is not None:
            lo.DataBodyRange.ClearContents()
        
        # Écriture
        first_cell = sht.range(lo.HeaderRowRange.Address).offset(1, 0)
        first_cell.options(index=False, header=False).value = df_to_write
        logger.info(f"[DEBUG] Données écrites. Vérification de la table...")
        logger.info(f"[DEBUG] DataBodyRange après écriture : {lo.DataBodyRange.Rows.Count if lo.DataBodyRange else 0} lignes")
        logger.info(f"[DEBUG] ListObject Range : {lo.Range.Address}")

        # Forcer un recalcul IMMÉDIAT
        try:
            wb.app.api.CalculateFullRebuild()
            logger.info(f"[DEBUG] Recalcul forcé effectué")
        except Exception as e:
            logger.warning(f"[DEBUG] Impossible de forcer le recalcul : {e}")
        
        wb.save()
        
        all_warn = {**base_warn, **warnings}
        return {
            "rows": n_rows,
            "cols": len(headers),
            "warnings": all_warn,
            "mode": "simple"
        }
    
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            app.quit()