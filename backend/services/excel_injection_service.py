from __future__ import annotations
import xlwings as xw
import pandas as pd
from loguru import logger

try:
    # si disponible (même module)
    from backend.services.dataset_service import align_df_to_expected_columns
except Exception:
    def align_df_to_expected_columns(df: pd.DataFrame, expected_columns):
        return df, {"missing": [], "extra": list(df.columns)}

def inject_dataframe(
    xlsx_path: str,
    sheet_name: str,
    table_name: str,
    df: pd.DataFrame,
    expected_columns: list[str] | None = None
) -> dict:
    """
    Injection:
      - ouvre le fichier Excel
      - trouve le ListObject table_name sur sheet_name
      - clear + écrit df (sans header)
    Si expected_columns est fourni: aligne DF (ajoute colonnes manquantes NA) et
    ne bloque pas l'injection (remonte des warnings).
    Retourne: {"rows": int, "cols": int, "warnings": {...}}
    """
    if not xlsx_path or not sheet_name or not table_name:
        raise ValueError("xlsx_path/sheet_name/table_name requis")

    df_to_write = df.copy()
    warnings = {"missing": [], "extra": []}
    if expected_columns:
        df_to_write, warnings = align_df_to_expected_columns(df_to_write, expected_columns)

    logger.info(f"[INJECT] {xlsx_path} -> {sheet_name}.{table_name} rows={len(df_to_write)} cols={len(df_to_write.columns)}")
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(xlsx_path)
        sht = wb.sheets[sheet_name]

        # trouver ListObject
        lo = None
        for tbl in sht.api.ListObjects:
            if tbl.Name == table_name:
                lo = tbl
                break
        if lo is None:
            raise RuntimeError(f"Table Excel '{table_name}' introuvable sur '{sheet_name}'")

        # vider la DataBodyRange si existe
        if lo.DataBodyRange is not None:
            lo.DataBodyRange.ClearContents()

        # première cellule sous l'entête
        first_cell = sht.range(lo.HeaderRowRange.Address).offset(1, 0)
        first_cell.options(index=False, header=False).value = df_to_write

        wb.save()
        return {"rows": len(df_to_write), "cols": len(df_to_write.columns), "warnings": warnings}
    finally:
        wb.close()
        app.quit()
