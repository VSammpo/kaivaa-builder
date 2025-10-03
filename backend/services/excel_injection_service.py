# backend/services/excel_injection_service.py
from __future__ import annotations
import xlwings as xw
import pandas as pd
from loguru import logger

def inject_dataframe(xlsx_path: str, sheet_name: str, table_name: str, df: pd.DataFrame) -> None:
    """
    Injection simple:
    - ouvre le fichier Excel
    - trouve le ListObject table_name sur sheet_name
    - clear + resize + écrit df[columns_enabled] (si table vide elle sera redimensionnée)
    """
    logger.info(f"[INJECT] {xlsx_path} -> {sheet_name}.{table_name} rows={len(df)} cols={len(df.columns)}")
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(xlsx_path)
        sht = wb.sheets[sheet_name]
        lo = None
        for tbl in sht.api.ListObjects:
            if tbl.Name == table_name:
                lo = tbl
                break
        if lo is None:
            raise RuntimeError(f"Table Excel '{table_name}' introuvable sur '{sheet_name}'")

        # clear data body range si existe
        if lo.DataBodyRange is not None:
            lo.DataBodyRange.ClearContents()

        # Écriture (xlwings se charge d'expand)
        first_cell = sht.range(lo.HeaderRowRange.Address).offset(1, 0)
        first_cell.options(index=False, header=False).value = df

        # Ajuste la taille si nécessaire
        # (Excel étend automatiquement le ListObject quand on écrit sous l'entête)
        wb.save()
    finally:
        wb.close()
        app.quit()
