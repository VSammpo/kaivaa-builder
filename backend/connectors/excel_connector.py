"""
Connecteur pour lire des données depuis Excel
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional
from loguru import logger

from backend.core.excel_handler import excel_app_context


class ExcelConnector:
    """Connecteur pour lire des tableaux structurés Excel"""
    
    def __init__(self, excel_path: str):
        """
        Initialise le connecteur.
        
        Args:
            excel_path: Chemin vers le fichier Excel
        """
        self.excel_path = Path(excel_path)
        
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Fichier Excel introuvable : {excel_path}")
    
    def read_table(self, table_name: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """
        Lit un tableau structuré Excel.
        
        Args:
            table_name: Nom du tableau structuré (ex: "Performance")
            sheet_name: Nom de la feuille (optionnel, recherche auto si None)
            
        Returns:
            DataFrame avec les données
        """
        logger.info(f"Lecture du tableau '{table_name}' depuis {self.excel_path.name}")
        
        with excel_app_context(str(self.excel_path), visible=False, read_only=True) as (app, wb):
            # Si sheet_name fourni, chercher directement
            if sheet_name:
                try:
                    sheet = wb.sheets[sheet_name]
                    df = self._read_table_from_sheet(sheet, table_name)
                    logger.success(f"Tableau '{table_name}' lu : {len(df)} lignes")
                    return df
                except Exception as e:
                    raise ValueError(f"Erreur lecture tableau '{table_name}' dans '{sheet_name}' : {e}")
            
            # Sinon, chercher dans toutes les feuilles
            for sheet in wb.sheets:
                try:
                    df = self._read_table_from_sheet(sheet, table_name)
                    logger.success(f"Tableau '{table_name}' trouvé dans '{sheet.name}' : {len(df)} lignes")
                    return df
                except:
                    continue
            
            raise ValueError(f"Tableau '{table_name}' introuvable dans {self.excel_path.name}")
    
    def _read_table_from_sheet(self, sheet, table_name: str) -> pd.DataFrame:
        """Lit un tableau depuis une feuille spécifique"""
        # Rechercher le tableau structuré
        table = None
        for t in sheet.api.ListObjects:
            if t.Name.strip().lower() == table_name.lower():
                table = t
                break
        
        if not table:
            raise ValueError(f"Tableau '{table_name}' non trouvé dans '{sheet.name}'")
        
        # Lire les données
        data_range = table.DataBodyRange
        if not data_range:
            return pd.DataFrame()
        
        # Convertir en DataFrame
        values = sheet.range(data_range.Address).options(pd.DataFrame, index=False, header=True).value
        
        return values
    
    def read_cell(self, sheet_name: str, cell: str) -> any:
        """
        Lit une cellule spécifique.
        
        Args:
            sheet_name: Nom de la feuille
            cell: Référence de la cellule (ex: "C3")
            
        Returns:
            Valeur de la cellule
        """
        with excel_app_context(str(self.excel_path), visible=False, read_only=True) as (app, wb):
            sheet = wb.sheets[sheet_name]
            value = sheet.range(cell).value
            return value
    
    def list_tables(self) -> Dict[str, List[str]]:
        """
        Liste tous les tableaux structurés du fichier.
        
        Returns:
            Dict {sheet_name: [table_names]}
        """
        tables_map = {}
        
        with excel_app_context(str(self.excel_path), visible=False, read_only=True) as (app, wb):
            for sheet in wb.sheets:
                sheet_tables = []
                try:
                    for t in sheet.api.ListObjects:
                        sheet_tables.append(t.Name)
                except:
                    pass
                
                if sheet_tables:
                    tables_map[sheet.name] = sheet_tables
        
        return tables_map