"""
Utilitaires pour l'injection de DataFrames dans Excel
Adapté de spirits_study pour KAIVAA Builder
"""

import pandas as pd
from typing import Dict, Optional
from loguru import logger
import os
from backend.core.excel_handler import excel_app_context


def inject_dataframe_to_excel(
    path: str,
    sheet_name: str,
    table_name: str,
    df: pd.DataFrame,
    filter_cells: Optional[Dict[str, str]] = None,
    validate_prix_au_volume: bool = False
) -> None:
    """
    Injecte un DataFrame dans un tableau nommé d'un fichier Excel existant.
    
    Args:
        path: Chemin complet vers le fichier Excel
        sheet_name: Nom de la feuille contenant le tableau
        table_name: Nom du tableau structuré Excel
        df: Données à injecter
        filter_cells: Dictionnaire de cellules à mettre à jour (ex: {"C2": "Gin"})
        validate_prix_au_volume: Si True, valide les colonnes de prix (optionnel)
    """
    if df.empty:
        logger.warning(f"DataFrame vide pour {table_name} - injection ignorée")
        return

    if filter_cells is None:
        filter_cells = {}

    with excel_app_context(path) as (app, wb):
        try:
            sht = wb.sheets[sheet_name]
        except Exception:
            raise ValueError(f"Feuille '{sheet_name}' introuvable dans {path}")

        # Application des filtres
        for cell, value in filter_cells.items():
            try:
                sht.range(cell).value = value
            except Exception as e:
                logger.warning(f"Erreur injection filtre {cell}={value} : {e}")

        # Récupération du tableau structuré
        try:
            table = sht.api.ListObjects(table_name)
        except Exception:
            raise ValueError(f"Tableau nommé '{table_name}' introuvable dans '{sheet_name}'")

        # Nettoyage et injection
        try:
            data_body_range = table.DataBodyRange
            if data_body_range:
                data_body_range.ClearContents()

            start_cell = sht.range((data_body_range.Row, data_body_range.Column))
            start_cell.options(index=False, header=False).value = df

            wb.save()
            logger.info(f"Données injectées dans '{sheet_name}' ({table_name}) : {len(df)} lignes")

        except Exception as e:
            raise RuntimeError(f"Erreur lors de l'injection dans {table_name} : {e}")


def inject_single_cell_value(path: str, sheet_name: str, cell: str, value: str) -> None:
    """
    Injecte une valeur unique dans une cellule spécifique.
    
    Args:
        path: Chemin vers le fichier Excel
        sheet_name: Nom de la feuille
        cell: Référence de la cellule (ex: "C4")
        value: Valeur à injecter
    """
    if not os.path.exists(path):
        logger.error(f"Fichier Excel non trouvé : {path}")
        return
    
    try:
        with excel_app_context(path) as (app, wb):
            sheet_names = [ws.name for ws in wb.sheets]
            if sheet_name not in sheet_names:
                logger.error(f"Feuille '{sheet_name}' non trouvée. Disponibles : {sheet_names}")
                return
            
            ws = wb.sheets[sheet_name]
            ws.range(cell).value = value
            wb.save()
            
            logger.debug(f"Valeur '{value}' injectée en {sheet_name}!{cell}")
    
    except Exception as e:
        logger.error(f"Erreur injection cellule {cell} : {e}")


def validate_excel_injection_config(
    path: str,
    injections: Dict[str, Dict[str, any]]
) -> bool:
    """
    Valide la configuration d'injection avant exécution.
    
    Args:
        path: Chemin vers le fichier Excel
        injections: Configuration à valider
        
    Returns:
        True si valide, False sinon
    """
    import os
    
    if not os.path.exists(path):
        logger.error(f"Fichier Excel introuvable : {path}")
        return False

    if not injections:
        logger.error("Configuration d'injection vide")
        return False

    for table_name, config in injections.items():
        if not isinstance(config, dict):
            logger.error(f"Configuration invalide pour {table_name}")
            return False
            
        if "sheet_name" not in config or "df" not in config:
            logger.error(f"Clés manquantes pour {table_name}")
            return False
            
        if not isinstance(config["df"], pd.DataFrame):
            logger.error(f"'df' doit être un DataFrame pour {table_name}")
            return False

    logger.info(f"Configuration d'injection validée ({len(injections)} tables)")
    return True