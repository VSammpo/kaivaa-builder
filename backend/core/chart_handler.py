"""
Conversion de graphiques du master Excel (template livrable) en images
"""


import os
from pathlib import Path
from typing import List, Dict, Optional
from loguru import logger

from backend.core.excel_handler import excel_app_context


class ChartExporter:
    """Exporte les graphiques Excel en images PNG"""
    
    def __init__(self, excel_path: str, output_dir: Optional[str] = None):
        """
        Initialise l'exporteur de graphiques.
        
        Args:
            excel_path: Chemin vers le fichier Excel
            output_dir: Dossier de sortie pour les images (défaut: temp)
        """
        self.excel_path = excel_path
        
        if output_dir:
            self.output_dir = Path(output_dir)
        else:
            import tempfile
            self.output_dir = Path(tempfile.gettempdir()) / "kaivaa_charts"
        
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def export_all_charts(self) -> Dict[str, List[str]]:
        """
        Exporte tous les graphiques du fichier Excel.
        
        Returns:
            Dict {sheet_name: [liste des chemins images]}
        """
        logger.info(f"Export des graphiques depuis {Path(self.excel_path).name}")
        
        charts_map = {}
        
        with excel_app_context(self.excel_path, visible=False, read_only=True) as (app, wb):
            for sheet in wb.sheets:
                sheet_charts = self._export_sheet_charts(sheet, sheet.name)
                if sheet_charts:
                    charts_map[sheet.name] = sheet_charts
        
        total_charts = sum(len(charts) for charts in charts_map.values())
        logger.success(f"{total_charts} graphiques exportés")
        
        return charts_map
    
    def export_chart_by_name(self, chart_name: str, sheet_name: Optional[str] = None) -> Optional[str]:
        """
        Exporte un graphique spécifique par son nom.
        
        Args:
            chart_name: Nom du graphique Excel
            sheet_name: Nom de la feuille (optionnel, recherche auto si None)
            
        Returns:
            Chemin de l'image exportée ou None
        """
        with excel_app_context(self.excel_path, visible=False, read_only=True) as (app, wb):
            if sheet_name:
                try:
                    sheet = wb.sheets[sheet_name]
                    return self._export_named_chart(sheet, chart_name, sheet_name)
                except Exception as e:
                    logger.warning(f"Graphique '{chart_name}' non trouvé dans '{sheet_name}' : {e}")
                    return None
            
            for sheet in wb.sheets:
                try:
                    image_path = self._export_named_chart(sheet, chart_name, sheet.name)
                    if image_path:
                        return image_path
                except:
                    continue
            
            logger.warning(f"Graphique '{chart_name}' introuvable")
            return None
    
    def _export_sheet_charts(self, sheet, sheet_name: str) -> List[str]:
        """Exporte tous les graphiques d'une feuille"""
        exported_paths = []
        
        try:
            charts = sheet.api.ChartObjects()
            chart_count = charts.Count
            
            if chart_count == 0:
                return []
            
            logger.debug(f"Export de {chart_count} graphiques depuis '{sheet_name}'")
            
            for i in range(1, chart_count + 1):
                try:
                    chart = charts(i)
                    chart_name = chart.Name
                    
                    safe_sheet_name = self._sanitize_filename(sheet_name)
                    safe_chart_name = self._sanitize_filename(chart_name)
                    image_name = f"{safe_sheet_name}_{safe_chart_name}_{i}.png"
                    image_path = self.output_dir / image_name
                    
                    chart.Chart.Export(str(image_path))
                    exported_paths.append(str(image_path))
                    
                    logger.debug(f"Graphique '{chart_name}' exporté : {image_path.name}")
                
                except Exception as e:
                    logger.warning(f"Erreur export graphique {i} de '{sheet_name}' : {e}")
                    continue
        
        except Exception as e:
            logger.warning(f"Erreur accès graphiques de '{sheet_name}' : {e}")
        
        return exported_paths
    
    def _export_named_chart(self, sheet, chart_name: str, sheet_name: str) -> Optional[str]:
        """Exporte un graphique spécifique par son nom"""
        try:
            charts = sheet.api.ChartObjects()
            
            for i in range(1, charts.Count + 1):
                chart = charts(i)
                if chart.Name == chart_name:
                    safe_sheet_name = self._sanitize_filename(sheet_name)
                    safe_chart_name = self._sanitize_filename(chart_name)
                    image_name = f"{safe_sheet_name}_{safe_chart_name}.png"
                    image_path = self.output_dir / image_name
                    
                    chart.Chart.Export(str(image_path))
                    logger.info(f"Graphique '{chart_name}' exporté : {image_path.name}")
                    
                    return str(image_path)
        
        except Exception as e:
            logger.warning(f"Erreur export graphique '{chart_name}' : {e}")
        
        return None
    
    def _sanitize_filename(self, name: str) -> str:
        """Nettoie un nom pour l'utiliser dans un nom de fichier"""
        import re
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', name)
        return sanitized[:50]
    
    def cleanup(self) -> None:
        """Supprime les images exportées"""
        try:
            import shutil
            if self.output_dir.exists():
                shutil.rmtree(self.output_dir)
                logger.debug(f"Dossier de graphiques nettoyé : {self.output_dir}")
        except Exception as e:
            logger.warning(f"Erreur nettoyage graphiques : {e}")