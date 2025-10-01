"""
Service de génération de rapports
"""

from pathlib import Path
from typing import Dict, Optional, Any
from datetime import datetime
from loguru import logger

from backend.config import PathConfig
from backend.models.template_config import TemplateConfig
from backend.connectors.excel_connector import ExcelConnector
from backend.core.excel_handler import load_replacement_tags, excel_app_context
from backend.core.ppt_handler import (
    powerpoint_app_context,
    replace_tags_in_shape,
    find_slide_by_id,
    check_and_remove_suppressed_slides
)
from backend.core.image_handler import inject_image_to_slide, find_slides_by_ids
from backend.core.batch_processor import BatchProcessor, SlideAxis
from backend.utils.file_utils import get_output_paths, ensure_directories
from backend.utils.cleanup import cleanup_before_run


class ReportService:
    """Service de génération de rapports à partir de templates"""
    
    def __init__(self, template_config: TemplateConfig):
        """
        Initialise le service.
        
        Args:
            template_config: Configuration du template
        """
        self.config = template_config
        self.template_dir = PathConfig.TEMPLATES / self.config.name
    
    def generate_report(
        self,
        parameters: Dict[str, Any],
        output_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Génère un rapport complet.
        
        Args:
            parameters: Paramètres du rapport (ex: {"entreprise": "ACME", "background": "Cuisine"})
            output_name: Nom personnalisé pour les fichiers de sortie
            
        Returns:
            Dict avec chemins des fichiers générés et métadonnées
        """
        logger.info(f"Génération du rapport '{self.config.name}'")
        logger.info(f"Paramètres : {parameters}")
        
        # Validation des paramètres
        self._validate_parameters(parameters)
        
        # Nettoyage préventif
        cleanup_before_run()
        
        # Génération des chemins de sortie
        output_paths = self._generate_output_paths(parameters, output_name)
        ensure_directories(output_paths['excel_path'], output_paths['pptx_path'])
        
        start_time = datetime.now()
        
        try:
            # Étape 1 : Préparation Excel
            logger.info("Étape 1/4 : Préparation Excel")
            excel_path = self._prepare_excel(parameters, output_paths['excel_path'])
            
            # Étape 2 : Lecture des données
            logger.info("Étape 2/4 : Lecture des données")
            data = self._load_data(excel_path)
            
            # Étape 3 : Génération PowerPoint
            logger.info("Étape 3/4 : Génération PowerPoint")
            ppt_path = self._generate_powerpoint(excel_path, output_paths['pptx_path'], parameters)
            
            # Étape 4 : Injection des images
            logger.info("Étape 4/4 : Injection des images")
            self._inject_images(ppt_path, excel_path)
            
            execution_time = (datetime.now() - start_time).total_seconds()
            
            result = {
                "success": True,
                "excel_path": str(excel_path),
                "pptx_path": str(ppt_path),
                "execution_time_seconds": execution_time,
                "parameters": parameters
            }
            
            logger.success(f"Rapport généré en {execution_time:.1f}s")
            return result
        
        except Exception as e:
            logger.error(f"Erreur génération rapport : {e}")
            execution_time = (datetime.now() - start_time).total_seconds()
            
            return {
                "success": False,
                "error": str(e),
                "execution_time_seconds": execution_time,
                "parameters": parameters
            }
    
    def _validate_parameters(self, parameters: Dict[str, Any]) -> None:
        """Valide que tous les paramètres requis sont fournis"""
        for param in self.config.parameters:
            if param.required and param.name not in parameters:
                raise ValueError(f"Paramètre requis manquant : {param.name}")
    
    def _generate_output_paths(self, parameters: Dict[str, Any], custom_name: Optional[str]) -> Dict[str, str]:
        """Génère les chemins de sortie"""
        if custom_name:
            base_name = custom_name
        else:
            # Générer un nom basé sur les paramètres
            param_values = "_".join([str(v) for v in parameters.values()])
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            base_name = f"{self.config.name}_{param_values}_{timestamp}"
        
        output_dir = PathConfig.OUTPUT / self.config.name
        
        return {
            "excel_path": str(output_dir / f"{base_name}.xlsx"),
            "pptx_path": str(output_dir / f"{base_name}.pptx")
        }
    
    def _prepare_excel(self, parameters: Dict[str, Any], output_path: str) -> Path:
        """Prépare le fichier Excel avec les paramètres"""
        import shutil
        
        # Copier le template Excel
        template_excel = self.template_dir / "master.xlsx"
        shutil.copy2(template_excel, output_path)
        
        logger.info(f"Excel copié : {output_path}")
        
        # Injecter les paramètres dans les balises
        with excel_app_context(output_path) as (app, wb):
            # Mettre à jour la feuille Balises
            try:
                balises_sheet = wb.sheets["Balises"]
                
                # Rechercher et mettre à jour les valeurs
                for param_name, param_value in parameters.items():
                    # Chercher la balise correspondante
                    balise_key = f"[{param_name.title()}]"
                    
                    # Parcourir le tableau Balises
                    for row in range(2, 100):  # Maximum 100 lignes
                        balise_cell = balises_sheet.range(f"A{row}").value
                        if balise_cell and balise_cell.lower() == balise_key.lower():
                            balises_sheet.range(f"C{row}").value = param_value
                            logger.debug(f"Paramètre '{param_name}' = '{param_value}' injecté")
                            break
            except Exception as e:
                logger.warning(f"Erreur mise à jour balises : {e}")
            
            wb.save()
        
        return Path(output_path)
    
    def _load_data(self, excel_path: Path) -> Dict[str, Any]:
        """Charge les données depuis Excel"""
        connector = ExcelConnector(str(excel_path))
        
        data = {}
        for table_name in self.config.data_source.required_tables:
            try:
                df = connector.read_table(table_name)
                data[table_name] = df
                logger.info(f"Table '{table_name}' chargée : {len(df)} lignes")
            except Exception as e:
                logger.warning(f"Impossible de charger '{table_name}' : {e}")
        
        return data
    
    def _generate_powerpoint(self, excel_path: Path, output_path: str, parameters: Dict[str, Any]) -> Path:
        """Génère le PowerPoint final"""
        import shutil
        
        # Copier le template PowerPoint
        template_ppt = self.template_dir / "master.pptx"
        shutil.copy2(template_ppt, output_path)
        
        logger.info(f"PowerPoint copié : {output_path}")
        
        # Charger les balises depuis Excel
        replacements = load_replacement_tags(str(excel_path))
        logger.info(f"{len(replacements)} balises chargées")
        
        # Remplacer les balises dans PowerPoint
        with powerpoint_app_context(output_path, visible=False) as (ppt_app, presentation):
            # Remplacement dans toutes les slides
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    replace_tags_in_shape(shape, replacements)
            
            # Supprimer les slides avec [@SUPR@]
            removed_slides = check_and_remove_suppressed_slides(presentation)
            if removed_slides:
                logger.info(f"Slides supprimées : {', '.join(removed_slides)}")
            
            presentation.Save()
        
        return Path(output_path)
    
    def _inject_images(self, ppt_path: Path, excel_path: Path) -> None:
        """Injecte les images dynamiques"""
        if not self.config.image_injections:
            logger.info("Aucune image à injecter")
            return
        
        replacements = load_replacement_tags(str(excel_path))
        
        with powerpoint_app_context(str(ppt_path), visible=False) as (ppt_app, presentation):
            for slide_id, images_config in self.config.image_injections.items():
                # Trouver la slide
                slide = find_slide_by_id(presentation, slide_id)
                if not slide:
                    logger.warning(f"Slide {slide_id} non trouvée pour injection d'images")
                    continue
                
                # Injecter chaque image
                for img_config in images_config:
                    try:
                        inject_image_to_slide(slide, img_config, replacements)
                        logger.info(f"Image injectée dans slide {slide_id}")
                    except Exception as e:
                        logger.warning(f"Erreur injection image dans {slide_id} : {e}")
            
            presentation.Save()