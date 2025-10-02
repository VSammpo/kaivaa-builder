"""
Service de génération de rapports
"""

from pathlib import Path
from typing import Dict, Optional, Any, List
from datetime import datetime
from loguru import logger

from backend.config import PathConfig
from backend.models.template_config import TemplateConfig, LoopConfig
from backend.connectors.excel_connector import ExcelConnector
from backend.core.excel_handler import load_replacement_tags, excel_app_context
from backend.core.ppt_handler import (
    powerpoint_app_context,
    replace_tags_in_shape,
    find_slide_by_id,
    check_and_remove_suppressed_slides
)
from backend.core.image_handler import inject_image_to_slide, find_slides_by_ids
from backend.core.batch_processor import BatchProcessor, SlideAxis, create_slide_axes_from_config
from backend.core.chart_handler import ChartExporter
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
            parameters: Paramètres du rapport
            output_name: Nom personnalisé pour les fichiers de sortie
            
        Returns:
            Dict avec chemins des fichiers générés et métadonnées
        """
        logger.info(f"Génération du rapport '{self.config.name}'")
        logger.info(f"Paramètres : {parameters}")
        
        self._validate_parameters(parameters)
        cleanup_before_run()
        
        output_paths = self._generate_output_paths(parameters, output_name)
        ensure_directories(output_paths['excel_path'], output_paths['pptx_path'])
        
        start_time = datetime.now()
        
        try:
            logger.info("Étape 1/8 : Préparation Excel")
            excel_path = self._prepare_excel(parameters, output_paths['excel_path'])
            
            logger.info("Étape 2/8 : Lecture des données")
            data = self._load_data(excel_path)
            
            logger.info("Étape 3/8 : Export des graphiques")
            chart_exporter = ChartExporter(str(excel_path))
            charts_map = chart_exporter.export_all_charts()
            
            logger.info("Étape 4/8 : Génération PowerPoint")
            ppt_path = self._generate_powerpoint(excel_path, output_paths['pptx_path'], parameters)
            
            logger.info("Étape 5/8 : Application des boucles")
            self._apply_loops(ppt_path, excel_path)
            
            logger.info("Étape 6/8 : Injection des tableaux")
            self._inject_tables_to_slides(ppt_path, excel_path)
            
            logger.info("Étape 7/8 : Injection des images")
            self._inject_images(ppt_path, excel_path)
            
            logger.info("Étape 8/8 : Injection des graphiques")
            self._inject_chart_images(ppt_path, charts_map)
            
            chart_exporter.cleanup()
            
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
        
        template_excel = self.template_dir / "master.xlsx"
        shutil.copy2(template_excel, output_path)
        
        logger.info(f"Excel copié : {output_path}")
        
        with excel_app_context(output_path) as (app, wb):
            try:
                balises_sheet = wb.sheets["Balises"]
                
                for param_name, param_value in parameters.items():
                    balise_key = f"[{param_name.title()}]"
                    
                    for row in range(2, 100):
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
        
        template_ppt = self.template_dir / "master.pptx"
        shutil.copy2(template_ppt, output_path)
        
        logger.info(f"PowerPoint copié : {output_path}")
        
        replacements = load_replacement_tags(str(excel_path))
        logger.info(f"{len(replacements)} balises chargées")
        
        with powerpoint_app_context(output_path, visible=True) as (ppt_app, presentation):
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    replace_tags_in_shape(shape, replacements)
            
            removed_slides = check_and_remove_suppressed_slides(presentation)
            if removed_slides:
                logger.info(f"Slides supprimées : {', '.join(removed_slides)}")
            
            presentation.Save()
        
        return Path(output_path)
    
    def _apply_loops(self, ppt_path: Path, excel_path: Path) -> None:
        """Applique les boucles pour dupliquer les slides"""
        if not self.config.loops:
            logger.info("Aucune boucle configurée")
            return
        
        logger.info(f"Application de {len(self.config.loops)} boucle(s)")
        
        from backend.core.batch_processor import BatchProcessor
        import time
        
        processor = BatchProcessor(str(excel_path))
        
        for loop_config in self.config.loops:
            logger.info(f"Traitement boucle '{loop_config.loop_id}'")
            
            param_count = self._read_loop_count(excel_path, loop_config)
            
            if not param_count or param_count <= 0:
                logger.warning(f"Aucune itération pour boucle '{loop_config.loop_id}'")
                continue
            
            logger.info(f"  → {param_count} itérations pour slides {loop_config.slides}")
            
            with powerpoint_app_context(str(ppt_path), visible=True) as (ppt_app, presentation):
                
                # Trouver les slides sources
                source_slides = {}
                for slide_id in loop_config.slides:
                    slide = find_slide_by_id(presentation, slide_id)
                    if slide:
                        source_slides[slide_id] = {
                            'slide': slide,
                            'original_index': slide.SlideIndex
                        }
                
                if not source_slides:
                    logger.error(f"Aucune slide source pour '{loop_config.loop_id}'")
                    continue
                
                # Créer les slides pour chaque itération
                created_slides = []
                
                for iteration in range(1, param_count + 1):
                    logger.debug(f"    → Itération {iteration}/{param_count}")
                    
                    # CORRECTION : Mettre à jour Excel AVANT de lire les balises
                    self._update_loop_iteration(excel_path, loop_config, iteration)
                    
                    # Attendre que Excel recalcule
                    time.sleep(0.5)
                    
                    # CORRECTION : Lire les balises APRÈS mise à jour
                    replacements = load_replacement_tags(str(excel_path))
                    logger.debug(f"      Balises rechargées pour itération {iteration}")
                    
                    for slide_id, slide_info in source_slides.items():
                        source_slide = slide_info['slide']
                        original_index = slide_info['original_index']
                        
                        # CORRECTION : Toujours dupliquer (même pour iteration 1)
                        new_slide = source_slide.Duplicate().Item(1)
                        
                        # Position cible
                        target_position = original_index + (iteration - 1)
                        
                        if target_position <= presentation.Slides.Count:
                            new_slide.MoveTo(target_position)
                        
                        created_slides.append(new_slide)
                        logger.debug(f"      Slide {slide_id} créée à position {target_position}")
                        
                        # Remplacer les balises avec les valeurs de CETTE itération
                        for shape in new_slide.Shapes:
                            replace_tags_in_shape(shape, replacements)
                        
                        # Injecter les images si configurées
                        if slide_id in self.config.image_injections:
                            for img_config in self.config.image_injections[slide_id]:
                                # Vérifier si l'attribut existe (compatibilité)
                                is_loop_dependent = getattr(img_config, 'loop_dependent', True)
                                if is_loop_dependent:
                                    try:
                                        inject_image_to_slide(new_slide, img_config.dict(), replacements)
                                        logger.debug(f"      Image injectée dans {slide_id}")
                                    except Exception as e:
                                        logger.warning(f"Erreur injection image : {e}")
                
                # CORRECTION : Supprimer les slides sources APRÈS avoir créé toutes les itérations
                logger.info(f"  → Suppression de {len(source_slides)} slide(s) source(s)")
                for slide_id, slide_info in sorted(source_slides.items(), 
                                                   key=lambda x: x[1]['slide'].SlideIndex, 
                                                   reverse=True):
                    try:
                        slide_info['slide'].Delete()
                        logger.debug(f"    Slide source {slide_id} supprimée")
                    except Exception as e:
                        logger.warning(f"Erreur suppression {slide_id} : {e}")
                
                presentation.Save()
        
        logger.success("Boucles appliquées avec succès")
    
    def _read_loop_count(self, excel_path: Path, loop_config: LoopConfig) -> Optional[int]:
        """Lit le nombre d'itérations depuis le tableau Loop"""
        try:
            with excel_app_context(str(excel_path), visible=False, read_only=True) as (app, wb):
                sheet = wb.sheets[loop_config.sheet_name]
                
                # Chercher le tableau Loop
                table = None
                for t in sheet.api.ListObjects:
                    if t.Name.strip().lower() == "loop":
                        table = t
                        break
                
                if not table:
                    logger.error(f"Tableau 'Loop' introuvable dans '{loop_config.sheet_name}'")
                    return None
                
                # Chercher la ligne correspondant au loop_id
                for row in table.DataBodyRange.Rows:
                    id_value = row.Columns(1).Value
                    if id_value and str(id_value).strip() == loop_config.loop_id:
                        count_value = row.Columns(3).Value  # Colonne "Nombre de tests"
                        return int(count_value) if count_value else 0
                
                logger.error(f"Loop ID '{loop_config.loop_id}' non trouvé dans tableau Loop")
                return None
        
        except Exception as e:
            logger.error(f"Erreur lecture Loop : {e}")
            return None
    
    def _update_loop_iteration(self, excel_path: Path, loop_config: LoopConfig, iteration: int) -> None:
        """Met à jour la valeur d'itération dans le tableau Loop"""
        try:
            with excel_app_context(str(excel_path)) as (app, wb):
                sheet = wb.sheets[loop_config.sheet_name]
                
                table = None
                for t in sheet.api.ListObjects:
                    if t.Name.strip().lower() == "loop":
                        table = t
                        break
                
                if not table:
                    return
                
                for row in table.DataBodyRange.Rows:
                    id_value = row.Columns(1).Value
                    if id_value and str(id_value).strip() == loop_config.loop_id:
                        row.Columns(2).Value = iteration
                        
                        # CORRECTION : Forcer le recalcul complet
                        wb.app.calculate()
                        wb.save()
                        
                        logger.debug(f"Loop '{loop_config.loop_id}' itération {iteration} - Excel recalculé")
                        return
        
        except Exception as e:
            logger.error(f"Erreur mise à jour Loop : {e}")
    
    def _inject_images(self, ppt_path: Path, excel_path: Path) -> None:
        """Injecte les images dynamiques"""
        if not self.config.image_injections:
            logger.info("Aucune image à injecter")
            return
        
        replacements = load_replacement_tags(str(excel_path))
        
        with powerpoint_app_context(str(ppt_path), visible=True) as (ppt_app, presentation):
            for slide_id, images_config in self.config.image_injections.items():
                slide = find_slide_by_id(presentation, slide_id)
                if not slide:
                    logger.warning(f"Slide {slide_id} non trouvée pour injection d'images")
                    continue
                
                for img_config in images_config:
                    try:
                        inject_image_to_slide(slide, img_config, replacements)
                        logger.info(f"Image injectée dans slide {slide_id}")
                    except Exception as e:
                        logger.warning(f"Erreur injection image dans {slide_id} : {e}")
            
            presentation.Save()
    
    def _inject_tables_to_slides(self, ppt_path: Path, excel_path: Path) -> None:
        """Injecte les données Excel dans les tableaux PowerPoint"""
        if not self.config.slide_mappings:
            logger.info("Aucun mapping de tableau configuré")
            return
        
        logger.info(f"Injection de {len(self.config.slide_mappings)} tableau(x)")
        
        from backend.core.excel_handler import read_excel_range_data
        
        with powerpoint_app_context(str(ppt_path), visible=True) as (ppt_app, presentation):
            for mapping in self.config.slide_mappings:
                slide = find_slide_by_id(presentation, mapping.slide_id)
                
                if not slide:
                    logger.warning(f"Slide {mapping.slide_id} non trouvée pour mapping")
                    continue
                
                # Lire les données Excel
                try:
                    data_text, hyperlinks_data = read_excel_range_data(
                        str(excel_path), 
                        mapping.sheet_name, 
                        mapping.excel_range
                    )
                except Exception as e:
                    logger.error(f"Erreur lecture {mapping.sheet_name}!{mapping.excel_range} : {e}")
                    continue
                
                if not data_text:
                    logger.warning(f"Aucune donnée pour {mapping.slide_id}")
                    continue
                
                # Trouver le tableau dans la slide
                table_shape = None
                for shape in slide.Shapes:
                    if hasattr(shape, 'HasTable') and shape.HasTable:
                        table_shape = shape
                        break
                
                if not table_shape:
                    logger.warning(f"Aucun tableau dans slide {mapping.slide_id}")
                    continue
                
                # Injection
                try:
                    self._inject_data_to_table(
                        table_shape.Table, 
                        data_text, 
                        mapping.has_header,
                        hyperlinks_data
                    )
                    logger.info(f"Tableau injecté dans {mapping.slide_id}")
                except Exception as e:
                    logger.error(f"Erreur injection tableau {mapping.slide_id} : {e}")
            
            presentation.Save()
    
    def _inject_data_to_table(self, table, data: list, has_header: bool, hyperlinks_data: dict = None) -> None:
        """Injecte des données dans un tableau PowerPoint"""
        offset = 1 if has_header else 0
        n_rows = min(len(data), table.Rows.Count - offset)
        n_cols = min(len(data[0]), table.Columns.Count) if n_rows > 0 else 0
        
        for r in range(n_rows):
            for c in range(n_cols):
                try:
                    value = data[r][c] if data[r][c] else ""
                    cell_shape = table.Cell(r + 1 + offset, c + 1).Shape
                    text_range = cell_shape.TextFrame2.TextRange
                    text_range.Text = str(value)
                    
                    # Hyperliens
                    if hyperlinks_data and (r, c) in hyperlinks_data:
                        url = hyperlinks_data[(r, c)]["url"]
                        try:
                            text_range.ActionSettings[1].Hyperlink.Address = url
                        except:
                            pass
                except:
                    continue
    
    def _inject_chart_images(self, ppt_path: Path, charts_map: Dict[str, List[str]]) -> None:
        """Injecte les graphiques Excel exportés comme images dans PowerPoint"""
        if not charts_map:
            logger.info("Aucun graphique à injecter")
            return
        
        total_charts = sum(len(charts) for charts in charts_map.values())
        logger.info(f"Injection de {total_charts} graphique(s)")
        
        with powerpoint_app_context(str(ppt_path), visible=True) as (ppt_app, presentation):
            all_chart_images = []
            for sheet_name, chart_paths in charts_map.items():
                all_chart_images.extend(chart_paths)
            
            replaced_count = 0
            
            for slide in presentation.Slides:
                shapes_to_process = []
                for shape in slide.Shapes:
                    if hasattr(shape, 'HasChart') and shape.HasChart:
                        shapes_to_process.append(shape)
                
                for shape in shapes_to_process:
                    if not all_chart_images:
                        break
                    
                    try:
                        left = shape.Left
                        top = shape.Top
                        width = shape.Width
                        height = shape.Height
                        
                        chart_image = all_chart_images.pop(0)
                        
                        shape.Delete()
                        
                        slide.Shapes.AddPicture(
                            FileName=chart_image,
                            LinkToFile=False,
                            SaveWithDocument=True,
                            Left=left,
                            Top=top,
                            Width=width,
                            Height=height
                        )
                        
                        replaced_count += 1
                        logger.debug(f"Graphique remplacé par image PNG")
                    except Exception as e:
                        logger.warning(f"Erreur remplacement graphique : {e}")
            
            logger.info(f"{replaced_count}/{total_charts} graphiques remplacés")
            presentation.Save()