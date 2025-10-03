"""
Service de génération de rapports
"""
import os
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
from zoneinfo import ZoneInfo


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
    
    def _now(self):
        """Datetime en Europe/Paris (évite les décalages si la machine est en UTC)."""
        return datetime.now(ZoneInfo("Europe/Paris"))

    def generate_report(
        self,
        parameters: Dict[str, Any],
        output_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Génère un rapport complet."""
        logger.info(f"Génération du rapport '{self.config.name}'")
        logger.info(f"Paramètres : {parameters}")
        
        self._validate_parameters(parameters)
        
        # Nettoyage préventif
        from backend.utils.cleanup import cleanup_before_run
        cleanup_before_run()
        
        output_paths = self._generate_output_paths(parameters, output_name)
        ensure_directories(output_paths['excel_path'], output_paths['pptx_path'])
        
        start_time = self._now()
        
        try:
            logger.info("Étape 1/6 : Préparation Excel")
            excel_path = self._prepare_excel(parameters, output_paths['excel_path'])
            
            logger.info("Étape 2/6 : Lecture des données")
            data = self._load_data(excel_path)
            
            logger.info("Étape 3/6 : Génération PowerPoint")
            ppt_path = self._generate_powerpoint(excel_path, output_paths['pptx_path'], parameters)

            logger.info("Conversion des graphiques statiques")
            self._convert_static_charts(ppt_path, excel_path)  

            logger.info("Étape 4/6 : Application des boucles")
            self._apply_loops(ppt_path, excel_path)
            
            logger.info("Étape 5/6 : Injection des tableaux")
            self._inject_tables_to_slides(ppt_path, excel_path)
            
            logger.info("Étape 6/6 : Injection des images")
            self._inject_images(ppt_path, excel_path)
            
            execution_time = (self._now() - start_time).total_seconds()
            
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
            execution_time = (self._now() - start_time).total_seconds()
            
            return {
                "success": False,
                "error": str(e),
                "execution_time_seconds": execution_time,
                "parameters": parameters
            }
        
        finally:
            # Forcer fermeture Excel/PowerPoint
            try:
                import time
                time.sleep(1)
                os.system("taskkill /f /im excel.exe 2>nul 1>nul")
                logger.debug("Excel fermé en fin de génération")
            except:
                pass


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
            param_values = "_".join([str(v) for v in parameters.values()]) if parameters else ""
            timestamp = self._now().strftime("%Y%m%d_%H%M")
            base_name = f"{self.config.name}"
            if param_values:
                base_name += f"_{param_values}"
            base_name += f"_{timestamp}"
        
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
        """Génère le PowerPoint final en préservant les slides qui seront bouclées"""
        import shutil
        import os
        
        template_ppt = self.template_dir / "master.pptx"
        shutil.copy2(template_ppt, output_path)
        
        logger.info(f"PowerPoint copié : {output_path}")
        # 🔗 Relinker le PPT de sortie vers l'Excel de sortie
        relinked = self._relink_excel_links_in_ppt(output_path, excel_path)
        logger.info(f"{relinked} lien(s) Excel relinkés vers l'Excel de sortie")

        
        replacements = load_replacement_tags(str(excel_path))
        logger.info(f"{len(replacements)} balises chargées")
        
        # Identifier les slides qui seront bouclées
        loop_slide_ids = set()
        for loop in self.config.loops:
            loop_slide_ids.update(loop.slides)
        
        logger.info(f"Slides loop à ignorer : {loop_slide_ids}")
        
        with powerpoint_app_context(output_path, visible=True) as (ppt_app, presentation):
            
            # Identifier les INDEX des slides loop
            loop_slide_indices = set()
            for slide_id in loop_slide_ids:
                slide = find_slide_by_id(presentation, slide_id)
                if slide:
                    loop_slide_indices.add(slide.SlideIndex)
                    logger.debug(f"Slide loop {slide_id} trouvée à index {slide.SlideIndex}")
            
            # Remplacer les balises SAUF pour les slides loop
            static_slides_processed = 0
            loop_slides_skipped = 0
            
            for slide in presentation.Slides:
                if slide.SlideIndex in loop_slide_indices:
                    loop_slides_skipped += 1
                    logger.debug(f"Slide index {slide.SlideIndex} ignorée (sera bouclée)")
                else:
                    for shape in slide.Shapes:
                        replace_tags_in_shape(shape, replacements)
                    static_slides_processed += 1
            
            logger.info(f"Balises remplacées : {static_slides_processed} slides statiques, {loop_slides_skipped} slides bouclées préservées")
            
            # Supprimer les slides [@SUPR@]
            removed_slides = check_and_remove_suppressed_slides(presentation)
            if removed_slides:
                logger.info(f"Slides supprimées : {', '.join(removed_slides)}")
            
            presentation.Save()
        
        return Path(output_path)
    
    def _apply_loops(self, ppt_path: Path, excel_path: Path) -> None:
        """Applique les boucles en gardant Excel ET PowerPoint ouverts simultanément"""
        if not self.config.loops:
            logger.info("Aucune boucle configurée")
            return
        
        logger.info(f"Application de {len(self.config.loops)} boucle(s)")
        
        import time
        
        for loop_config in self.config.loops:
            logger.info(f"Traitement boucle '{loop_config.loop_id}'")
            
            param_count = self._read_loop_count(excel_path, loop_config)
            
            if not param_count or param_count <= 0:
                logger.warning(f"Aucune itération pour boucle '{loop_config.loop_id}'")
                continue
            
            logger.info(f"  → {param_count} itérations pour slides {loop_config.slides}")
            
            # Ouvrir Excel UNE SEULE FOIS
            try:
                with excel_app_context(str(excel_path)) as (app, wb):
                    # Ouvrir PowerPoint
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
                        for iteration in range(1, param_count + 1):
                            logger.debug(f"    → Itération {iteration}/{param_count}")
                            
                            # 1. Mettre à jour Excel (qui reste ouvert)
                            self._update_loop_iteration_with_wb(wb, loop_config, iteration)
                            time.sleep(0.5)
                            
                            # 2. Lire les balises APRÈS mise à jour
                            replacements = self._load_replacement_tags_from_wb(wb)
                            logger.debug(f"      Balises rechargées pour itération {iteration}")
                            
                            for slide_id, slide_info in source_slides.items():
                                source_slide = slide_info['slide']
                                original_index = slide_info['original_index']
                                
                                # 3. Rafraîchir les graphiques de la SLIDE SOURCE (Excel ouvert)
                                self._refresh_chart_links_in_slide_live(source_slide)
                                
                                # 4. Dupliquer APRÈS rafraîchissement
                                new_slide = source_slide.Duplicate().Item(1)
                                
                                # Position cible
                                target_position = original_index + (iteration - 1)
                                if target_position <= presentation.Slides.Count:
                                    new_slide.MoveTo(target_position)
                                
                                logger.debug(f"      Slide {slide_id} créée à position {target_position}")
                                
                                # 5. Remplacer les balises sur la COPIE
                                for shape in new_slide.Shapes:
                                    replace_tags_in_shape(shape, replacements)
                                
                                # 6. Injecter les images
                                if slide_id in self.config.image_injections:
                                    for img_config in self.config.image_injections[slide_id]:
                                        is_loop_dependent = getattr(img_config, 'loop_dependent', True)
                                        if is_loop_dependent:
                                            try:
                                                inject_image_to_slide(new_slide, img_config.dict(), replacements)
                                                logger.debug(f"      Image injectée dans {slide_id}")
                                            except Exception as e:
                                                logger.warning(f"Erreur injection image : {e}")
                                
                                # 7. Convertir les graphiques de la COPIE en images
                                self._convert_charts_in_slide(new_slide)
                        
                        # Supprimer les slides sources
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
            
            except Exception as e:
                logger.error(f"Erreur dans la boucle : {e}")
                raise
        
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
                        # Convertir en dict pour utilisation avec inject_image_to_slide
                        img_dict = img_config.dict() if hasattr(img_config, 'dict') else img_config
                        inject_image_to_slide(slide, img_dict, replacements)
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

    def _convert_charts_in_slide(self, slide) -> None:
        """Convertit tous les graphiques d'une slide en images PNG"""
        import time
        
        try:
            charts_converted = 0
            shapes_to_process = []
            
            # Collecter tous les graphiques
            for shape in slide.Shapes:
                if hasattr(shape, 'HasChart') and shape.HasChart:
                    shapes_to_process.append(shape)
            
            if not shapes_to_process:
                return
            
            # Convertir chaque graphique
            for shape in shapes_to_process:
                try:
                    # Sauvegarder position et taille
                    left = shape.Left
                    top = shape.Top
                    width = shape.Width
                    height = shape.Height
                    
                    # Copier le graphique
                    shape.Copy()
                    time.sleep(0.2)
                    
                    # Supprimer l'original
                    shape.Delete()
                    
                    # Coller comme image
                    try:
                        slide.Shapes.PasteSpecial(14)  # ppPasteEnhancedMetafile
                    except:
                        try:
                            slide.Shapes.PasteSpecial(2)  # ppPastePicture
                        except:
                            slide.Shapes.Paste()
                    
                    # Repositionner
                    try:
                        new_shape = slide.Shapes(slide.Shapes.Count)
                        new_shape.Left = left
                        new_shape.Top = top
                        new_shape.Width = width
                        new_shape.Height = height
                    except:
                        pass
                    
                    charts_converted += 1
                
                except Exception as e:
                    logger.warning(f"Erreur conversion graphique : {e}")
                    continue
            
            if charts_converted > 0:
                logger.debug(f"      {charts_converted} graphiques convertis en images")
        
        except Exception as e:
            logger.error(f"Erreur conversion graphiques slide : {e}")

    def _refresh_chart_links_in_slide(self, slide, excel_path: Path) -> None:
        """
        Rafraîchit RÉELLEMENT les graphiques en forçant Excel à recalculer.
        Cette méthode DOIT être appelée sur la slide source AVANT duplication.
        """
        import time
        
        try:
            charts_refreshed = 0
            
            for shape in slide.Shapes:
                if hasattr(shape, 'HasChart') and shape.HasChart:
                    try:
                        chart = shape.Chart
                        
                        # Méthode 1 : Activer le ChartData (ouvre Excel en arrière-plan)
                        try:
                            chart.ChartData.Activate()
                            charts_refreshed += 1
                        except:
                            pass
                        
                        # Méthode 2 : Forcer le recalcul via le Workbook
                        try:
                            workbook = chart.ChartData.Workbook
                            workbook.Application.Calculate()
                            workbook.Application.CalculateFullRebuild()
                        except:
                            pass
                        
                        # Méthode 3 : Refresh du graphique lui-même
                        try:
                            chart.Refresh()
                        except:
                            pass
                        
                    except Exception as e:
                        logger.debug(f"      Erreur rafraîchissement graphique : {e}")
                        continue
            
            if charts_refreshed > 0:
                logger.debug(f"      {charts_refreshed} graphique(s) rafraîchi(s)")
                # IMPORTANT : Pause pour laisser Excel recalculer
                time.sleep(0.5)
        
        except Exception as e:
            logger.warning(f"Erreur rafraîchissement graphiques : {e}")

    def _update_loop_iteration_with_wb(self, excel_wb, loop_config: LoopConfig, iteration: int) -> None:
        """Met à jour la valeur d'itération dans le tableau Loop avec workbook ouvert"""
        try:
            sheet = excel_wb.sheets[loop_config.sheet_name]
            
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
                    
                    # Forcer le recalcul complet
                    excel_wb.app.calculate()
                    excel_wb.save()
                    
                    logger.debug(f"Loop '{loop_config.loop_id}' itération {iteration} - Excel recalculé")
                    return
        
        except Exception as e:
            logger.error(f"Erreur mise à jour Loop : {e}")

    def _load_replacement_tags_from_wb(self, excel_wb, sheet_name: str = "Balises", table_name: str = "Balises") -> Dict[str, str]:
        """Lit les balises depuis un workbook déjà ouvert"""
        try:
            sht = excel_wb.sheets[sheet_name]
            
            table = None
            for t in sht.api.ListObjects:
                if t.Name.strip().lower() == table_name.lower():
                    table = t
                    break
            
            if not table:
                logger.error(f"Tableau '{table_name}' introuvable")
                return {}
            
            replacements = {}
            data_range = table.DataBodyRange
            if data_range is None:
                return {}
            
            for row in data_range.Rows:
                try:
                    balise = row.Columns(1).Value
                    valeur = row.Columns(3).Text
                    if balise and valeur is not None:
                        replacements[balise] = str(valeur)
                except:
                    continue
            
            logger.debug(f"{len(replacements)} balises lues depuis workbook ouvert")
            return replacements
        
        except Exception as e:
            logger.error(f"Erreur lecture balises : {e}")
            return {}

    def _refresh_chart_links_in_slide_live(self, slide) -> None:
        """
        Rafraîchit les graphiques avec Excel ouvert simultanément.
        CRITIQUE : Excel doit être ouvert pour que les données se mettent à jour.
        """
        import time
        
        try:
            charts_refreshed = 0
            
            for shape in slide.Shapes:
                if hasattr(shape, 'HasChart') and shape.HasChart:
                    try:
                        chart = shape.Chart
                        
                        # Activer le ChartData (ouvre la connexion Excel)
                        chart.ChartData.Activate()
                        
                        # Forcer le refresh
                        chart.Refresh()
                        
                        charts_refreshed += 1
                        
                    except Exception as e:
                        logger.debug(f"      Erreur rafraîchissement graphique : {e}")
                        continue
            
            if charts_refreshed > 0:
                logger.debug(f"      {charts_refreshed} graphique(s) rafraîchi(s)")
                time.sleep(0.5)
        
        except Exception as e:
            logger.warning(f"Erreur rafraîchissement graphiques : {e}")


    def _relink_excel_links_in_ppt(self, ppt_path: Path, excel_path: Path) -> int:
        """Pointe tous les liens Excel du PPT vers l'Excel de sortie."""
        from backend.core.ppt_handler import powerpoint_app_context
        import os
        excel_path_abs = os.path.abspath(str(excel_path))
        relinked = 0
        with powerpoint_app_context(str(ppt_path), visible=True) as (ppt_app, presentation):
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    try:
                        # Cas 1: objets liés (OLE, images liées, etc.)
                        if hasattr(shape, "LinkFormat") and shape.LinkFormat:
                            old_src = shape.LinkFormat.SourceFullName
                            # Ne relinker que s'il s'agit d'un fichier Excel
                            if old_src and (old_src.lower().endswith(".xlsx") or ".xlsx" in old_src.lower()):
                                shape.LinkFormat.SourceFullName = excel_path_abs
                                relinked += 1
                        # Cas 2: certains graphiques natives PPT avec data link externe (rare)
                        if hasattr(shape, "HasChart") and shape.HasChart:
                            try:
                                # Si le chart est "linked", certains environnements exposent LinkFormat
                                if hasattr(shape, "LinkFormat") and shape.LinkFormat:
                                    old_src2 = shape.LinkFormat.SourceFullName
                                    if old_src2 and (old_src2.lower().endswith(".xlsx") or ".xlsx" in old_src2.lower()):
                                        shape.LinkFormat.SourceFullName = excel_path_abs
                                        relinked += 1
                            except:
                                pass
                    except:
                        continue
            presentation.Save()
        return relinked

    def _log_chart_sources(self, ppt_path: Path) -> None:
        from backend.core.ppt_handler import powerpoint_app_context
        with powerpoint_app_context(str(ppt_path), visible=False) as (ppt_app, presentation):
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if hasattr(shape, "HasChart") and shape.HasChart:
                        try:
                            chart = shape.Chart
                            # Selon les cas, on peut accéder au classeur
                            wb = chart.ChartData.Workbook
                            if hasattr(wb, "FullName"):
                                from loguru import logger
                                logger.debug(f"Slide {slide.SlideIndex} chart uses workbook: {wb.FullName}")
                        except Exception:
                            pass

    def _convert_static_charts(self, ppt_path: Path, excel_path: Path) -> None:
        """Rafraîchit puis convertit les graphiques des slides statiques en images"""
        if not self.config.loops:
            logger.info("Aucune boucle configurée, conversion de tous les graphiques")
            loop_slide_ids = set()
        else:
            loop_slide_ids = set()
            for loop in self.config.loops:
                loop_slide_ids.update(loop.slides)
        
        logger.info(f"Rafraîchissement et conversion des graphiques statiques (slides loop ignorées : {loop_slide_ids})")
        
        # Ouvrir Excel ET PowerPoint simultanément pour le rafraîchissement
        with excel_app_context(str(excel_path)) as (excel_app, excel_wb):
            with powerpoint_app_context(str(ppt_path), visible=True) as (ppt_app, presentation):
                converted_count = 0
                
                for slide in presentation.Slides:
                    # Vérifier si la slide contient un ID de loop
                    is_loop_slide = False
                    for shape in slide.Shapes:
                        if hasattr(shape, 'HasTextFrame') and shape.HasTextFrame:
                            try:
                                text = shape.TextFrame2.TextRange.Text
                                for slide_id in loop_slide_ids:
                                    if slide_id in text:
                                        is_loop_slide = True
                                        break
                            except:
                                continue
                        if is_loop_slide:
                            break
                    
                    if not is_loop_slide:
                        # 1. Rafraîchir les graphiques de cette slide
                        charts_refreshed = 0
                        for shape in slide.Shapes:
                            if hasattr(shape, 'HasChart') and shape.HasChart:
                                try:
                                    chart = shape.Chart
                                    chart.ChartData.Activate()
                                    chart.Refresh()
                                    charts_refreshed += 1
                                except Exception as e:
                                    logger.debug(f"Erreur rafraîchissement graphique : {e}")
                                    continue
                        
                        if charts_refreshed > 0:
                            logger.debug(f"Slide {slide.SlideIndex} : {charts_refreshed} graphique(s) rafraîchi(s)")
                        
                        # 2. Convertir les graphiques en images
                        self._convert_charts_in_slide(slide)
                        converted_count += 1
                
                logger.info(f"{converted_count} slides statiques avec graphiques rafraîchis et convertis")
                presentation.Save()