"""
Batch Processor - Gestion du traitement par batch pour les slides dynamiques
Adapté de spirits_study pour KAIVAA Builder
"""

from typing import Dict, List, Optional, Any, Callable
from dataclasses import dataclass
from loguru import logger

from backend.core.excel_handler import (
    excel_app_context,
    read_loop_table_count
)


@dataclass
class SlideAxis:
    """
    Configuration d'un axe de données pour la génération de slides.
    
    Attributes:
        name: Nom de l'axe (ex: "produits", "distributeurs")
        loop_id: ID dans le tableau Loop (ex: "Produit")
        slides: Liste des slide IDs concernés (ex: ["A003", "A004"])
        sheet_name: Feuille Excel contenant le tableau Loop
    """
    name: str
    loop_id: str
    slides: List[str]
    sheet_name: str = "Charts_settings"


@dataclass
class BatchResult:
    """
    Résultat du traitement d'un batch.
    
    Attributes:
        axis_name: Nom de l'axe traité
        parameter_value: Valeur du paramètre traité
        slide_count: Nombre de slides générées
        replacements: Balises utilisées
        success: Succès ou échec
        error_message: Message d'erreur si échec
    """
    axis_name: str
    parameter_value: int
    slide_count: int
    replacements: Dict[str, str]
    success: bool
    error_message: Optional[str] = None


class BatchProcessor:
    """
    Processeur pour traiter les slides par batch selon leurs axes de données.
    Optimise les performances en regroupant les opérations Excel.
    """
    
    def __init__(self, excel_path: str):
        """
        Initialise le batch processor.
        
        Args:
            excel_path: Chemin vers le fichier Excel
        """
        self.excel_path = excel_path
        self.results: List[BatchResult] = []
    
    def validate_axis_config(self, axes: Dict[str, SlideAxis]) -> bool:
        """
        Valide la configuration des axes.
        
        Args:
            axes: Configuration des axes à valider
            
        Returns:
            True si valide, False sinon
        """
        if not axes:
            logger.error("Configuration des axes vide")
            return False
        
        errors = []
        
        for axis_name, axis in axes.items():
            if not isinstance(axis, SlideAxis):
                errors.append(f"Axe {axis_name} doit être une instance de SlideAxis")
                continue
                
            if not axis.slides:
                errors.append(f"Axe {axis_name} : liste de slides vide")
                
            if not axis.loop_id:
                errors.append(f"Axe {axis_name} : loop_id manquant")
        
        if errors:
            logger.error("Erreurs configuration axes :")
            for error in errors:
                logger.error(f"  - {error}")
            return False
        
        logger.info(f"Configuration {len(axes)} axes validée")
        return True
    
    def get_axis_parameters_count(self, axis: SlideAxis) -> Optional[int]:
        """
        Récupère le nombre de paramètres à traiter pour un axe depuis le tableau Loop.
        
        Args:
            axis: Configuration de l'axe
            
        Returns:
            Nombre de paramètres, ou None si erreur
        """
        try:
            count = read_loop_table_count(self.excel_path, axis.sheet_name, axis.loop_id)
            if count is None or count <= 0:
                logger.warning(f"Aucun paramètre à traiter pour l'axe {axis.name}")
                return None
            return count
        except Exception as e:
            logger.error(f"Erreur lecture count pour axe {axis.name} : {e}")
            return None
    
    def process_axis_batch(
        self, 
        axis: SlideAxis, 
        processor_callback: Callable[[int, Dict[str, str], List[str]], None]
    ) -> List[BatchResult]:
        """
        Traite tous les paramètres d'un axe en batch optimisé.
        
        Args:
            axis: Configuration de l'axe
            processor_callback: Fonction appelée pour chaque paramètre
                               Signature: (parameter_value, replacements, slide_ids)
        
        Returns:
            Liste des résultats de batch
        """
        import time
        
        logger.info(f"Traitement par batch de l'axe '{axis.name}'")
        
        param_count = self.get_axis_parameters_count(axis)
        if param_count is None:
            return []
        
        logger.info(f"  → {param_count} paramètres à traiter pour {len(axis.slides)} slides")
        
        batch_results = []
        
        # Ouverture d'Excel UNE SEULE FOIS pour tout l'axe
        with excel_app_context(self.excel_path) as (app, wb):
            for param_value in range(1, param_count + 1):
                try:
                    logger.debug(f"    → Paramètre {param_value}/{param_count}")
                    
                    # Injection du paramètre dans la session ouverte
                    self._update_loop_table_iteration_in_session(
                        wb, axis.sheet_name, axis.loop_id, param_value
                    )
                    
                    # Forcer le recalcul Excel
                    wb.app.calculate()
                    wb.save()
                    
                    # Délai pour laisser PowerPoint récupérer les changements
                    time.sleep(1.0)
                    
                    # Lecture des balises fraîches
                    replacements = self._read_replacement_tags_from_session(wb)
                    
                    # Appel du callback pour traiter les slides
                    try:
                        processor_callback(param_value, replacements, axis.slides)
                        
                        batch_results.append(BatchResult(
                            axis_name=axis.name,
                            parameter_value=param_value,
                            slide_count=len(axis.slides),
                            replacements=replacements,
                            success=True
                        ))
                    
                    except Exception as callback_error:
                        logger.error(f"Erreur callback pour paramètre {param_value} : {callback_error}")
                        batch_results.append(BatchResult(
                            axis_name=axis.name,
                            parameter_value=param_value,
                            slide_count=len(axis.slides),
                            replacements=replacements,
                            success=False,
                            error_message=str(callback_error)
                        ))
                
                except Exception as e:
                    logger.error(f"Erreur traitement paramètre {param_value} : {e}")
                    batch_results.append(BatchResult(
                        axis_name=axis.name,
                        parameter_value=param_value,
                        slide_count=len(axis.slides),
                        replacements={},
                        success=False,
                        error_message=str(e)
                    ))
        
        successful = len([r for r in batch_results if r.success])
        logger.info(f"Axe '{axis.name}' : {successful}/{len(batch_results)} paramètres traités avec succès")
        
        self.results.extend(batch_results)
        return batch_results
    
    def _update_loop_table_iteration_in_session(
        self, 
        wb, 
        sheet_name: str, 
        loop_id: str, 
        iteration_value: int
    ) -> None:
        """
        Met à jour la valeur d'itération dans le tableau Loop dans une session Excel ouverte.
        
        Args:
            wb: Workbook Excel ouvert
            sheet_name: Nom de la feuille
            loop_id: ID de la boucle
            iteration_value: Nouvelle valeur d'itération
        """
        try:
            sht = wb.sheets[sheet_name]
            
            # Recherche du tableau Loop
            table = None
            for t in sht.api.ListObjects:
                if t.Name.strip().lower() == "loop":
                    table = t
                    break
            
            if not table:
                raise ValueError(f"Tableau 'Loop' introuvable dans '{sheet_name}'")
            
            # Recherche et mise à jour de la ligne
            for row in table.DataBodyRange.Rows:
                id_value = row.Columns(1).Value
                if id_value and str(id_value).strip() == loop_id:
                    row.Columns(2).Value = iteration_value
                    logger.debug(f"Loop {loop_id} mis à jour : itération = {iteration_value}")
                    return
            
            raise ValueError(f"Loop ID '{loop_id}' non trouvé dans le tableau Loop")
        
        except Exception as e:
            raise RuntimeError(f"Erreur mise à jour tableau Loop : {e}")
    
    def _read_replacement_tags_from_session(
        self, 
        wb, 
        sheet_name: str = "Balises", 
        table_name: str = "Balises"
    ) -> Dict[str, str]:
        """
        Lit les balises depuis une session Excel déjà ouverte.
        
        Args:
            wb: Workbook Excel ouvert
            sheet_name: Nom de la feuille
            table_name: Nom du tableau
            
        Returns:
            Dict des balises
        """
        try:
            sht = wb.sheets[sheet_name]
        except Exception:
            raise ValueError(f"Feuille '{sheet_name}' introuvable")
        
        # Recherche du tableau
        table = None
        try:
            for t in sht.api.ListObjects:
                if t.Name.strip().lower() == table_name.lower():
                    table = t
                    break
        except Exception as e:
            raise RuntimeError(f"Erreur accès tableaux : {e}")
        
        if not table:
            raise ValueError(f"Tableau '{table_name}' introuvable dans '{sheet_name}'")
        
        replacements = {}
        try:
            for row in table.DataBodyRange.Rows:
                balise = row.Columns(1).Value
                valeur = row.Columns(3).Text
                if balise and valeur is not None:
                    replacements[balise] = str(valeur)
        except Exception as e:
            raise RuntimeError(f"Erreur lecture balises : {e}")
        
        return replacements
    
    def get_processing_summary(self) -> Dict[str, Any]:
        """
        Retourne un résumé du traitement effectué.
        
        Returns:
            Dict avec statistiques du traitement
        """
        if not self.results:
            return {"total": 0, "success": 0, "errors": 0, "axes": []}
        
        total = len(self.results)
        success = len([r for r in self.results if r.success])
        errors = total - success
        
        axes_summary = {}
        for result in self.results:
            if result.axis_name not in axes_summary:
                axes_summary[result.axis_name] = {"total": 0, "success": 0}
            axes_summary[result.axis_name]["total"] += 1
            if result.success:
                axes_summary[result.axis_name]["success"] += 1
        
        return {
            "total": total,
            "success": success, 
            "errors": errors,
            "success_rate": round(success / total * 100, 1) if total > 0 else 0,
            "axes": axes_summary
        }


def create_slide_axes_from_config(config: Dict[str, Dict[str, Any]]) -> Dict[str, SlideAxis]:
    """
    Convertit une configuration dictionnaire en objets SlideAxis.
    
    Args:
        config: Configuration au format dict
        
    Returns:
        Dict[str, SlideAxis]: Axes configurés
        
    Example:
        config = {
            "produits": {
                "loop_id": "Produit",
                "slides": ["A003", "A004"]
            }
        }
        axes = create_slide_axes_from_config(config)
    """
    axes = {}
    
    for axis_name, axis_config in config.items():
        try:
            axes[axis_name] = SlideAxis(
                name=axis_name,
                loop_id=axis_config["loop_id"],
                slides=axis_config["slides"],
                sheet_name=axis_config.get("sheet_name", "Charts_settings")
            )
        except KeyError as e:
            logger.error(f"Configuration incomplète pour axe {axis_name} : clé manquante {e}")
            continue
        except Exception as e:
            logger.error(f"Erreur création axe {axis_name} : {e}")
            continue
    
    return axes