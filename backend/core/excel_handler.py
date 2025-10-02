"""
Excel Handler - Gestion des opérations Excel via xlwings
Adapté de spirits_study pour KAIVAA Builder
"""

import os
from pathlib import Path
from contextlib import contextmanager
from typing import Dict, Optional, List, Tuple, Any
import xlwings as xw
import pandas as pd
from loguru import logger


@contextmanager
def excel_app_context(path: str, visible: bool = False, read_only: bool = False):
    """
    Context manager pour gérer proprement les ressources Excel.
    
    Args:
        path: Chemin vers le fichier Excel
        visible: Si True, affiche l'application Excel
        read_only: Si True, ouvre en lecture seule
        
    Yields:
        Tuple[xlwings.App, xlwings.Book]: Application et workbook Excel
    """
    app = None
    wb = None
    try:
        logger.debug(f"Ouverture Excel: {path}")
        app = xw.App(visible=visible)
        wb = app.books.open(path, read_only=read_only)  # <- AJOUTER read_only ici
        yield app, wb
    except Exception as e:
        logger.error(f"Erreur ouverture Excel: {e}")
        raise RuntimeError(f"Erreur lors de l'ouverture d'Excel ({path}): {e}")
    finally:
        if wb is not None:
            try:
                wb.close()
                logger.debug("Workbook fermé")
            except:
                pass
        if app is not None:
            try:
                app.quit()
                logger.debug("Application Excel fermée")
            except:
                pass

def copy_template_excel(template_path: str, dest_path: str) -> None:
    """
    Copie un template Excel vers un chemin de travail.
    
    Args:
        template_path: Chemin du template source
        dest_path: Chemin de destination
        
    Raises:
        FileNotFoundError: Si le template n'existe pas
    """
    if not os.path.exists(template_path):
        logger.error(f"Template Excel introuvable: {template_path}")
        raise FileNotFoundError(f"Template Excel introuvable : {template_path}")
    
    from shutil import copyfile
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    copyfile(template_path, dest_path)
    logger.info(f"Template copié: {os.path.basename(dest_path)}")


def inject_filter_values(
    excel_path: str, 
    values: Dict[str, Any], 
    sheet_name: str = "Charts_settings"
) -> None:
    """
    Injecte des valeurs dans des cellules spécifiques d'Excel.
    
    Args:
        excel_path: Chemin du fichier Excel
        values: Dictionnaire {cellule: valeur} ex: {"C2": "Leclerc"}
        sheet_name: Nom de la feuille cible
        
    Example:
        inject_filter_values("data.xlsx", {"C2": "Paris", "C3": "2025"})
    """
    if not values:
        logger.debug("Aucune valeur à injecter")
        return
    
    logger.debug(f"Injection dans {sheet_name}: {values}")
    
    with excel_app_context(excel_path) as (app, wb):
        try:
            sheet = wb.sheets[sheet_name]
            logger.debug(f"Feuille '{sheet_name}' trouvée")
        except Exception:
            available_sheets = [ws.name for ws in wb.sheets]
            logger.error(f"Feuille '{sheet_name}' introuvable. Disponibles: {available_sheets}")
            raise ValueError(f"Feuille '{sheet_name}' introuvable dans {excel_path}")
        
        for cell, val in values.items():
            try:
                sheet.range(cell).value = val
                logger.debug(f"  {cell} = {val}")
            except Exception as e:
                logger.error(f"Erreur injection {cell}: {e}")
                raise ValueError(f"Erreur lors de l'injection en cellule {cell} : {e}")
        
        wb.save()
        logger.info(f"Valeurs injectées: {len(values)} cellules")


def load_replacement_tags(
    excel_path: str, 
    sheet_name: str = "Balises", 
    table_name: str = "Balises"
) -> Dict[str, str]:
    """
    Lit les balises de remplacement depuis un tableau structuré Excel.
    
    Args:
        excel_path: Chemin du fichier Excel
        sheet_name: Nom de la feuille contenant les balises
        table_name: Nom du tableau structuré
        
    Returns:
        Dict des balises {balise: valeur}
        
    Example:
        tags = load_replacement_tags("data.xlsx")
        # {"[Marque]": "BOMBAY", "[Segment]": "Gin"}
    """
    max_retries = 3
    
    for attempt in range(max_retries):
        try:
            if attempt > 0:
                _force_close_excel_instances()
                import time
                time.sleep(2.0)
            
            with excel_app_context(excel_path) as (app, wb):
                try:
                    sht = wb.sheets[sheet_name]
                except Exception:
                    raise ValueError(f"Feuille '{sheet_name}' introuvable")
                
                if not hasattr(sht, 'api') or sht.api is None:
                    raise RuntimeError("API Excel corrompue")
                
                # Recherche du tableau
                table = None
                try:
                    list_objects = sht.api.ListObjects
                    if list_objects is None:
                        raise ValueError(f"Aucun tableau accessible dans '{sheet_name}'")
                    
                    try:
                        table = list_objects(table_name)
                        if table:
                            logger.debug(f"Tableau '{table_name}' trouvé par accès direct")
                    except:
                        for t in list_objects:
                            if t.Name.strip().lower() == table_name.lower():
                                table = t
                                logger.debug(f"Tableau '{table_name}' trouvé par itération")
                                break
                
                except Exception as list_error:
                    raise RuntimeError(f"Erreur accès tableaux : {list_error}")
                
                if not table:
                    raise ValueError(f"Tableau '{table_name}' introuvable dans '{sheet_name}'")
                
                # Lecture des balises
                replacements = {}
                try:
                    data_range = table.DataBodyRange
                    if data_range is None:
                        logger.warning(f"Tableau '{table_name}' vide")
                        return {}
                    
                    for row in data_range.Rows:
                        try:
                            balise = row.Columns(1).Value
                            valeur = row.Columns(3).Text
                            if balise and valeur is not None:
                                replacements[balise] = str(valeur)
                        except Exception:
                            continue
                
                except Exception as e:
                    raise RuntimeError(f"Erreur lecture balises : {e}")
                
                logger.info(f"{len(replacements)} balises chargées")
                return replacements
        
        except Exception as e:
            error_msg = str(e).lower()
            is_com_error = any(keyword in error_msg for keyword in [
                "enumeration", "rejeté", "rejected", "automation", "com_error"
            ])
            
            if is_com_error and attempt < max_retries - 1:
                logger.warning(f"Tentative {attempt + 1} échouée (erreur COM), retry...")
                continue
            else:
                logger.error(f"Erreur lecture balises après {attempt + 1} tentatives")
                raise
    
    return {}


def read_excel_range_data(
    excel_path: str, 
    sheet_name: str, 
    excel_range: str
) -> Tuple[List[List[str]], Dict[Tuple[int, int], Dict[str, str]]]:
    """
    Lit les données d'une plage Excel avec les hyperliens.
    
    Args:
        excel_path: Chemin du fichier Excel
        sheet_name: Nom de la feuille
        excel_range: Plage à lire (ex: "A1:C10")
        
    Returns:
        Tuple (data_text, hyperlinks_data)
        - data_text: Liste de listes avec le texte formaté
        - hyperlinks_data: Dict {(row, col): {"text": str, "url": str}}
    """
    with excel_app_context(excel_path) as (app, wb):
        try:
            sht = wb.sheets[sheet_name]
        except Exception:
            raise ValueError(f"Feuille '{sheet_name}' introuvable")
        
        try:
            range_obj = sht.range(excel_range)
            data = range_obj.api.Value2
            
            if not data:
                return [], {}
            
            hyperlinks_data = {}
            
            # Gestion dimensions
            if isinstance(data, (list, tuple)) and len(data) > 0:
                if isinstance(data[0], (list, tuple)):
                    num_rows = len(data)
                    num_cols = len(data[0]) if data[0] else 0
                else:
                    num_rows = 1
                    num_cols = len(data)
                    data = [data]
            else:
                num_rows = 1
                num_cols = 1
                data = [[data]]
            
            # Lecture hyperliens et texte formaté
            data_text = []
            for r in range(num_rows):
                row_data = []
                for c in range(num_cols):
                    cell = range_obj.api.Cells(r + 1, c + 1)
                    formatted_text = str(cell.Text)
                    row_data.append(formatted_text)
                    
                    # Vérification formules HYPERLINK
                    try:
                        formula = str(cell.Formula)
                        if formula and "HYPERLINK(" in formula.upper():
                            url = _extract_url_from_hyperlink_formula(formula)
                            if url:
                                hyperlinks_data[(r, c)] = {
                                    "text": formatted_text,
                                    "url": url
                                }
                                continue
                    except:
                        pass
                    
                    # Hyperliens natifs Excel
                    try:
                        if cell.Hyperlinks.Count > 0:
                            hyperlink = cell.Hyperlinks.Item(1)
                            hyperlinks_data[(r, c)] = {
                                "text": formatted_text,
                                "url": hyperlink.Address
                            }
                    except:
                        pass
                
                data_text.append(row_data)
            
            logger.debug(f"Lecture {excel_range}: {num_rows}x{num_cols}, {len(hyperlinks_data)} hyperliens")
            return data_text, hyperlinks_data
        
        except Exception as e:
            raise RuntimeError(f"Erreur lecture plage {excel_range} : {e}")


def read_loop_table_count(excel_path: str, sheet_name: str, loop_id: str) -> Optional[int]:
    """
    Lit le nombre de tests depuis le tableau Loop pour un ID donné.
    
    Args:
        excel_path: Chemin du fichier Excel
        sheet_name: Nom de la feuille contenant le tableau Loop
        loop_id: ID de la boucle (ex: "Produit", "Concurrent")
        
    Returns:
        Nombre de tests ou None si non trouvé
    """
    max_retries = 3
    
    for attempt in range(max_retries):
        try:
            if attempt > 0:
                _force_close_excel_instances()
                import time
                time.sleep(2.0)
            
            with excel_app_context(excel_path) as (app, wb):
                try:
                    sht = wb.sheets[sheet_name]
                except Exception:
                    raise ValueError(f"Feuille '{sheet_name}' introuvable")
                
                if not hasattr(sht, 'api') or sht.api is None:
                    raise RuntimeError("API Excel corrompue")
                
                try:
                    table = None
                    list_objects = sht.api.ListObjects
                    
                    if list_objects is None:
                        raise ValueError(f"Aucun tableau accessible dans '{sheet_name}'")
                    
                    try:
                        table = list_objects("Loop")
                    except:
                        for t in list_objects:
                            if t.Name.strip().lower() == "loop":
                                table = t
                                break
                    
                    if not table:
                        available_tables = [t.Name for t in list_objects]
                        raise ValueError(f"Tableau 'Loop' introuvable. Disponibles: {available_tables}")
                    
                    data_range = table.DataBodyRange
                    if data_range is None:
                        logger.warning("Tableau Loop vide")
                        return None
                    
                    for row in data_range.Rows:
                        try:
                            id_value = row.Columns(1).Value
                            count_value = row.Columns(3).Value
                            
                            if id_value and str(id_value).strip() == loop_id:
                                result = int(count_value) if count_value is not None else None
                                logger.debug(f"Loop '{loop_id}' trouvé: {result} tests")
                                return result
                        except Exception:
                            continue
                    
                    logger.warning(f"Loop ID '{loop_id}' non trouvé dans le tableau")
                    return None
                
                except Exception as e:
                    raise RuntimeError(f"Erreur lecture tableau Loop : {e}")
        
        except Exception as e:
            error_msg = str(e).lower()
            is_com_error = any(keyword in error_msg for keyword in [
                "enumeration", "rejeté", "rejected", "automation", "com_error"
            ])
            
            if is_com_error and attempt < max_retries - 1:
                logger.warning(f"Lecture Loop tentative {attempt + 1} échouée, retry...")
                continue
            else:
                logger.error(f"Erreur lecture Loop après {attempt + 1} tentatives")
                return None
    
    return None


def update_loop_table_iteration(
    excel_path: str, 
    sheet_name: str, 
    loop_id: str, 
    iteration_value: int
) -> None:
    """
    Met à jour la valeur d'itération dans le tableau Loop.
    
    Args:
        excel_path: Chemin du fichier Excel
        sheet_name: Nom de la feuille
        loop_id: ID de la boucle
        iteration_value: Nouvelle valeur d'itération
    """
    max_retries = 2
    
    for attempt in range(max_retries):
        try:
            with excel_app_context(excel_path) as (app, wb):
                try:
                    sht = wb.sheets[sheet_name]
                except Exception:
                    raise ValueError(f"Feuille '{sheet_name}' introuvable")
                
                if not hasattr(sht, 'api') or sht.api is None:
                    raise RuntimeError("Session Excel corrompue")
                
                try:
                    table = None
                    list_objects = sht.api.ListObjects
                    
                    if list_objects is None:
                        raise RuntimeError("ListObjects inaccessible")
                    
                    for t in list_objects:
                        if t.Name.strip().lower() == "loop":
                            table = t
                            break
                    
                    if not table:
                        raise ValueError(f"Tableau 'Loop' introuvable dans '{sheet_name}'")
                    
                    updated = False
                    for row in table.DataBodyRange.Rows:
                        id_value = row.Columns(1).Value
                        if id_value and str(id_value).strip() == loop_id:
                            row.Columns(2).Value = iteration_value
                            updated = True
                            break
                    
                    if not updated:
                        raise ValueError(f"Loop ID '{loop_id}' non trouvé")
                    
                    wb.save()
                    logger.debug(f"Loop '{loop_id}' mis à jour: itération={iteration_value}")
                    return
                
                except Exception as e:
                    if attempt < max_retries - 1:
                        logger.warning(f"Tentative {attempt + 1} échouée: {e}")
                        import time
                        time.sleep(1.0)
                        continue
                    else:
                        raise RuntimeError(f"Erreur mise à jour Loop : {e}")
        
        except Exception as e:
            if attempt < max_retries - 1:
                import time
                time.sleep(1.0)
                continue
            else:
                raise


def _extract_url_from_hyperlink_formula(formula: str) -> Optional[str]:
    """Extrait l'URL d'une formule HYPERLINK Excel."""
    import re
    
    try:
        pattern = r'HYPERLINK\s*\(\s*"([^"]+)"[,;]'
        match = re.search(pattern, formula, re.IGNORECASE)
        
        if match:
            return match.group(1).strip()
        
        pattern2 = r'HYPERLINK\s*\(\s*([^,;)]+)[,;)]'
        match2 = re.search(pattern2, formula, re.IGNORECASE)
        
        if match2:
            url = match2.group(1).strip().strip('"\'')
            return url
    
    except Exception:
        pass
    
    return None


def _force_close_excel_instances() -> None:
    """Force la fermeture de toutes les instances Excel ouvertes."""
    try:
        import psutil
        import os
        
        excel_processes = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'].lower() in ['excel.exe', 'xlwings32.exe']:
                    excel_processes.append(proc.info['pid'])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        if excel_processes:
            logger.debug(f"Fermeture forcée de {len(excel_processes)} processus Excel")
            for pid in excel_processes:
                try:
                    os.kill(pid, 9)
                except:
                    continue
    except ImportError:
        try:
            import os
            os.system("taskkill /f /im excel.exe 2>nul")
            os.system("taskkill /f /im xlwings32.exe 2>nul")
        except:
            pass