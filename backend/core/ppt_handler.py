"""
PowerPoint Handler - Gestion des opérations PowerPoint via win32com
Adapté de spirits_study pour KAIVAA Builder
"""

import os
import uuid
import re
from contextlib import contextmanager
from typing import Dict, List, Optional
import win32com.client as win32
from loguru import logger
import pythoncom
import pywintypes
from win32com.client import constants

@contextmanager
def powerpoint_app_context(ppt_path: str, visible: bool = True, read_only: bool = False):
    """
    Contexte PowerPoint avec initialisation COM correcte sur le thread courant.
    - Initialise COM (STA) avant Dispatch
    - Ouvre éventuellement la présentation
    - Ferme proprement et désinitialise COM
    """
    ppt_app = None
    presentation = None
    initialized_here = False
    try:
        # --- Initialiser COM en STA sur ce thread ---
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            initialized_here = True
        except pywintypes.com_error as e:
            # RPC_E_CHANGED_MODE (-2147417850) = déjà initialisé en MTA : on continue.
            if e.hresult != -2147417850:
                raise

        # --- Lancer PowerPoint ---
        logger.debug(f"Ouverture PowerPoint: {ppt_path}")
        try:
            ppt_app = win32.Dispatch("PowerPoint.Application")
        except pywintypes.com_error as e:
            raise RuntimeError(f"Echec Dispatch PowerPoint.Application: {e}") from e

        ppt_app.Visible = True if visible else False


        # --- Ouvrir la présentation ---
        try:
            presentation = ppt_app.Presentations.Open(
                os.path.abspath(ppt_path),
                WithWindow=visible,
                ReadOnly=read_only,
                Untitled=False
            )
        except pywintypes.com_error as e:
            # Nettoyage si l'ouverture échoue
            try:
                ppt_app.Quit()
            except Exception:
                pass
            raise RuntimeError(f"Erreur lors de l'ouverture de PowerPoint ({ppt_path}): {e}") from e

        yield ppt_app, presentation

    finally:
        # Fermer la présentation
        try:
            if presentation is not None:
                presentation.Close()
                logger.debug("Présentation fermée")
        except Exception:
            pass

        # Quitter PowerPoint
        try:
            if ppt_app is not None:
                ppt_app.Quit()
                logger.debug("Application PowerPoint fermée")
        except Exception:
            pass

        # Désinitialiser COM si on l'a initialisé ici
        if initialized_here:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def replace_tags_in_text_range(text_range, replacements: Dict[str, str]) -> None:
    """
    Remplace les balises dans un TextRange PowerPoint.
    
    Args:
        text_range: Objet TextRange PowerPoint
        replacements: Dictionnaire {balise: valeur}
    """
    full_text = text_range.Text
    for tag, value in replacements.items():
        pos = full_text.find(tag)
        while pos != -1:
            try:
                text_range.Characters(pos + 1, len(tag)).Text = str(value)
                full_text = text_range.Text
                pos = full_text.find(tag)
            except:
                full_text = full_text.replace(tag, str(value))
                text_range.Text = full_text
                break


def replace_tags_in_shape(shape, replacements: Dict[str, str]) -> None:
    """
    Remplace les balises dans une shape PowerPoint (texte, tableau, groupe).
    
    Args:
        shape: Objet Shape PowerPoint
        replacements: Dictionnaire {balise: valeur}
    """
    try:
        if shape.Type == 6:  # Groupe
            for i in range(1, shape.GroupItems.Count + 1):
                replace_tags_in_shape(shape.GroupItems.Item(i), replacements)
        elif shape.HasTable:
            table = shape.Table
            for row in range(1, table.Rows.Count + 1):
                for col in range(1, table.Columns.Count + 1):
                    try:
                        text_range = table.Cell(row, col).Shape.TextFrame2.TextRange
                        replace_tags_in_text_range(text_range, replacements)
                    except:
                        continue
        elif shape.HasTextFrame:
            replace_tags_in_text_range(shape.TextFrame2.TextRange, replacements)
    except Exception as e:
        logger.debug(f"Erreur remplacement balises shape: {e}")


def find_slide_by_id(presentation, slide_id: str) -> Optional[object]:
    """
    Trouve une slide contenant un ID spécifique.
    
    Args:
        presentation: Présentation PowerPoint
        slide_id: ID à rechercher (ex: "A003")
        
    Returns:
        Slide PowerPoint ou None si non trouvée
    """
    slide_id_pattern = re.compile(r"\b" + re.escape(slide_id) + r"\b")
    
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if hasattr(shape, 'HasTextFrame') and shape.HasTextFrame:
                try:
                    text = shape.TextFrame2.TextRange.Text
                    if slide_id_pattern.search(text):
                        logger.debug(f"Slide {slide_id} trouvée à l'index {slide.SlideIndex}")
                        return slide
                except:
                    continue
    
    logger.warning(f"Slide {slide_id} non trouvée")
    return None


def check_and_remove_suppressed_slides(presentation) -> List[str]:
    """
    Vérifie et supprime les slides contenant la balise [@SUPR@].
    
    Args:
        presentation: Présentation PowerPoint
        
    Returns:
        Liste des slides supprimées (pour logging)
    """
    def has_suppression_tag(slide) -> bool:
        """Vérifie si une slide contient [@SUPR@]"""
        def check_shape_for_tag(shape) -> bool:
            try:
                if shape.Type == 6:  # Groupe
                    for i in range(1, shape.GroupItems.Count + 1):
                        if check_shape_for_tag(shape.GroupItems.Item(i)):
                            return True
                elif shape.HasTable:
                    table = shape.Table
                    for row in range(1, table.Rows.Count + 1):
                        for col in range(1, table.Columns.Count + 1):
                            try:
                                text = table.Cell(row, col).Shape.TextFrame2.TextRange.Text
                                if "[@SUPR@]" in text:
                                    return True
                            except:
                                continue
                elif shape.HasTextFrame:
                    text = shape.TextFrame2.TextRange.Text
                    if "[@SUPR@]" in text:
                        return True
            except:
                pass
            return False
        
        try:
            for shape in slide.Shapes:
                if check_shape_for_tag(shape):
                    return True
        except:
            pass
        return False
    
    slides_to_remove = []
    
    try:
        slides_count = presentation.Slides.Count
        logger.debug(f"Vérification de {slides_count} slides pour [@SUPR@]")
        
        for i in range(slides_count, 0, -1):
            try:
                slide = presentation.Slides(i)
                if has_suppression_tag(slide):
                    slides_to_remove.append(f"Slide {i}")
                    try:
                        slide.Delete()
                        logger.info(f"Slide {i} supprimée (balise [@SUPR@])")
                    except Exception as delete_error:
                        logger.warning(f"Erreur suppression slide {i}: {delete_error}")
                        continue
            except Exception as slide_error:
                logger.warning(f"Erreur accès slide {i}: {slide_error}")
                continue
    
    except Exception as count_error:
        logger.warning(f"Impossible d'accéder aux slides: {count_error}")
        return []
    
    return slides_to_remove