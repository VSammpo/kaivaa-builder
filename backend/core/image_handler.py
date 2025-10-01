"""
Image Handler - Gestion de l'injection d'images dans PowerPoint
Adapté de spirits_study pour KAIVAA Builder
"""

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from loguru import logger


def resolve_image_path(
    pattern: str, 
    replacements: Dict[str, str], 
    default_path: Optional[str] = None
) -> Optional[str]:
    """
    Résout le chemin d'une image en remplaçant les variables par les valeurs des balises.
    Teste automatiquement PNG puis JPG.
    
    Args:
        pattern: Pattern avec variables (ex: "assets/products/{Marque}/{Produit}.png")
        replacements: Dict des balises (ex: {"[Marque]": "BOMBAY"})
        default_path: Chemin par défaut si l'image principale n'existe pas
        
    Returns:
        Chemin de l'image à utiliser, ou None si aucune image trouvée
        
    Example:
        pattern = "assets/{Marque}/{Produit}.png"
        replacements = {"[Marque]": "BOMBAY", "[Produit]": "Gin"}
        path = resolve_image_path(pattern, replacements)
    """
    try:
        variables = re.findall(r'\{([^}]+)\}', pattern)
        resolved_path = pattern
        
        for var in variables:
            balise_key = f"[{var}]"
            
            if balise_key in replacements:
                value = replacements[balise_key] if var == "Produit" else clean_for_filename(replacements[balise_key])
                resolved_path = resolved_path.replace(f"{{{var}}}", value)
                logger.debug(f"Variable {var} résolue : '{value}'")
            else:
                found_value = _find_balise_value_flexible(var, replacements)
                if found_value:
                    resolved_path = resolved_path.replace(f"{{{var}}}", found_value)
                    logger.debug(f"Variable {var} trouvée via recherche flexible : '{found_value}'")
                else:
                    logger.warning(f"Variable {var} non trouvée dans les balises")
                    return _try_default_path(default_path)
        
        # Test PNG puis JPG
        if os.path.exists(resolved_path):
            absolute_path = os.path.abspath(resolved_path)
            logger.debug(f"Image PNG trouvée : {absolute_path}")
            return absolute_path
        else:
            logger.debug(f"Image PNG non trouvée : {resolved_path}")
            
            if resolved_path.lower().endswith('.png'):
                jpg_path = resolved_path[:-4] + '.jpg'
                if os.path.exists(jpg_path):
                    absolute_path = os.path.abspath(jpg_path)
                    logger.debug(f"Image JPG trouvée : {absolute_path}")
                    return absolute_path
            
            return _try_default_path(default_path)
    
    except Exception as e:
        logger.error(f"Erreur résolution image : {e}")
        return _try_default_path(default_path)


def inject_image_to_slide(slide, image_config: dict, replacements: dict) -> None:
    """
    Injecte une image selon image_config, en résolvant le chemin via les balises.
    
    Args:
        slide: Slide PowerPoint
        image_config: Configuration de l'image (pattern, position, size, background, etc.)
        replacements: Dict des balises pour résolution
        
    Example:
        image_config = {
            "pattern": "assets/{Marque}/{Produit}.png",
            "default_path": "assets/default.png",
            "position": {"left": 10, "top": 155},
            "size": {"max_width": 30, "max_height": 90},
            "background": False
        }
    """
    # Constantes MsoZOrderCmd
    MSO_BRINGTOFRONT = 0
    MSO_SENDBACKWARD = 1
    MSO_BRINGFORWARD = 2
    MSO_SENDTOBACK = 3

    def _set_image_z_order(shape, send_to_background: bool) -> None:
        try:
            if send_to_background:
                shape.ZOrder(MSO_SENDTOBACK)
                try:
                    while getattr(shape, "ZOrderPosition", 1) > 1:
                        shape.ZOrder(MSO_SENDBACKWARD)
                except Exception:
                    pass
            else:
                shape.ZOrder(MSO_BRINGTOFRONT)
        except Exception as e:
            logger.debug(f"Erreur positionnement Z-order : {e}")

    # Résolution du chemin
    pattern = image_config.get("pattern")
    default_path = image_config.get("default_path")
    
    if not isinstance(pattern, str) or not pattern:
        logger.warning("Aucun 'pattern' string fourni pour l'image")
        return

    image_path = resolve_image_path(pattern=pattern, replacements=replacements, default_path=default_path)
    if not image_path:
        logger.warning(f"Aucune image trouvée pour pattern='{pattern}'")
        return
    
    if not os.path.exists(image_path):
        logger.warning(f"Image introuvable (après résolution): {image_path}")
        return

    # Dimensions slide
    try:
        slide_width = slide.Parent.PageSetup.SlideWidth
        slide_height = slide.Parent.PageSetup.SlideHeight
    except Exception:
        slide_width, slide_height = 960, 540

    # Lecture options
    pos = image_config.get("position", {})
    size = image_config.get("size", {})
    left = pos.get("left", 0)
    top = pos.get("top", 0)
    width = size.get("width")
    height = size.get("height")

    fit_to_slide = bool(image_config.get("fit_to_slide", False))
    keep_aspect = image_config.get("keep_aspect", None)
    name_override = image_config.get("name")
    background = bool(image_config.get("background", False))

    if fit_to_slide and (width is None or height is None):
        left, top, width, height = 0, 0, slide_width, slide_height

    # Insertion
    try:
        if width is not None and height is not None:
            new_shape = slide.Shapes.AddPicture(
                FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                Left=left, Top=top, Width=width, Height=height
            )
        else:
            new_shape = slide.Shapes.AddPicture(
                FileName=image_path, LinkToFile=False, SaveWithDocument=True,
                Left=left, Top=top
            )
    except Exception as e:
        logger.error(f"Échec AddPicture pour {image_path} : {e}")
        return

    # Options shape
    if keep_aspect is not None:
        try:
            new_shape.LockAspectRatio = -1 if keep_aspect else 0
        except Exception:
            pass
    
    if name_override:
        try:
            new_shape.Name = str(name_override)
        except Exception:
            pass

    # Z-order
    _set_image_z_order(new_shape, send_to_background=background)
    
    action = "arrière-plan" if background else "avant-plan"
    logger.debug(f"Image '{os.path.basename(image_path)}' positionnée à l'{action}")


def find_slides_by_ids(presentation, target_slide_ids: List[str]) -> Dict[str, object]:
    """
    Trouve les slides PowerPoint contenant les IDs spécifiés.
    Gère les shapes groupées.
    
    Args:
        presentation: Présentation PowerPoint
        target_slide_ids: Liste des IDs à rechercher (ex: ["A001", "A002"])
        
    Returns:
        Dict {slide_id: slide_object}
    """
    def search_text_in_shape(shape, target_ids):
        """Recherche récursive dans les shapes"""
        found_ids = {}
        
        try:
            if shape.Type == 6:  # Groupe
                for i in range(1, shape.GroupItems.Count + 1):
                    sub_results = search_text_in_shape(shape.GroupItems.Item(i), target_ids)
                    found_ids.update(sub_results)
            elif hasattr(shape, 'HasTextFrame') and shape.HasTextFrame:
                text = shape.TextFrame2.TextRange.Text
                if text:
                    for slide_id in target_ids:
                        pattern = re.compile(r'\b' + re.escape(slide_id) + r'\b')
                        if pattern.search(text):
                            found_ids[slide_id] = True
        except:
            pass
        
        return found_ids
    
    slides_found = {}
    
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            shape_results = search_text_in_shape(shape, target_slide_ids)
            
            for slide_id in shape_results:
                if slide_id not in slides_found:
                    slides_found[slide_id] = slide
                    logger.debug(f"Slide {slide_id} trouvée (index {slide.SlideIndex})")
    
    missing_slides = [slide_id for slide_id in target_slide_ids if slide_id not in slides_found]
    if missing_slides:
        logger.warning(f"Slides non trouvées : {', '.join(missing_slides)}")
    
    return slides_found


def clean_for_filename(text: str) -> str:
    """
    Nettoie un texte pour l'utiliser dans un nom de fichier.
    
    Args:
        text: Texte à nettoyer
        
    Returns:
        Texte nettoyé
    """
    if not text:
        return "unknown"
    
    cleaned = text.strip()
    
    replacements = {
        " ": "_", "/": "_", "\\": "_", ":": "_", "*": "_", 
        "?": "_", '"': "_", "<": "_", ">": "_", "|": "_",
        "-": "_", ".": "_", "&": "and", "%": "pct"
    }
    
    for old, new in replacements.items():
        cleaned = cleaned.replace(old, new)
    
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    
    return cleaned.strip("_") if cleaned else "unknown"


def _find_balise_value_flexible(var_name: str, replacements: Dict[str, str]) -> Optional[str]:
    """Recherche flexible d'une valeur de balise avec différentes variantes."""
    search_patterns = [
        f"[{var_name}]",
        f"[{var_name.lower()}]",
        f"[{var_name.upper()}]",
        f"[{var_name.title()}]",
    ]
    
    name_variants = {
        "Catégorie": ["Category", "Segment", "Type"],
        "Marque": ["Brand", "Sous_marque", "SousMarque"],
        "Distributeur": ["Distributor", "Enseigne"],
        "Produit": ["Product"]
    }
    
    if var_name in name_variants:
        for variant in name_variants[var_name]:
            search_patterns.extend([
                f"[{variant}]",
                f"[{variant.lower()}]",
                f"[{variant.upper()}]"
            ])
    
    for pattern in search_patterns:
        if pattern in replacements:
            value = replacements[pattern]
            if value and value.strip():
                return value if var_name == "Produit" else clean_for_filename(value)
    
    return None


def _try_default_path(default_path: Optional[str]) -> Optional[str]:
    """Essaie d'utiliser l'image par défaut."""
    if default_path and os.path.exists(default_path):
        absolute_path = os.path.abspath(default_path)
        logger.debug(f"Utilisation image par défaut : {absolute_path}")
        return absolute_path
    elif default_path:
        logger.warning(f"Image par défaut non trouvée : {default_path}")
    return None


# Compat pour ancienne API
def inject_images_to_slide(slide, images_config: list, replacements: dict) -> int:
    """
    Ancienne signature pour compatibilité.
    Injecte plusieurs images sur une slide.
    
    Returns:
        Nombre d'images injectées
    """
    injected = 0
    if not images_config:
        return 0
    
    for img_cfg in images_config:
        try:
            inject_image_to_slide(slide, img_cfg, replacements)
            injected += 1
        except Exception as e:
            logger.warning(f"Erreur injection image {img_cfg.get('type', 'image')} : {e}")
    
    return injected