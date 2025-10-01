"""
Utilitaires de gestion de fichiers et chemins
Adapté de spirits_study pour KAIVAA Builder
"""

import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional
from loguru import logger


def get_output_paths(
    study_name: str,
    category: str,
    brand: str,
    batch: str,
    distributor: str,
    template_name: str
) -> Dict[str, str]:
    """
    Génère une structure de chemins intelligente pour les outputs.
    
    Args:
        study_name: Nom de l'étude
        category: Catégorie (ex: "Gin", "Whisky")
        brand: Marque
        batch: ID du batch
        distributor: Distributeur
        template_name: Nom du template utilisé
        
    Returns:
        Dict avec les chemins excel, pptx, temp, etc.
        
    Structure générée:
        output/{study_name}/{category}/{brand}/Excel/{batch}_{brand}_{category}_{distributor}_{template}.xlsx
        output/{study_name}/{category}/{brand}/PowerPoint/{batch}_{brand}_{category}_{distributor}_{template}.pptx
    """
    from backend.config import PathConfig
    
    # Nettoyage des noms
    study_clean = clean_filename(study_name)
    category_clean = clean_filename(category)
    brand_clean = clean_filename(brand)
    batch_clean = clean_filename(batch)
    distributor_clean = clean_filename(distributor)
    template_clean = clean_filename(template_name)
    
    # Structure de base
    base_dir = PathConfig.OUTPUT / study_clean / category_clean / brand_clean
    excel_dir = base_dir / "Excel"
    pptx_dir = base_dir / "PowerPoint"
    
    # Noms des fichiers
    excel_filename = f"{batch_clean}_{brand_clean}_{category_clean}_{distributor_clean}_{template_clean}.xlsx"
    pptx_filename = f"{batch_clean}_{brand_clean}_{category_clean}_{distributor_clean}_{template_clean}.pptx"
    temp_excel_filename = f"TEMP_{excel_filename}"
    
    # Chemins complets
    return {
        "excel_dir": str(excel_dir),
        "pptx_dir": str(pptx_dir),
        "excel_path": str(excel_dir / excel_filename),
        "pptx_path": str(pptx_dir / pptx_filename),
        "temp_excel": str(excel_dir / temp_excel_filename),
        "base_dir": str(base_dir)
    }


def generate_batch_id(prefix: str = "") -> str:
    """
    Génère un identifiant de batch unique.
    
    Args:
        prefix: Préfixe optionnel
        
    Returns:
        ID au format [prefix_]YYYYMMDD_HHmm
        
    Example:
        generate_batch_id("test")  # "test_20251001_1430"
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    
    if prefix:
        prefix_clean = clean_filename(prefix)
        return f"{prefix_clean}_{timestamp}"
    
    return timestamp


def ensure_directories(*paths: str) -> None:
    """
    S'assure que les répertoires existent.
    
    Args:
        *paths: Chemins de fichiers ou dossiers
    """
    for path in paths:
        if not path:
            continue
            
        directory = os.path.dirname(path) if os.path.isfile(path) or '.' in os.path.basename(path) else path
        
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
            logger.debug(f"Dossier créé : {directory}")


def clean_filename(name: str) -> str:
    """
    Nettoie un nom pour l'utiliser dans un nom de fichier.
    
    Args:
        name: Nom à nettoyer
        
    Returns:
        Nom nettoyé (caractères spéciaux remplacés par _)
        
    Example:
        clean_filename("Mon Fichier/Test")  # "Mon_Fichier_Test"
    """
    if not name:
        return "unknown"
    
    replacements = {
        " ": "_", "/": "_", "\\": "_", ":": "_", "*": "_", 
        "?": "_", '"': "_", "<": "_", ">": "_", "|": "_"
    }
    
    result = name
    for old, new in replacements.items():
        result = result.replace(old, new)
    
    # Suppression des underscores multiples
    while "__" in result:
        result = result.replace("__", "_")
    
    return result.strip("_")


def get_file_size_mb(file_path: str) -> float:
    """
    Retourne la taille d'un fichier en MB.
    
    Args:
        file_path: Chemin du fichier
        
    Returns:
        Taille en MB
    """
    if not os.path.exists(file_path):
        return 0.0
    
    size_bytes = os.path.getsize(file_path)
    return round(size_bytes / (1024 * 1024), 2)