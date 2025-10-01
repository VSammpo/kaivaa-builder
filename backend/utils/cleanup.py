"""
Utilitaires de nettoyage des processus Office
Adapté de spirits_study pour KAIVAA Builder
"""

import os
import tempfile
import glob
from loguru import logger


def cleanup_before_run() -> None:
    """
    Nettoyage préventif avant exécution pour éviter les conflits COM Excel/PowerPoint.
    """
    logger.info("Nettoyage préventif...")
    
    _force_close_office_apps()
    _cleanup_temp_files()
    
    import time
    time.sleep(2.0)
    
    logger.info("Nettoyage terminé")


def _force_close_office_apps() -> None:
    """Force la fermeture d'Excel et PowerPoint"""
    try:
        os.system("taskkill /f /im excel.exe 2>nul 1>nul")
        os.system("taskkill /f /im powerpnt.exe 2>nul 1>nul") 
        os.system("taskkill /f /im xlwings32.exe 2>nul 1>nul")
        logger.debug("Applications Office fermées")
    except:
        pass


def _cleanup_temp_files() -> None:
    """Nettoie les fichiers temporaires Office"""
    try:
        temp_dir = tempfile.gettempdir()
        
        patterns = [
            "ppt_temp_*.pptx",
            "temp_output_*.pptx", 
            "TEMP_*.xlsx",
            "chart_*.png"
        ]
        
        cleaned = 0
        for pattern in patterns:
            files = glob.glob(os.path.join(temp_dir, pattern))
            for file_path in files:
                try:
                    os.remove(file_path)
                    cleaned += 1
                except:
                    continue
        
        if cleaned > 0:
            logger.debug(f"{cleaned} fichiers temporaires supprimés")
    except:
        pass