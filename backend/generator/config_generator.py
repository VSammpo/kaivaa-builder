"""
Générateur de fichiers config.yaml
"""

from pathlib import Path
from loguru import logger

from backend.models.template_config import TemplateConfig


class ConfigGenerator:
    """Génère les fichiers config.yaml pour les templates"""
    
    def __init__(self, config: TemplateConfig):
        self.config = config
    
    def generate(self, output_path: Path) -> None:
        """
        Génère le fichier config.yaml.
        
        Args:
            output_path: Chemin de sortie
        """
        yaml_content = self.config.to_yaml()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(yaml_content)
        
        logger.info(f"Fichier config.yaml généré : {output_path}")