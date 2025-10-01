"""
Configuration centralisée de l'application KAIVAA Builder.
Gère les variables d'environnement et les chemins.
"""

import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

# Chargement des variables d'environnement
load_dotenv()

# Chemins de base
PROJECT_ROOT = Path(__file__).parent.parent
BACKEND_DIR = PROJECT_ROOT / "backend"
FRONTEND_DIR = PROJECT_ROOT / "frontend"
TEMPLATES_DIR = PROJECT_ROOT / "templates"
OUTPUT_DIR = PROJECT_ROOT / "output"
LOGS_DIR = PROJECT_ROOT / "logs"

# Création des dossiers s'ils n'existent pas
for directory in [TEMPLATES_DIR, OUTPUT_DIR, LOGS_DIR]:
    directory.mkdir(exist_ok=True)


class DatabaseConfig:
    """Configuration PostgreSQL"""
    
    HOST: str = os.getenv("DB_HOST", "localhost")
    PORT: int = int(os.getenv("DB_PORT", "5432"))
    NAME: str = os.getenv("DB_NAME", "kaivaa_builder")
    USER: str = os.getenv("DB_USER", "postgres")
    PASSWORD: str = os.getenv("DB_PASSWORD", "")
    
    @classmethod
    def get_connection_string(cls) -> str:
        """Retourne la chaîne de connexion PostgreSQL"""
        return f"postgresql://{cls.USER}:{cls.PASSWORD}@{cls.HOST}:{cls.PORT}/{cls.NAME}"


class AppConfig:
    """Configuration de l'application"""
    
    ENV: str = os.getenv("APP_ENV", "development")
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    SECRET_KEY: str = os.getenv("SECRET_KEY", "dev-secret-key-change-in-prod")
    
    # Versions
    VERSION: str = "0.1.0"
    APP_NAME: str = "KAIVAA Builder"
    
    @classmethod
    def is_production(cls) -> bool:
        return cls.ENV == "production"
    
    @classmethod
    def is_development(cls) -> bool:
        return cls.ENV == "development"


class PathConfig:
    """Configuration des chemins"""
    
    TEMPLATES = TEMPLATES_DIR
    OUTPUT = OUTPUT_DIR
    LOGS = LOGS_DIR
    
    @classmethod
    def get_template_path(cls, template_name: str) -> Path:
        """Retourne le chemin vers un template spécifique"""
        return cls.TEMPLATES / template_name
    
    @classmethod
    def get_output_path(cls, filename: str) -> Path:
        """Retourne le chemin de sortie pour un fichier"""
        return cls.OUTPUT / filename


# Exports principaux
__all__ = ["DatabaseConfig", "AppConfig", "PathConfig", "PROJECT_ROOT"]