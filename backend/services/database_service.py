"""
Service de gestion de la connexion à la base de données
"""

from contextlib import contextmanager
from typing import Generator
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, Session as SessionType
from loguru import logger

from backend.config import DatabaseConfig
from backend.database.models import Base
import os
import shutil
from datetime import datetime
from typing import Optional
from backend.database.models import ExecutionJob
from sqlalchemy.orm import Session
from zoneinfo import ZoneInfo
from datetime import datetime


class DatabaseService:
    """Service de gestion de la base de données"""
    
    _engine = None
    _session_factory = None
    
    @classmethod
    def initialize(cls):
        """Initialise la connexion à la base de données"""
        if cls._engine is None:
            connection_string = DatabaseConfig.get_connection_string()
            logger.info(f"Initialisation de la base de données")
            
            cls._engine = create_engine(
                connection_string,
                pool_pre_ping=True,
                pool_size=5,
                max_overflow=10
            )
            
            cls._session_factory = sessionmaker(bind=cls._engine)
            logger.success("Base de données initialisée")
    
    @classmethod
    def create_tables(cls):
        """Crée toutes les tables si elles n'existent pas"""
        cls.initialize()
        Base.metadata.create_all(cls._engine)
        logger.info("Tables créées/vérifiées")
    
    @classmethod
    @contextmanager
    def get_session(cls) -> Generator[SessionType, None, None]:
        """
        Context manager pour obtenir une session de base de données.
        
        Yields:
            Session SQLAlchemy
            
        Example:
            with DatabaseService.get_session() as db:
                templates = db.query(Template).all()
        """
        cls.initialize()
        session = cls._session_factory()
        try:
            yield session
            session.commit()
        except Exception as e:
            session.rollback()
            logger.error(f"Erreur base de données : {e}")
            raise
        finally:
            session.close()

    @staticmethod
    def _move_to_trash(file_path: Optional[str]) -> Optional[str]:
        """
        Déplace un fichier vers le dossier '00 - Supr' (même dossier que le fichier source).
        Retourne le chemin de destination si déplacement OK, sinon None.
        """
        if not file_path or not os.path.isfile(file_path):
            return None
        try:
            template_dir = os.path.dirname(file_path)
            trash_dir = os.path.join(template_dir, "00 - Supr")
            os.makedirs(trash_dir, exist_ok=True)

            base = os.path.basename(file_path)
            stem, ext = os.path.splitext(base)
            ts = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y%m%d_%H%M%S")
            dest = os.path.join(trash_dir, f"{stem}__DELETED_{ts}{ext}")

            shutil.move(file_path, dest)
            logger.info(f"Fichier déplacé vers la corbeille: {dest}")
            return dest
        except Exception as e:
            logger.warning(f"Impossible de déplacer '{file_path}' vers la corbeille: {e}")
            return None

    @staticmethod
    def delete_job_and_files(db: Session, job_id: int) -> bool:
        """
        Supprime une exécution :
        - Déplace les fichiers Excel/PPT associés vers '00 - Supr' (soft delete)
        - Supprime la ligne d'historique en base
        Retourne True si une ligne a été supprimée, False sinon.
        """
        job = db.query(ExecutionJob).filter(ExecutionJob.id == job_id).first()
        if not job:
            return False

        DatabaseService._move_to_trash(job.output_excel_path)
        DatabaseService._move_to_trash(job.output_ppt_path)

        db.delete(job)
        db.commit()
        logger.info(f"Exécution {job_id} supprimée (fichiers envoyés en corbeille).")
        return True