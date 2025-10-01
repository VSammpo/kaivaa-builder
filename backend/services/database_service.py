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