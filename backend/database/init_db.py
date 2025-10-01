"""
Script d'initialisation de la base de données.
Crée toutes les tables et un utilisateur admin par défaut.
"""

import sys
from pathlib import Path

# Ajouter la racine du projet au PYTHONPATH
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from loguru import logger

from backend.config import DatabaseConfig
from backend.database.models import Base, User


def init_database():
    """Initialise la base de données avec toutes les tables"""
    
    connection_string = DatabaseConfig.get_connection_string()
    logger.info(f"Connexion à la base de données : {DatabaseConfig.HOST}:{DatabaseConfig.PORT}/{DatabaseConfig.NAME}")
    
    try:
        # Création du moteur SQLAlchemy
        engine = create_engine(connection_string, echo=False)
        
        # Création de toutes les tables
        logger.info("Création des tables...")
        Base.metadata.create_all(engine)
        logger.success("Tables créées avec succès")
        
        # Création d'un utilisateur admin par défaut
        Session = sessionmaker(bind=engine)
        session = Session()
        
        # Vérifier si un admin existe déjà
        existing_admin = session.query(User).filter_by(username="admin").first()
        
        if not existing_admin:
            admin_user = User(
                username="admin",
                email="admin@kaivaa.local",
                role="admin",
                is_active=True
            )
            session.add(admin_user)
            session.commit()
            logger.success("Utilisateur admin créé (username: admin)")
        else:
            logger.info("Utilisateur admin existe déjà")
        
        session.close()
        logger.success("Base de données initialisée avec succès !")
        
    except Exception as e:
        logger.error(f"Erreur lors de l'initialisation : {e}")
        raise


if __name__ == "__main__":
    init_database()