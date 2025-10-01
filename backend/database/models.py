"""
Modèles de base de données pour KAIVAA Builder.
Architecture prête pour multi-utilisateurs et gestion des droits.
"""

from datetime import datetime
from typing import Optional
from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, Boolean, JSON
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship

Base = declarative_base()


class User(Base):
    """Utilisateurs du système (préparation future)"""
    __tablename__ = "users"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(50), unique=True, nullable=False)
    email = Column(String(100), unique=True, nullable=False)
    role = Column(String(20), default="user")  # admin, user, viewer
    created_at = Column(DateTime, default=datetime.utcnow)
    is_active = Column(Boolean, default=True)
    
    # Relations
    templates = relationship("Template", back_populates="creator")
    jobs = relationship("ExecutionJob", back_populates="user")


class Template(Base):
    """Templates créés par les utilisateurs"""
    __tablename__ = "templates"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(100), unique=True, nullable=False)
    description = Column(Text)
    version = Column(String(20), default="1.0")
    
    # Métadonnées
    created_by = Column(Integer, ForeignKey("users.id"))
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Configuration (stockée en JSON)
    config = Column(JSON, nullable=False)
    
    # Chemins des fichiers
    config_file_path = Column(String(500))
    ppt_template_path = Column(String(500))
    excel_template_path = Column(String(500))
    
    # Statistiques
    execution_count = Column(Integer, default=0)
    last_executed = Column(DateTime, nullable=True)
    
    # État
    is_active = Column(Boolean, default=True)
    is_public = Column(Boolean, default=False)
    
    # Relations
    creator = relationship("User", back_populates="templates")
    versions = relationship("TemplateVersion", back_populates="template", cascade="all, delete-orphan")
    jobs = relationship("ExecutionJob", back_populates="template")
    permissions = relationship("TemplatePermission", back_populates="template", cascade="all, delete-orphan")


class TemplateVersion(Base):
    """Versioning des templates"""
    __tablename__ = "template_versions"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    template_id = Column(Integer, ForeignKey("templates.id", ondelete="CASCADE"))
    version = Column(String(20), nullable=False)
    
    # Configuration de cette version
    config_snapshot = Column(JSON, nullable=False)
    
    # Métadonnées
    created_at = Column(DateTime, default=datetime.utcnow)
    created_by = Column(Integer, ForeignKey("users.id"))
    change_description = Column(Text)
    
    # Relations
    template = relationship("Template", back_populates="versions")


class TemplatePermission(Base):
    """Gestion des droits d'accès aux templates (préparation future)"""
    __tablename__ = "template_permissions"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    template_id = Column(Integer, ForeignKey("templates.id", ondelete="CASCADE"))
    user_id = Column(Integer, ForeignKey("users.id"))
    
    # Niveaux : view, edit, execute, admin
    permission_level = Column(String(20), default="view")
    
    granted_at = Column(DateTime, default=datetime.utcnow)
    granted_by = Column(Integer, ForeignKey("users.id"))
    
    # Relations
    template = relationship("Template", back_populates="permissions")


class ExecutionJob(Base):
    """Historique des générations de rapports"""
    __tablename__ = "execution_jobs"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    template_id = Column(Integer, ForeignKey("templates.id"))
    user_id = Column(Integer, ForeignKey("users.id"))
    
    # Paramètres d'exécution
    parameters = Column(JSON, nullable=False)
    
    # État
    status = Column(String(20), default="pending")  # pending, running, completed, failed
    
    # Chemins de sortie
    output_excel_path = Column(String(500))
    output_ppt_path = Column(String(500))
    
    # Timing
    created_at = Column(DateTime, default=datetime.utcnow)
    started_at = Column(DateTime, nullable=True)
    completed_at = Column(DateTime, nullable=True)
    
    # Résultats
    error_message = Column(Text, nullable=True)
    execution_time_seconds = Column(Integer, nullable=True)
    
    # Relations
    template = relationship("Template", back_populates="jobs")
    user = relationship("User", back_populates="jobs")


class CustomTable(Base):
    """Tables personnalisées avec SQL + Python"""
    __tablename__ = "custom_tables"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    template_id = Column(Integer, ForeignKey("templates.id"))
    
    # Identification
    table_name = Column(String(100), nullable=False)
    description = Column(Text)
    
    # Code
    sql_query = Column(Text, nullable=False)
    python_code = Column(Text, nullable=True)  # Optionnel
    
    # Métadonnées
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Validation
    is_validated = Column(Boolean, default=False)
    last_validation_result = Column(Text)