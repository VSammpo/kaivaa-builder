"""
Service de gestion des templates
"""

from pathlib import Path
from typing import List, Dict, Optional, Any
from datetime import datetime
from sqlalchemy.orm import Session
from loguru import logger

from backend.config import DatabaseConfig, PathConfig
from backend.database.models import Template, User, TemplateVersion
from backend.models.template_config import TemplateConfig
from backend.generator.template_generator import TemplateGenerator
import re


class TemplateService:
    """Service CRUD pour les templates"""
    
    def __init__(self, db_session: Session):
        """
        Initialise le service.
        
        Args:
            db_session: Session SQLAlchemy
        """
        self.db = db_session
    
    def create_template(
        self,
        config: TemplateConfig,
        user_id: int,
        ppt_source: Optional[Path] = None,
        excel_source: Optional[Path] = None
    ) -> Template:
        """
        Crée un nouveau template.
        
        Args:
            config: Configuration du template
            user_id: ID de l'utilisateur créateur
            ppt_source: Fichier PowerPoint source
            excel_source: Fichier Excel source
            
        Returns:
            Template créé
        """
        logger.info(f"Création du template '{config.name}'")
        
        # Vérifier si le nom existe déjà
        existing = self.db.query(Template).filter_by(name=config.name).first()
        if existing:
            raise ValueError(f"Un template nommé '{config.name}' existe déjà")
        
        # Générer les fichiers du template
        generator = TemplateGenerator(config)
        created_files = generator.generate(
            ppt_source=ppt_source,
            excel_source=excel_source,
            create_new=(ppt_source is None and excel_source is None)
        )
        
        # Créer l'entrée en base
        template = Template(
            name=config.name,
            description=config.description,
            version=config.version,
            created_by=user_id,
            config=config.model_dump(mode='json'), 
            config_file_path=str(created_files['config']),
            ppt_template_path=str(created_files['ppt']),
            excel_template_path=str(created_files['excel']),
            is_active=True
        )
        
        self.db.add(template)
        self.db.commit()
        self.db.refresh(template)
        
        # Créer la première version
        self._create_version(template, user_id, "Création initiale")
        
        # Faire un dernier refresh pour avoir l'objet complet
        self.db.refresh(template)
        
        logger.success(f"Template '{config.name}' créé (ID: {template.id})")
        return template
    
    def get_template(self, template_id: int) -> Optional[Template]:
        """
        Récupère un template par son ID.
        
        Args:
            template_id: ID du template
            
        Returns:
            Template ou None
        """
        return self.db.query(Template).filter_by(id=template_id).first()
    
    def get_template_by_name(self, name: str) -> Optional[Template]:
        """
        Récupère un template par son nom.
        
        Args:
            name: Nom du template
            
        Returns:
            Template ou None
        """
        return self.db.query(Template).filter_by(name=name).first()
    
    def list_templates(
        self,
        active_only: bool = True,
        user_id: Optional[int] = None
    ) -> List[Template]:
        """
        Liste tous les templates.
        
        Args:
            active_only: Si True, retourne uniquement les templates actifs
            user_id: Si fourni, filtre par créateur
            
        Returns:
            Liste des templates
        """
        query = self.db.query(Template)
        
        if active_only:
            query = query.filter_by(is_active=True)
        
        if user_id:
            query = query.filter_by(created_by=user_id)
        
        return query.order_by(Template.created_at.desc()).all()
    
    def update_template(
        self,
        template_id: int,
        updates: Dict[str, Any],
        user_id: int
    ) -> Template:
        """
        Met à jour un template.
        
        Args:
            template_id: ID du template
            updates: Dictionnaire des champs à mettre à jour
            user_id: ID de l'utilisateur faisant la modification
            
        Returns:
            Template mis à jour
        """
        template = self.get_template(template_id)
        if not template:
            raise ValueError(f"Template {template_id} non trouvé")
        
        logger.info(f"Mise à jour du template '{template.name}'")
        
        # Champs autorisés à la mise à jour
        allowed_fields = ['description', 'version', 'config', 'is_public', 'card_image_path']
        
        change_description = []
        for field, value in updates.items():
            if field in allowed_fields:
                old_value = getattr(template, field)
                setattr(template, field, value)
                change_description.append(f"{field}: {old_value} → {value}")
        
        template.updated_at = datetime.utcnow()
        
        self.db.commit()
        self.db.refresh(template)
        
        # Créer une nouvelle version
        self._create_version(
            template,
            user_id,
            f"Mise à jour: {', '.join(change_description)}"
        )
        
        logger.success(f"Template '{template.name}' mis à jour")
        return template
    
    def delete_template(self, template_id: int, hard_delete: bool = False) -> bool:
        """
        Supprime un template.
        
        Args:
            template_id: ID du template
            hard_delete: Si True, suppression définitive, sinon désactivation
            
        Returns:
            True si succès
        """
        template = self.get_template(template_id)
        if not template:
            raise ValueError(f"Template {template_id} non trouvé")
        
        if hard_delete:
            logger.warning(f"Suppression DÉFINITIVE du template '{template.name}'")
            
            # Supprimer les fichiers physiques
            template_dir = PathConfig.TEMPLATES / template.name
            if template_dir.exists():
                import shutil
                shutil.rmtree(template_dir)
                logger.info(f"Dossier supprimé : {template_dir}")
            
            # Supprimer de la base
            self.db.delete(template)
            self.db.commit()
            
            logger.success(f"Template '{template.name}' supprimé définitivement")
        else:
            logger.info(f"Désactivation du template '{template.name}'")
            template.is_active = False
            self.db.commit()
            logger.success(f"Template '{template.name}' désactivé")
        
        return True
    
    def get_template_stats(self, template_id: int) -> Dict[str, Any]:
        """
        Récupère les statistiques d'un template.
        
        Args:
            template_id: ID du template
            
        Returns:
            Dict avec statistiques
        """
        template = self.get_template(template_id)
        if not template:
            raise ValueError(f"Template {template_id} non trouvé")
        
        from backend.database.models import ExecutionJob
        
        total_executions = self.db.query(ExecutionJob).filter_by(
            template_id=template_id
        ).count()
        
        successful_executions = self.db.query(ExecutionJob).filter_by(
            template_id=template_id,
            status='completed'
        ).count()
        
        failed_executions = self.db.query(ExecutionJob).filter_by(
            template_id=template_id,
            status='failed'
        ).count()
        
        avg_execution_time = self.db.query(ExecutionJob).filter_by(
            template_id=template_id,
            status='completed'
        ).with_entities(
            ExecutionJob.execution_time_seconds
        ).all()
        
        avg_time = sum([t[0] for t in avg_execution_time if t[0]]) / len(avg_execution_time) if avg_execution_time else 0
        
        return {
            "template_id": template_id,
            "name": template.name,
            "total_executions": total_executions,
            "successful_executions": successful_executions,
            "failed_executions": failed_executions,
            "success_rate": round(successful_executions / total_executions * 100, 1) if total_executions > 0 else 0,
            "avg_execution_time_seconds": round(avg_time, 2),
            "last_execution": template.last_executed.isoformat() if template.last_executed else None
        }
    
    def _create_version(
        self,
        template: Template,
        user_id: int,
        description: str
    ) -> TemplateVersion:
        """Crée une nouvelle version d'un template"""
        version = TemplateVersion(
            template_id=template.id,
            version=template.version,
            config_snapshot=template.config,
            created_by=user_id,
            change_description=description
        )
        
        self.db.add(version)
        self.db.commit()
        
        logger.debug(f"Version {template.version} créée pour template {template.name}")
        return version
    
    def load_template_config(self, template_id: int) -> TemplateConfig:
        """
        Charge la configuration d'un template.
        
        Args:
            template_id: ID du template
            
        Returns:
            TemplateConfig
        """
        template = self.get_template(template_id)
        if not template:
            raise ValueError(f"Template {template_id} non trouvé")
        
        return TemplateConfig(**template.config)
    
    def _slugify(self, text: str) -> str:
        """Transforme un nom en slug 'propre' pour le nom de fichier."""
        import re as _re
        text = (text or "").lower()
        text = _re.sub(r'[^a-z0-9]+', '-', text).strip('-')
        return text or 'image'

    def save_card_image(self, template_id: int, file_bytes: bytes, original_filename: str) -> str:
        """
        Enregistre physiquement l'image de carte dans assets/background/card/
        et met à jour template.card_image_path en base.
        Retourne le chemin ABSOLU enregistré.
        """
        from pathlib import Path as _Path
        from datetime import datetime as _dt

        # 1) Récupérer le template
        template = self.get_template(template_id)
        if not template:
            raise ValueError(f"Template {template_id} non trouvé")

        # 2) Trouver la racine du projet
        #    a) si PathConfig.ROOT existe, on l'utilise
        #    b) sinon, on remonte depuis ce fichier: .../backend/services/template_service.py -> racine = parents[2]
        try:
            from backend.config import PathConfig  # optionnel
            project_root = _Path(PathConfig.ROOT)
        except Exception:
            project_root = _Path(__file__).resolve().parents[2]

        # 3) Dossier de sortie
        assets_dir = project_root / "assets" / "background" / "card"
        assets_dir.mkdir(parents=True, exist_ok=True)

        # 4) Nom de fichier propre et unique
        stem = self._slugify(template.name)
        ext = _Path(original_filename).suffix.lower() or ".png"
        fname = f"{stem}_{_dt.utcnow().strftime('%Y%m%d%H%M%S')}{ext}"
        out_path = assets_dir / fname

        # 5) Écriture du fichier
        with open(out_path, "wb") as f:
            f.write(file_bytes)

        # 6) Mise à jour du chemin en base (chemin ABSOLU)
        abs_path_str = str(out_path.resolve())
        template.card_image_path = abs_path_str
        template.updated_at = _dt.utcnow()
        self.db.commit()
        self.db.refresh(template)

        return abs_path_str
