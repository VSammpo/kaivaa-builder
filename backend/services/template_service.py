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
from zoneinfo import ZoneInfo
from datetime import datetime

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
        
        template.updated_at = datetime.now(ZoneInfo("Europe/Paris"))
        
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
        fname = f"{stem}_{_dt.now(ZoneInfo('Europe/Paris')).strftime('%Y%m%d%H%M%S')}{ext}"
        out_path = assets_dir / fname

        # 5) Écriture du fichier
        with open(out_path, "wb") as f:
            f.write(file_bytes)

        # 6) Mise à jour du chemin en base (chemin ABSOLU)
        abs_path_str = str(out_path.resolve())
        template.card_image_path = abs_path_str
        template.updated_at = _dt.now(ZoneInfo("Europe/Paris"))
        self.db.commit()
        self.db.refresh(template)

        return abs_path_str
    

    def get_config(self, template_id: int) -> dict:
        """
        Retourne le JSON config du template, toujours avec des clés par défaut.
        IMPORTANT: self.db est une Session SQLAlchemy, ne pas appeler get_session() ici.
        """
        tpl = self.db.query(Template).get(template_id)
        cfg = tpl.config or {}
        if not isinstance(cfg, dict):
            cfg = {}

        # Compat legacy: ancienne clé 'contracts' (on la garde vide, mais on n'en dépend plus)
        if "contracts" not in cfg or not isinstance(cfg["contracts"], dict):
            cfg["contracts"] = {}

        # Clés MVP: usages & sources de gabarits par livrable
        if "gabarit_usages" not in cfg or not isinstance(cfg["gabarit_usages"], list):
            cfg["gabarit_usages"] = []
        if "gabarit_sources" not in cfg or not isinstance(cfg["gabarit_sources"], list):
            cfg["gabarit_sources"] = []

        return cfg


    def update_config(self, template_id: int, new_config: dict) -> None:
        """
        Écrase la config du template par new_config (et garantit les clés par défaut).
        """
        cfg = new_config or {}
        if not isinstance(cfg, dict):
            cfg = {}

        # Compat legacy
        if "contracts" not in cfg or not isinstance(cfg["contracts"], dict):
            cfg["contracts"] = {}

        # Clés MVP
        if "gabarit_usages" not in cfg or not isinstance(cfg["gabarit_usages"], list):
            cfg["gabarit_usages"] = []
        if "gabarit_sources" not in cfg or not isinstance(cfg["gabarit_sources"], list):
            cfg["gabarit_sources"] = []

        tpl = self.db.query(Template).get(template_id)
        tpl.config = cfg
        self.db.add(tpl)
        self.db.commit()
        self.db.refresh(tpl)


    def list_gabarit_sources(self, template_id: int) -> list[dict]:
        cfg = self.get_config(template_id)
        sources = cfg.get("gabarit_sources", [])
        return sources if isinstance(sources, list) else []

    def upsert_gabarit_source(self, template_id: int, gabarit_name: str, gabarit_version: str, source: dict) -> None:
        """
        source (MVP CSV) :
        {
          "type": "csv",
          "path": "C:/.../file.csv",
          "sep": ";",
          "encoding": "utf-8-sig"
        }
        """
        cfg = self.get_config(template_id)
        sources = cfg.get("gabarit_sources", [])
        if not isinstance(sources, list):
            sources = []

        gabarit_name = (gabarit_name or "").strip()
        gabarit_version = (gabarit_version or "v1").strip()

        # Remplacer si déjà présent (name+version)
        sources = [
            s for s in sources
            if not (s.get("gabarit_name") == gabarit_name and s.get("gabarit_version") == gabarit_version)
        ]
        sources.append({
            "gabarit_name": gabarit_name,
            "gabarit_version": gabarit_version,
            "source": source
        })

        cfg["gabarit_sources"] = sources
        self.update_config(template_id, cfg)

    def get_gabarit_source(self, template_id: int, gabarit_name: str, gabarit_version: str):
        for s in self.list_gabarit_sources(template_id):
            if s.get("gabarit_name") == gabarit_name and s.get("gabarit_version") == gabarit_version:
                return s.get("source")
        return None

    def delete_gabarit_source(self, template_id: int, gabarit_name: str, gabarit_version: str) -> bool:
        cfg = self.get_config(template_id)
        sources = cfg.get("gabarit_sources", [])
        if not isinstance(sources, list):
            sources = []

        new_sources = [
            s for s in sources
            if not (s.get("gabarit_name") == gabarit_name and s.get("gabarit_version") == gabarit_version)
        ]
        if len(new_sources) == len(sources):
            return False

        cfg["gabarit_sources"] = new_sources
        self.update_config(template_id, cfg)
        return True


    def list_gabarit_usages(self, template_id: int) -> list[dict]:
        """
        Liste des *tables demandées* (usages de gabarit) rattachées au livrable
        (templates.config['gabarit_usages']).

        Chaque item est normalisé au format :
        {
        "gabarit_name": str,
        "gabarit_version": str,
        "columns_enabled": [str, ...],
        "excel_target": {"sheet": str, "table": str},
        "methods": [str, ...]                # NEW (MVP: noms symboliques)
        }
        """
        cfg = self.get_config(template_id)
        usages = cfg.get("gabarit_usages", [])
        usages = usages if isinstance(usages, list) else []

        # Normalisation rétrocompatible
        norm: list[dict] = []
        for u in usages:
            if not isinstance(u, dict):
                continue
            norm.append({
                "gabarit_name": (u.get("gabarit_name") or "").strip(),
                "gabarit_version": (u.get("gabarit_version") or "v1").strip(),
                "columns_enabled": [c.strip() for c in (u.get("columns_enabled") or []) if c and str(c).strip()],
                "excel_target": {
                    "sheet": (u.get("excel_target", {}).get("sheet") or "").strip(),
                    "table": (u.get("excel_target", {}).get("table") or "").strip(),
                },
                "methods": [m.strip() for m in (u.get("methods") or []) if m and str(m).strip()],  # NEW
            })
        return norm


    def upsert_gabarit_usage(
        self,
        template_id: int,
        gabarit_name: str,
        gabarit_version: str,
        columns_enabled: list[str],
        excel_sheet: str,
        excel_table: str,
        methods: Optional[list[str]] = None,   # NEW
    ) -> None:
        """
        Crée/MAJ une *table demandée* (usage de gabarit) pour ce livrable (clé: gabarit_name+version).

        Args:
            template_id: id du template livrable
            gabarit_name: nom du gabarit
            gabarit_version: version du gabarit (ex: 'v1')
            columns_enabled: colonnes cochées pour ce livrable (sous-ensemble du gabarit)
            excel_sheet: feuille Excel cible (ListObject)
            excel_table: nom du ListObject Excel cible
            methods: liste de méthodes à appliquer (noms symboliques)  # NEW
        """
        cfg = self.get_config(template_id)
        usages = cfg.get("gabarit_usages", [])
        if not isinstance(usages, list):
            usages = []

        gabarit_name = (gabarit_name or "").strip()
        gabarit_version = (gabarit_version or "v1").strip()
        excel_sheet = (excel_sheet or "").strip()
        excel_table = (excel_table or "").strip()
        columns_enabled = [c.strip() for c in (columns_enabled or []) if c and str(c).strip()]
        methods = [m.strip() for m in (methods or []) if m and str(m).strip()]  # NEW

        # Remplacer si déjà présent (name+version)
        usages = [
            u for u in usages
            if not (u.get("gabarit_name") == gabarit_name and u.get("gabarit_version") == gabarit_version)
        ]
        usages.append({
            "gabarit_name": gabarit_name,
            "gabarit_version": gabarit_version,
            "columns_enabled": columns_enabled,
            "excel_target": {"sheet": excel_sheet, "table": excel_table},
            "methods": methods,  # NEW
        })

        cfg["gabarit_usages"] = usages
        self.update_config(template_id, cfg)


    def delete_gabarit_usage(self, template_id: int, gabarit_name: str, gabarit_version: str) -> bool:
        """
        Supprime l'usage (name+version) pour ce livrable. Renvoie True si supprimé.
        """
        cfg = self.get_config(template_id)
        usages = cfg.get("gabarit_usages", [])
        if not isinstance(usages, list):
            usages = []
        new_usages = [
            u for u in usages
            if not (u.get("gabarit_name") == gabarit_name and u.get("gabarit_version") == gabarit_version)
        ]
        if len(new_usages) == len(usages):
            return False
        cfg["gabarit_usages"] = new_usages
        self.update_config(template_id, cfg)
        return True

    def get_gabarit_usage(self, template_id: int, gabarit_name: str, gabarit_version: str = "v1") -> dict | None:
        """
        Retourne l'usage (table demandée) pour (gabarit_name, gabarit_version) si présent,
        sinon None.
        """
        gabarit_name = (gabarit_name or "").strip()
        gabarit_version = (gabarit_version or "v1").strip()
        for u in self.list_gabarit_usages(template_id):
            if u.get("gabarit_name") == gabarit_name and u.get("gabarit_version") == gabarit_version:
                return u
        return None
    
    def set_usage_methods(
        self,
        template_id: int,
        gabarit_name: str,
        gabarit_version: str,
        methods: list[str],
    ) -> None:
        """
        Remplace la liste des 'methods' pour une table demandée (usage).
        """
        cfg = self.get_config(template_id)
        usages = cfg.get("gabarit_usages", [])
        if not isinstance(usages, list):
            usages = []

        gabarit_name = (gabarit_name or "").strip()
        gabarit_version = (gabarit_version or "v1").strip()
        methods = [m.strip() for m in (methods or []) if m and str(m).strip()]

        new_usages = []
        updated = False
        for u in usages:
            if u.get("gabarit_name") == gabarit_name and u.get("gabarit_version") == gabarit_version:
                u = dict(u)
                u["methods"] = methods
                updated = True
            new_usages.append(u)

        if not updated:
            # si l'usage n'existe pas, on le crée minimalement (colonnes et cible vides)
            new_usages.append({
                "gabarit_name": gabarit_name,
                "gabarit_version": gabarit_version,
                "columns_enabled": [],
                "excel_target": {"sheet": "", "table": ""},
                "methods": methods,
            })

        cfg["gabarit_usages"] = new_usages
        self.update_config(template_id, cfg)


    def add_method_to_usage(
        self,
        template_id: int,
        gabarit_name: str,
        gabarit_version: str,
        method_name: str,
    ) -> None:
        """
        Ajoute une méthode (si absente) sur une table demandée.
        """
        method = (method_name or "").strip()
        if not method:
            return
        usage = self.get_gabarit_usage(template_id, gabarit_name, gabarit_version)
        methods = (usage or {}).get("methods", [])
        if method not in methods:
            methods.append(method)
        self.set_usage_methods(template_id, gabarit_name, gabarit_version, methods)


    def remove_method_from_usage(
        self,
        template_id: int,
        gabarit_name: str,
        gabarit_version: str,
        method_name: str,
    ) -> None:
        """
        Retire une méthode (si présente) d'une table demandée.
        """
        method = (method_name or "").strip()
        usage = self.get_gabarit_usage(template_id, gabarit_name, gabarit_version)
        methods = [m for m in (usage or {}).get("methods", []) if m != method]
        self.set_usage_methods(template_id, gabarit_name, gabarit_version, methods)

    def get_tables_demandees(self, template_id: int) -> list[dict]:
        """
        Vue synthétique pour l'UI : tables demandées (usages normalisés).
        """
        return self.list_gabarit_usages(template_id)
    
    def resolve_usage_expected_columns(self, template_id: int, gabarit_name: str, gabarit_version: str = "v1") -> list[str]:
        """
        Colonnes 'attendues' pour une table demandée = colonnes cochées (columns_enabled)
        ∪ colonnes exigées par les méthodes sélectionnées (get_method_requirements).
        L'ordre conserve d'abord 'columns_enabled', puis on ajoute les manquantes.
        """
        usage = self.get_gabarit_usage(template_id, gabarit_name, gabarit_version) or {}
        enabled = [c for c in (usage.get("columns_enabled") or []) if isinstance(c, str) and c.strip()]
        methods = [m for m in (usage.get("methods") or []) if isinstance(m, str) and m.strip()]

        # dépendances de colonnes par méthodes
        try:
            from backend.services.gabarit_registry import get_method_requirements
            required = get_method_requirements(gabarit_name, gabarit_version, methods) or []
        except Exception:
            required = []

        # union ordonnée: cochées d'abord, puis exigences méthodes
        out: list[str] = []
        for c in enabled:
            if c not in out:
                out.append(c)
        for c in required:
            if c not in out:
                out.append(c)
        return out
