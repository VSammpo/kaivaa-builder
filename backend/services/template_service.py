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
from backend.services.database_service import DatabaseService
# === Helpers de chemin/version ===
from pathlib import Path

def _templates_root_dir() -> Path:
    # Utilise configuration/templates comme racine file-based
    return Path(__file__).resolve().parents[2] / "configuration" / "templates"

def _ensure_version_layout(name: str, version: str) -> Path:
    """
    Garantit l’arborescence versionnée et migre l'ancien fichier <version>.json
    qui aurait pu être écrit à la racine du template vers <Version>/config.json.
    """
    root = _templates_root_dir() / name
    version_dir = root / str(version)
    version_dir.mkdir(parents=True, exist_ok=True)

    legacy_json = root / f"{version}.json"
    cfg_path = version_dir / "config.json"

    # Migration silencieuse si un ancien fichier existe
    if legacy_json.exists() and not cfg_path.exists():
        try:
            cfg_path.write_bytes(legacy_json.read_bytes())
            legacy_json.unlink()
        except Exception:
            pass  # on n'arrête pas l'app si la migration échoue

    # Si par erreur un <version>.json existe encore, on l'efface
    if legacy_json.exists():
        try:
            legacy_json.unlink()
        except Exception:
            pass

    return version_dir

# --- NOUVEL EMPLACEMENT DES FICHIERS DE TEMPLATES ---
def _templates_root_dir() -> Path:
    # si PathConfig.TEMPLATES existe, on l’ignore ici au profit de configuration/templates
    return Path(__file__).resolve().parents[2] / "configuration" / "templates"


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
        
        generator = TemplateGenerator(template_config=config)



        # Vérifier unicité par (nom, version)
        existing = (
            self.db.query(Template)
            .filter(Template.name == config.name, Template.version == config.version)
            .first()
        )
        if existing:
            raise ValueError(f"Un template '{config.name}' en version '{config.version}' existe déjà")

        
        # Générer les fichiers du template
        created_files = generator.generate(
            ppt_source=ppt_source,
            excel_source=excel_source,
            create_new=(ppt_source is None and excel_source is None)
        )

        # === CORRECTION : Gérer tous les dossiers legacy (templates/ ET template/) ===
        try:
            from backend.config import PathConfig
            project_root = Path(PathConfig.ROOT)
        except Exception:
            project_root = Path(__file__).resolve().parents[2]

        # ✅ 1. Déplacer depuis templates/ (pluriel)
        legacy_dir_plural = project_root / "templates" / config.name
        if legacy_dir_plural.exists():
            import shutil
            target_dir = PathConfig.TEMPLATES / config.name
            target_dir.parent.mkdir(parents=True, exist_ok=True)
            if target_dir.exists():
                shutil.rmtree(target_dir)
            shutil.move(str(legacy_dir_plural), str(target_dir))
            try:
                legacy_dir_plural.parent.rmdir()  # Supprime templates/ si vide
            except Exception:
                pass

        # ✅ 2. Déplacer depuis template/ (singulier)
        legacy_dir_singular = project_root / "template" / config.name
        if legacy_dir_singular.exists():
            import shutil
            target_dir = PathConfig.TEMPLATES / config.name
            target_dir.parent.mkdir(parents=True, exist_ok=True)
            # Fusionner avec l'existant si nécessaire
            if target_dir.exists():
                for item in legacy_dir_singular.iterdir():
                    shutil.move(str(item), str(target_dir / item.name))
            else:
                shutil.move(str(legacy_dir_singular), str(target_dir))
            try:
                legacy_dir_singular.rmdir()  # Supprime template/ si vide
            except Exception:
                pass

        # ✅ 3. Supprimer les dossiers racine s'ils sont vides
        for legacy_root in [project_root / "templates", project_root / "template"]:
            try:
                if legacy_root.exists() and not any(legacy_root.iterdir()):
                    legacy_root.rmdir()
            except Exception:
                pass

        # === Sortie sous configuration/templates/<Nom>/<Version>/... ===
        version_dir = _ensure_version_layout(config.name, config.version)

        conf_dest  = version_dir / "config.json"
        ppt_dest   = version_dir / "master.pptx"
        excel_dest = version_dir / "master.xlsx"

        from shutil import move
        if created_files.get('config'):
            move(str(created_files['config']), str(conf_dest))
        if created_files.get('ppt'):
            move(str(created_files['ppt']), str(ppt_dest))
        if created_files.get('excel'):
            move(str(created_files['excel']), str(excel_dest))

        created_files = {
            'config': conf_dest if conf_dest.exists() else None,
            'ppt': ppt_dest if ppt_dest.exists() else None,
            'excel': excel_dest if excel_dest.exists() else None,
        }




        # === Déplacer aussi les artefacts restants générés dans l'ancien dossier racine ===
        from shutil import move
        project_root = Path(__file__).resolve().parents[2]

        # Traiter templates/ ET template/ (singulier)
        for legacy_base in ("templates", "template"):
            legacy_root = project_root / legacy_base / config.name
            if legacy_root.exists() and legacy_root.is_dir():
                for item in legacy_root.iterdir():
                    try:
                        if item.is_dir():
                            # ex: queries/
                            dest = version_dir / item.name
                            dest.mkdir(parents=True, exist_ok=True)
                            for sub in item.iterdir():
                                move(str(sub), str(dest / sub.name))
                            try:
                                item.rmdir()
                            except Exception:
                                pass
                        else:
                            # ex: README, autres fichiers
                            move(str(item), str(version_dir / item.name))
                    except Exception:
                        # on ne bloque pas la création si un move échoue
                        pass
                # supprimer le dossier '<legacy_base>/<Nom>' s'il est vide
                try:
                    legacy_root.rmdir()
                except Exception:
                    pass

        # tenter aussi de supprimer les dossiers racine 'templates' et 'template' s'ils sont vides
        for maybe_empty in (project_root / "templates", project_root / "template"):
            try:
                next(maybe_empty.iterdir())
            except StopIteration:
                try:
                    maybe_empty.rmdir()
                except Exception:
                    pass
            except Exception:
                pass


        
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
    
    def delete_template(self, template_id: int, *, hard_delete: bool = False) -> bool:
        template = self.get_template(template_id)
        if not template:
            raise ValueError(f"Template {template_id} non trouvé")

        if hard_delete:
            logger.warning(f"Suppression DÉFINITIVE du template '{template.name}'")
            template_dir = _templates_root_dir() / template.name
            if template_dir.exists():
                import shutil
                shutil.rmtree(template_dir)
                logger.info(f"Dossier supprimé : {template_dir}")
            self.db.delete(template)
            self.db.commit()
            logger.success(f"Template '{template.name}' supprimé définitivement")
        else:
            logger.info(f"Soft delete du template '{template.name}'")
            info = self.soft_delete_template(template_id)
            logger.success(f"Template archivé sous _trash ({info['new_name']})")

        return True

    
    def soft_delete_template(self, template_id: int) -> dict:
        """
        Archive le template :
        - déplace configuration/templates/<Nom>/<Version> vers configuration/templates/_trash/<Nom_Supr_nXXXX>/<Version>
        - renomme le Template en base pour <Nom_Supr_nXXXX> + is_active=False
        Retour: { "old_name": ..., "new_name": ..., "version": ... }
        """
        import shutil
        with DatabaseService.get_session() as db:
            tpl = db.query(Template).filter(Template.id == template_id).first()
            if not tpl:
                raise ValueError("Template introuvable")

            old_name = tpl.name
            version = tpl.version or "v1"

            root = _templates_root_dir()
            src_dir = root / old_name / str(version)

            trash_root = root / "_trash"
            trash_root.mkdir(parents=True, exist_ok=True)

            prefix = f"{old_name}_Supr_n"
            existing = [p.name for p in trash_root.iterdir() if p.is_dir() and p.name.startswith(prefix)]
            if existing:
                try:
                    k = max(int(x.split(prefix, 1)[-1]) for x in existing) + 1
                except Exception:
                    k = len(existing) + 1
            else:
                k = 1
            new_name = f"{old_name}_Supr_n{str(k).zfill(4)}"

            dest_dir = trash_root / new_name / str(version)
            dest_dir.parent.mkdir(parents=True, exist_ok=True)

            if src_dir.exists():
                shutil.move(str(src_dir), str(dest_dir))
                try:
                    (root / old_name).rmdir()
                except Exception:
                    pass

            tpl.name = new_name
            tpl.is_active = False
            db.add(tpl)
            db.commit()

        return {"old_name": old_name, "new_name": new_name, "version": version}

    
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
        Lit la config file-based : configuration/templates/<Nom>/<Version>/config.json
        (migre/élimine un éventuel <version>.json à la racine du template).
        """
        with DatabaseService.get_session() as db:
            tpl = db.query(Template).filter(Template.id == template_id).first()
            if not tpl:
                return {}
            name = tpl.name
            version = tpl.version or "1.0"

        version_dir = _ensure_version_layout(name, version)
        cfg_path = version_dir / "config.json"
        if not cfg_path.exists():
            return {"name": name, "version": version, "parameters": [], "gabarit_usages": []}

        import json
        try:
            return json.loads(cfg_path.read_text(encoding="utf-8"))
        except Exception:
            return {}



    def update_config(self, template_id: int, new_config: dict) -> None:
        """
        Écrit configuration/templates/<Nom>/<Version>/config.json
        (fusion simple avec l'existant) et supprime tout <version>.json résiduel.
        """
        with DatabaseService.get_session() as db:
            tpl = db.query(Template).filter(Template.id == template_id).first()
            if not tpl:
                return
            name = tpl.name
            version = tpl.version or "1.0"

        import json, logging
        logger = logging.getLogger("kaivaa.template")

        version_dir = _ensure_version_layout(name, version)
        cfg_path = version_dir / "config.json"

        current = {}
        if cfg_path.exists():
            try:
                current = json.loads(cfg_path.read_text(encoding="utf-8"))
            except Exception:
                current = {}

        merged = dict(current)
        for k, v in (new_config or {}).items():
            merged[k] = v

        merged.setdefault("name", name)
        merged.setdefault("version", version)
        merged.setdefault("parameters", current.get("parameters", []))
        merged.setdefault("gabarit_usages", current.get("gabarit_usages", []))

        tmp = cfg_path.with_suffix(".tmp")
        tmp.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(cfg_path)

        # Nettoyage si un ancien <version>.json subsiste
        legacy_json = (_templates_root_dir() / name / f"{version}.json")
        if legacy_json.exists():
            try:
                legacy_json.unlink()
            except Exception:
                pass

        logger.info(f"[template_service] update_config -> {cfg_path}")


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
        Liste des *tables demandées* (usages de gabarit) rattachées au livrable.
        Tous les champs utiles sont normalisés et renvoyés pour ne rien perdre.
        """
        cfg = self.get_config(template_id)
        usages = cfg.get("gabarit_usages", [])
        usages = usages if isinstance(usages, list) else []

        norm: list[dict] = []
        for u in usages:
            if not isinstance(u, dict):
                continue
            excel = u.get("excel_target") or {}
            norm.append({
                "gabarit_name": (u.get("gabarit_name") or "").strip(),
                "gabarit_version": (u.get("gabarit_version") or "v1").strip(),
                "columns_enabled": [str(c).strip() for c in (u.get("columns_enabled") or []) if str(c).strip()],
                "excel_target": {
                    "sheet": (excel.get("sheet") or "").strip(),
                    "table": (excel.get("table") or "").strip(),
                },
                "methods": [str(m).strip() for m in (u.get("methods") or []) if str(m).strip()],
                "enrichments": u.get("enrichments") or [],          # ← garder tel quel
                "overlay_python": (u.get("overlay_python") or "").strip(),
                "final_order": list(u.get("final_order") or []),     # ← garder ordre final
                "final_excludes": [str(c).strip() for c in (u.get("final_excludes") or []) if str(c).strip()],
            })
        return norm


    def upsert_gabarit_usage(
        self,
        *,
        template_id: int,
        gabarit_name: str,
        gabarit_version: str = "v1",
        excel_sheet: str,
        excel_table: str,
        columns_enabled: list[str] | None = None,
        methods: list[str] | None = None,
        enrichments: list[dict] | None = None,
        final_order: list[str] | None = None,
        final_excludes: list[str] | None = None,
    ) -> dict:
        """
        Upsert d'un 'usage' (gabarit + feuille + table) dans le JSON de config du template.
        On préserve les champs existants non fournis (ex: overlay_python).
        """
        config = self.get_config(template_id) or {}
        allu = list(config.get("gabarit_usages") or [])

        key = (
            gabarit_name,
            (gabarit_version or "v1"),
            excel_sheet.strip(),
            excel_table.strip(),
        )

        # construire la charge utile (normalisée)
        payload = {
            "gabarit_name": gabarit_name,
            "gabarit_version": (gabarit_version or "v1"),
            "excel_target": {"sheet": key[2], "table": key[3]},
            "columns_enabled": list(columns_enabled or []),
            "methods": list(methods or []),
            "enrichments": list(enrichments or []),
        }
        if final_order is not None:
            payload["final_order"] = list(final_order)
        if final_excludes is not None:
            payload["final_excludes"] = list(final_excludes)

        # rechercher si un usage existe déjà pour cette clé
        idx = None
        for i, u in enumerate(allu):
            tgt = (u.get("excel_target") or {})
            k = (
                u.get("gabarit_name"),
                (u.get("gabarit_version") or "v1"),
                (tgt.get("sheet") or "").strip(),
                (tgt.get("table") or "").strip(),
            )
            if k == key:
                idx = i
                break

        if idx is None:
            # création
            allu.append(payload)
        else:
            # mise à jour en préservant les champs non gérés ici (ex: overlay_python)
            current = dict(allu[idx])
            current.update(payload)
            allu[idx] = current

        config["gabarit_usages"] = allu
        self.update_config(template_id, config)
        return payload


    def update_usage_final_view(
        self,
        template_id: int,
        gabarit_name: str,
        gabarit_version: str = "v1",
        final_order: list[str] | None = None,
        final_excludes: list[str] | None = None,
        final_renames: dict[str, str] | None = None,
    ) -> None:
        cfg = self.get_config(template_id)
        usages = cfg.get("gabarit_usages", [])
        if not isinstance(usages, list):
            usages = []

        gname = (gabarit_name or "").strip()
        gver = (gabarit_version or "v1").strip()

        new_usages = []
        updated = False
        for u in usages:
            if u.get("gabarit_name") == gname and (u.get("gabarit_version") or "v1") == gver:
                u = dict(u)
                if final_order is not None:
                    u["final_order"] = list(final_order or [])
                if final_excludes is not None:
                    u["final_excludes"] = [c for c in (final_excludes or []) if str(c).strip()]
                if final_renames is not None:
                    clean = {str(k): str(v) for k, v in (final_renames or {}).items() if str(k).strip() and str(v).strip()}
                    u["final_renames"] = clean
                updated = True
            new_usages.append(u)

        if updated:
            cfg["gabarit_usages"] = new_usages
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

    def get_table_role(self, template_id: int, gabarit_name: str, gabarit_version: str="v1") -> str | None:
        cfg = self.get_config(template_id)
        for r in (cfg.get("gabarit_roles") or []):
            if r.get("gabarit_name")==gabarit_name and (r.get("gabarit_version") or "v1")==gabarit_version:
                return r.get("table_role")
        return None

    # ---------- RELATIONS (catalogue) -------------------------------------

    def list_relations(self, template_id: int, from_gabarit: str | None=None, from_version: str | None=None) -> list[dict]:
        """
        Retourne la liste des relations autorisées (sans type de jointure).
        Si from_gabarit est fourni, filtre sur les relations sortantes depuis ce gabarit.
        """
        cfg = self.get_config(template_id)
        rels = cfg.get("gabarit_relations", [])
        rels = rels if isinstance(rels, list) else []
        if from_gabarit:
            fv = (from_version or "v1").strip()
            return [r for r in rels if r.get("from_gabarit")==from_gabarit and (r.get("from_version") or "v1")==fv]
        return rels

    def get_relation_by_id(self, template_id: int, relation_id: str) -> dict | None:
        cfg = self.get_config(template_id)
        for r in (cfg.get("gabarit_relations") or []):
            if r.get("relation_id") == relation_id:
                return r
        return None

    def add_relation(self, template_id: int,
                    from_gabarit: str, from_version: str,
                    to_gabarit: str, to_version: str,
                    left_key: str, right_key: str,
                    cardinality: str | None=None) -> None:
        """
        Déclare une relation autorisée entre 2 gabarits (clé ↔ clé).
        """
        cfg = self.get_config(template_id)
        rels = cfg.get("gabarit_relations", [])
        if not isinstance(rels, list):
            rels = []

        item = {
            "from_gabarit": (from_gabarit or "").strip(),
            "from_version": (from_version or "v1").strip(),
            "to_gabarit": (to_gabarit or "").strip(),
            "to_version": (to_version or "v1").strip(),
            "left_key": (left_key or "").strip(),
            "right_key": (right_key or "").strip(),
        }
        rid = (
            f"{item['from_gabarit']}|{item['from_version']}->"
            f"{item['to_gabarit']}|{item['to_version']}::"
            f"{item['left_key']}={item['right_key']}"
        )
        item["relation_id"] = rid
        if cardinality:
            item["cardinality"] = cardinality

        # anti-duplication exacte
        if item not in rels:
            rels.append(item)

        cfg["gabarit_relations"] = rels
        self.update_config(template_id, cfg)

    def delete_relation(self, template_id: int,
                        from_gabarit: str, from_version: str,
                        to_gabarit: str, to_version: str,
                        left_key: str, right_key: str) -> bool:
        cfg = self.get_config(template_id)
        rels = cfg.get("gabarit_relations", [])
        if not isinstance(rels, list):
            rels = []
        new_rels = [
            r for r in rels
            if not (
                r.get("from_gabarit")==from_gabarit and (r.get("from_version") or "v1")==from_version
                and r.get("to_gabarit")==to_gabarit and (r.get("to_version") or "v1")==to_version
                and r.get("left_key")==left_key and r.get("right_key")==right_key
            )
        ]
        if len(new_rels)==len(rels):
            return False
        cfg["gabarit_relations"] = new_rels
        self.update_config(template_id, cfg)
        return True

    def list_enrichment_paths(self, start_name: str, start_version: str="v1", max_depth: int=2) -> list[list[tuple]]:
        """
        Retourne des chemins possibles sous forme de séquences (g_from, col_from, g_to, col_to).
        max_depth=2 autorise 2 sauts (A->B->C).
        """
        from backend.services.gabarit_registry import get_relations

        paths = []
        frontier = [[(start_name, start_version)]]
        depth = 0
        while frontier and depth < max_depth:
            new_frontier = []
            for p in frontier:
                cur_name, cur_ver = p[-1]
                rels = get_relations(cur_name, cur_ver) or []
                for r in rels:
                    step = (r["from_gabarit"], r.get("from_version","v1"),
                            r["left_key"], r["to_gabarit"], r.get("to_version","v1"),
                            r["right_key"])
                    path_for_ui = [(step[0], step[2], step[3], step[5])]
                    paths.append(path_for_ui)
                    new_frontier.append(p + [(r["to_gabarit"], r.get("to_version","v1"))])
            frontier = new_frontier
            depth += 1

        return paths


    def get_tables_demandees(self, template_id: int) -> list[dict]:
        """
        Vue synthétique pour l'UI : tables demandées (usages normalisés).
        """
        return self.list_gabarit_usages(template_id)

    def get_required_input_columns_for_usage(self, template_id: int, gname: str, gver: str = "v1") -> list[str]:
        """
        Colonnes MINIMALES à fournir dans la table d'entrée (gabarit de base) pour cet usage :
        - columns_enabled (ou toutes les colonnes du gabarit si vide)
        - + toutes les left_key du PREMIER saut des enrichissements (pour pouvoir démarrer les chaînes)
        - + colonnes d'entrée des méthodes sélectionnées (registry: "inputs": ["colA","colB",...])
        """
        from backend.services.gabarit_registry import get_gabarit, list_methods_for_gabarit

        u = self.get_gabarit_usage(template_id, gname, gver) or {}
        g = get_gabarit(gname, gver)
        base_cols = [c.name for c in (g.columns or [])]

        required = u.get("columns_enabled") or base_cols[:]

        # 1) toutes les left_key du premier saut des enrichissements
        for e in (u.get("enrichments") or []):
            p = e.get("path") or []
            if p:
                first_left_key = p[0][1]  # [from, left_key, to, right_key]
                if first_left_key and first_left_key not in required:
                    required.append(first_left_key)

        # 2) inputs des méthodes sélectionnées
        sel = set(u.get("methods") or [])
        if sel:
            allm = list_methods_for_gabarit(gname, gver) or []
            it = (allm.values() if isinstance(allm, dict) else allm)
            for m in it:
                if isinstance(m, dict) and m.get("name") in sel:
                    for cin in (m.get("inputs") or []):
                        if cin not in required:
                            required.append(cin)

        return list(dict.fromkeys(required))
    
    def compute_required_tables_for_usage(self, template_id: int, gname: str, gver: str = "v1") -> dict[tuple[str, str], list[str]]:
        """
        Retourne un dict { (gabarit_name, gabarit_version): [colonnes requises] }
        en tenant compte des enrichissements multi-sauts et des méthodes.
        """
        from backend.services.gabarit_registry import get_gabarit, list_methods_for_gabarit

        u = self.get_gabarit_usage(template_id, gname, gver) or {}
        result: dict[tuple[str, str], list[str]] = {}

        # Helpers
        def _ensure(d: dict, key: tuple[str, str]) -> list[str]:
            if key not in d:
                d[key] = []
            return d[key]

        # 0) Base
        g = get_gabarit(gname, gver)
        base_cols = [c.name for c in (g.columns or [])]
        base_req = u.get("columns_enabled") or base_cols[:]
        # + inputs de méthodes
        sel = set(u.get("methods") or [])
        if sel:
            allm = list_methods_for_gabarit(gname, gver) or []
            it = (allm.values() if isinstance(allm, dict) else allm)
            for m in it:
                if isinstance(m, dict) and m.get("name") in sel:
                    for cin in (m.get("inputs") or []):
                        if cin not in base_req:
                            base_req.append(cin)

        # + toutes les left_key des 1ers sauts
        for e in (u.get("enrichments") or []):
            p = e.get("path") or []
            if p:
                lk = p[0][1]
                if lk and lk not in base_req:
                    base_req.append(lk)

        result[(gname, gver)] = list(dict.fromkeys(base_req))

        # 1) Parcourir les chemins pour poser les clés sur chaque table
        for e in (u.get("enrichments") or []):
            path = e.get("path") or []
            # poser les clés de pas en pas
            for step in path:
                frm, lkey, to, rkey = step  # [from, left_key, to, right_key]
                frm_key = (frm, "v1")
                to_key = (to, "v1")
                lst_from = _ensure(result, frm_key)
                lst_to = _ensure(result, to_key)
                if lkey and lkey not in lst_from:
                    lst_from.append(lkey)
                if rkey and rkey not in lst_to:
                    lst_to.append(rkey)

            # ajouter les colonnes sélectionnées sur la table cible (dernier 'to')
            if path:
                last_to = (path[-1][2], "v1")
                lst_last = _ensure(result, last_to)
                for c in (e.get("columns") or []):
                    if c not in lst_last:
                        lst_last.append(c)

        # dédup ordonnée
        for k, cols in result.items():
            result[k] = list(dict.fromkeys(cols))
        return result


    
    def get_gabarit_usage_by_target(
        self,
        template_id: int,
        gabarit_name: str,
        gabarit_version: str,
        excel_sheet: str,
        excel_table: str,
    ) -> dict | None:
        """Retourne l'usage correspondant à (gabarit, version, sheet, table)."""
        cfg = self.get_config(template_id)
        usages = cfg.get("gabarit_usages", [])
        if not isinstance(usages, list):
            return None
        gname = (gabarit_name or "").strip()
        gver  = (gabarit_version or "v1").strip()
        sheet = (excel_sheet or "").strip()
        table = (excel_table or "").strip()
        for u in usages:
            if (
                u.get("gabarit_name") == gname
                and (u.get("gabarit_version") or "v1") == gver
                and ((u.get("excel_target") or {}).get("sheet", "") or "") == sheet
                and ((u.get("excel_target") or {}).get("table", "") or "") == table
            ):
                return u
        return None
    
    def resolve_usage_expected_columns(self, template_id: int, gabarit_name: str, gabarit_version: str) -> list[str]:
        """
        Retourne la liste finale des colonnes attendues pour un usage,
        en tenant compte de l'ordre final, des exclusions et des renommages.
        """
        usage = self.get_gabarit_usage(template_id, gabarit_name, gabarit_version)
        if not usage:
            return []
        
        from backend.services.gabarit_registry import get_gabarit, list_methods_for_gabarit
        
        # Ordre final (ou colonnes par défaut)
        final_order = usage.get("final_order") or []
        if not final_order:
            g = get_gabarit(gabarit_name, gabarit_version)
            base_cols = [c.name for c in (g.columns or [])]
            final_order = usage.get("columns_enabled") or base_cols[:]
            
            # Ajouter colonnes enrichies
            for e in (usage.get("enrichments") or []):
                for c in (e.get("columns") or []):
                    if c not in final_order:
                        final_order.append(c)
            
            # Ajouter sorties de méthodes
            methods_selected = set(usage.get("methods") or [])
            if methods_selected:
                all_methods = list_methods_for_gabarit(gabarit_name, gabarit_version) or []
                it = (all_methods.values() if isinstance(all_methods, dict) else all_methods)
                for m in it:
                    if isinstance(m, dict) and m.get("name") in methods_selected:
                        out_col = (m.get("output_column") or "").strip()
                        if out_col and out_col not in final_order:
                            final_order.append(out_col)
        
        # Exclusions
        final_excludes = set(usage.get("final_excludes") or [])
        
        # Renommages (appliquer les nouveaux noms)
        final_renames = usage.get("final_renames") or {}
        
        result = []
        for col in final_order:
            if col in final_excludes:
                continue
            # Utiliser le nom renommé si présent
            new_name = final_renames.get(col, col)
            if new_name and new_name not in result:
                result.append(new_name)
        
        return result
    
    def cache_parameter_options(self, template_id: int, param_name: str, values: list[str], source: dict) -> None:
        """
        Persiste dans configuration/templates/<Nom>/<Version>/config.json
        la liste 'values' pour le paramètre 'param_name' sous la clé parameters[].options_cache.
        """
        cfg = self.get_config(template_id) or {}
        plist = list(cfg.get("parameters") or [])

        updated = False
        for p in plist:
            if p.get("name") == param_name:
                p["options_cache"] = {
                    "values": list(values or [])[:500],  # borne raisonnable
                    "source": {
                        "gabarit": (source or {}).get("gabarit"),
                        "version": (source or {}).get("version"),
                        "column":  (source or {}).get("column"),
                    },
                }
                updated = True
                break

        if not updated:
            plist.append({
                "name": param_name,
                "options_mode": "from_column",
                "options_source": {
                    "gabarit": (source or {}).get("gabarit"),
                    "version": (source or {}).get("version"),
                    "column":  (source or {}).get("column"),
                },
                "options_cache": {
                    "values": list(values or [])[:500],
                    "source": {
                        "gabarit": (source or {}).get("gabarit"),
                        "version": (source or {}).get("version"),
                        "column":  (source or {}).get("column"),
                    },
                },
            })

        cfg["parameters"] = plist
        self.update_config(template_id, cfg)

