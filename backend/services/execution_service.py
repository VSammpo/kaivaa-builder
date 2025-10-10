# backend/services/generation_service.py
from typing import Dict, Any, Optional
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

from loguru import logger

from backend.services.project_service import ProjectService
from backend.services.template_service import TemplateService
from backend.services.execution_service import ExecutionService
from backend.database.models import ExecutionJob


class GenerationService:
    """
    Service d'orchestration de la génération des livrables.
    """
    
    def __init__(self, db_session):
        self.db = db_session
        self.ps = ProjectService(db_session)
        self.ts = TemplateService(db_session)
        self.es = ExecutionService(db_session)
    
    def generate_deliverable(
        self,
        project_id: str,
        template_id: int,
        parameters: Dict[str, Any],
        generate_excel: bool = True,
        generate_ppt: bool = True
    ) -> Dict[str, Any]:
        """
        Génère un livrable pour un projet avec les paramètres donnés.
        
        Returns:
            Dict avec 'excel_path', 'ppt_path', 'job_id'
        """
        # 1. Charger le projet et le livrable
        proj = self.ps.load_project(project_id)
        deliverable = next((d for d in proj.get("deliverables", []) if d["template_id"] == template_id), None)
        
        if not deliverable:
            raise ValueError(f"Livrable {template_id} introuvable dans le projet {project_id}")
        
        # 2. Vérifier que le livrable est fonctionnel
        if not deliverable.get("is_functional"):
            raise ValueError("Le livrable n'est pas prêt (données manquantes)")
        
        # 3. Récupérer les masters personnalisés du projet
        masters = deliverable.get("custom_masters", {})
        excel_master = masters.get("excel_path")
        ppt_master = masters.get("ppt_path")
        
        # 4. Créer un job d'exécution
        job = ExecutionJob(
            project_id=project_id,
            template_id=template_id,
            status="running",
            parameters=parameters
        )
        self.db.add(job)
        self.db.commit()
        
        try:
            # 5. Préparer les données (sources projet > défaut)
            data_sources = self._prepare_data_sources(project_id, template_id)
            
            # 6. Lancer l'exécution
            result = self.es.execute_template(
                template_id=template_id,
                parameters=parameters,
                data_sources=data_sources,
                excel_template_path=excel_master,
                ppt_template_path=ppt_master,
                generate_excel=generate_excel,
                generate_ppt=generate_ppt,
                output_dir=Path("output") / project_id
            )
            
            # 7. Mettre à jour le job
            job.status = "completed"
            job.output_excel_path = result.get("excel_path")
            job.output_ppt_path = result.get("ppt_path")
            job.completed_at = datetime.now(ZoneInfo("Europe/Paris"))
            self.db.commit()
            
            # 8. Mettre à jour le livrable dans le projet
            deliverable["last_generated_at"] = datetime.now(ZoneInfo("Europe/Paris")).isoformat()
            self.ps.save_project(proj)
            
            logger.success(f"Génération réussie : job #{job.id}")
            
            return {
                "job_id": job.id,
                "excel_path": result.get("excel_path"),
                "ppt_path": result.get("ppt_path")
            }
        
        except Exception as e:
            # Échec
            job.status = "failed"
            job.error_message = str(e)
            job.completed_at = datetime.now(ZoneInfo("Europe/Paris"))
            self.db.commit()
            
            logger.error(f"Échec génération job #{job.id} : {e}")
            raise
    
    def _prepare_data_sources(self, project_id: str, template_id: int) -> Dict[str, Any]:
        """
        Prépare les sources de données pour la génération.
        Priorité : sources client > sources par défaut du gabarit
        """
        # Récupérer les gabarits requis par le template
        usages = self.ts.list_gabarit_usages(template_id)
        
        data_sources = {}
        
        for usage in usages:
            gname = usage.get("gabarit_name")
            gver = usage.get("gabarit_version", "v1")
            
            if not gname:
                continue
            
            # Récupérer la source configurée dans le projet
            source = self.ps.get_data_source(project_id, gname, gver)
            
            if source:
                # Source configurée dans le projet
                data_sources[f"{gname}_{gver}"] = source.get("source_config")
            else:
                # Fallback : source par défaut du gabarit
                from backend.services.gabarit_registry import get_default_source
                default_src = get_default_source(gname, gver)
                
                if default_src:
                    data_sources[f"{gname}_{gver}"] = default_src
                else:
                    logger.warning(f"Aucune source disponible pour {gname} v{gver}")
        
        return data_sources