# backend/services/generation_service.py
"""
Service d'orchestration de la génération des livrables.
Fait le pont entre les projets et le moteur de génération (ReportService).
"""

from typing import Dict, Any
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

from loguru import logger

from backend.services.project_service import ProjectService
from backend.services.template_service import TemplateService
from backend.services.report_service import ReportService
from backend.database.models import ExecutionJob
from backend.services.iteration_service import IterationService
from pathlib import Path


class GenerationService:
    """
    Service d'orchestration de la génération des livrables.
    """
    
    def __init__(self, db_session):
        self.db = db_session
        self.ps = ProjectService(db_session)
        self.ts = TemplateService(db_session)
    
    def generate_deliverable(self, project_id: str, template_id: int, parameters: Dict[str, Any], 
                         generate_excel: bool = True, generate_ppt: bool = True) -> Dict[str, Any]:
        logger.info(f"[GenerationService] Génération livrable template={template_id} projet={project_id}")
        
        # 1. Charger le projet et le livrable
        proj = self.ps.load_project(project_id)
        deliverable = next((d for d in proj.get("deliverables", []) if d["template_id"] == template_id), None)
        
        if not deliverable:
            raise ValueError(f"Livrable {template_id} introuvable dans le projet {project_id}")
        
        # 2. Vérifier que le livrable est fonctionnel
        if not deliverable.get("is_functional"):
            raise ValueError("Le livrable n'est pas prêt (données manquantes)")
        
        # 3. Récupérer les masters personnalisés du projet ✅
        masters = deliverable.get("custom_masters", {})
        excel_master = masters.get("excel_path")
        ppt_master = masters.get("ppt_path")
        
        # 4. Récupérer la config du template
        template_config = self.ts.load_template_config(template_id)
        
        # 5. Créer un job d'exécution
        job = ExecutionJob(
            project_id=project_id,
            template_id=template_id,
            status="running",
            parameters=parameters,
            started_at=datetime.now(ZoneInfo("Europe/Paris"))
        )
        self.db.add(job)
        self.db.commit()
        self.db.refresh(job)
        
        try:
            # 6. Créer le ReportService avec la config du template
            report_service = ReportService(template_config)
            
            # 7. Préparer le nom de sortie
            timestamp = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y%m%d_%H%M%S")
            output_name = f"{proj.get('name', 'projet')}_{template_config.name}_{timestamp}"
            
            # 8. Générer via ReportService avec les masters du projet ✅
            result = report_service.generate_report(
                parameters=parameters,
                output_name=output_name,
                project_id=project_id,
                excel_master_path=excel_master,  # ✅ Maintenant défini
                ppt_master_path=ppt_master        # ✅ Maintenant défini
            )
            
            if not result.get("success"):
                raise RuntimeError(result.get("error", "Génération échouée"))
            
            # 8. Normaliser les clés de sortie
            excel_path = result.get("excel_path") or result.get("xlsx_path")
            ppt_path   = result.get("ppt_path")   or result.get("pptx_path")

            # 9. Mettre à jour le job
            job.status = "completed"
            job.output_excel_path = excel_path
            job.output_ppt_path   = ppt_path
            job.execution_time_seconds = int(result.get("execution_time_seconds", 0))
            job.completed_at = datetime.now(ZoneInfo("Europe/Paris"))
            self.db.commit()

            try:
                if project_id and ppt_path:
                    IterationService(self.db).register_generated_output(
                        project_id=project_id,
                        template_name=template_config.name,
                        ppt_path=Path(ppt_path) if ppt_path else None,
                        parameters=parameters,
                        excel_path=Path(excel_path) if excel_path else None,
                    )
            except Exception as meta_err:
                logger.warning(f"[GenerationService] Impossible d'enregistrer l'itération : {meta_err}")


            # 10. Mettre à jour le livrable dans le projet
            deliverable["last_generated_at"] = datetime.now(ZoneInfo("Europe/Paris")).isoformat()
            self.ps.save_project(proj)

            logger.success(f"[GenerationService] Génération réussie : job #{job.id}")

            return {
                "job_id": job.id,
                "excel_path": excel_path,
                "ppt_path": ppt_path,
                "success": True
            }

        
        except Exception as e:
            # Échec
            job.status = "failed"
            job.error_message = str(e)
            job.completed_at = datetime.now(ZoneInfo("Europe/Paris"))
            self.db.commit()
            
            logger.error(f"[GenerationService] Échec génération job #{job.id} : {e}")
            raise