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


class GenerationService:
    """
    Service d'orchestration de la génération des livrables.
    """
    
    def __init__(self, db_session):
        self.db = db_session
        self.ps = ProjectService(db_session)
        self.ts = TemplateService(db_session)
    
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
        
        Args:
            project_id: ID du projet
            template_id: ID du template/livrable
            parameters: Paramètres de génération (dict nom -> valeur)
            generate_excel: Générer le fichier Excel
            generate_ppt: Générer le fichier PowerPoint
        
        Returns:
            Dict avec 'job_id', 'excel_path', 'ppt_path'
        """
        logger.info(f"[GenerationService] Génération livrable template={template_id} projet={project_id}")
        
        # 1. Charger le projet et le livrable
        proj = self.ps.load_project(project_id)
        deliverable = next((d for d in proj.get("deliverables", []) if d["template_id"] == template_id), None)
        
        if not deliverable:
            raise ValueError(f"Livrable {template_id} introuvable dans le projet {project_id}")
        
        # 2. Vérifier que le livrable est fonctionnel
        if not deliverable.get("is_functional"):
            raise ValueError("Le livrable n'est pas prêt (données manquantes)")
        
        # 3. Récupérer la config du template
        template_config = self.ts.load_template_config(template_id)
        
        # 4. Créer un job d'exécution
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
            # 5. Créer le ReportService avec la config du template
            report_service = ReportService(template_config)
            
            # 6. Préparer le nom de sortie
            timestamp = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y%m%d_%H%M%S")
            output_name = f"{proj.get('name', 'projet')}_{template_config.name}_{timestamp}"
            
            # 7. Générer via ReportService (qui gère l'injection depuis le projet)
            result = report_service.generate_report(
                parameters=parameters,
                output_name=output_name,
                project_id=project_id
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