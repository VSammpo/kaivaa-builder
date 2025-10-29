from pathlib import Path
import json
import uuid
from datetime import datetime
from backend.services.project_service import ProjectService
from backend.core.ppt_handler import merge_presentations


class IterationService:
    """
    Gère :
    - stockage des livrables générés par projet
    - versionnage WIP / FINAL / PRESENTED / SENT
    - consolidation multi-livrables
    """

    def __init__(self, db):
        self.db = db
        self.project_service = ProjectService(db)

    # ----------------------------------------------------------------------
    # UTILS
    # ----------------------------------------------------------------------

    def _project_output_dir(self, project_id: str) -> Path:
        p = self.project_service.get_project_dir(project_id) / "output"
        p.mkdir(parents=True, exist_ok=True)
        return p
    
    def _raw_output_dir(self, project_id: str) -> Path:
        """Dossier pour les livrables bruts"""
        p = self._project_output_dir(project_id) / "raw"
        p.mkdir(parents=True, exist_ok=True)
        return p
    
    def _workdocs_dir(self, project_id: str) -> Path:
        """Dossier pour les documents de travail"""
        p = self._project_output_dir(project_id) / "workdocs"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def _iteration_dir(self, project_id: str, template_name: str) -> Path:
        p = self._raw_output_dir(project_id) / template_name
        p.mkdir(parents=True, exist_ok=True)
        return p

    # ----------------------------------------------------------------------
    # ENREGISTREMENT D’UN NOUVEAU LIVRABLE GÉNÉRÉ
    # ----------------------------------------------------------------------

    def register_generated_output(
        self,
        project_id: str,
        template_name: str,
        ppt_path: Path | None,
        parameters: dict,
        excel_path: Path | None = None,
    ):

        """
        On est appelé juste après la génération dans generation_service.
        """
        iterations_dir = self._iteration_dir(project_id, template_name)

        iteration_id = uuid.uuid4().hex
        meta = {
            "id": iteration_id,
            "project_id": project_id,
            "template_name": template_name,
            "generated_at": datetime.now().isoformat(),
            "state": "WIP",
            "version": 1,
            "parameters": parameters,
            "ppt_path": str(ppt_path) if ppt_path else None,
            "excel_path": str(excel_path) if excel_path else None,

        }

        meta_file = iterations_dir / f"{iteration_id}.json"
        with open(meta_file, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        return iteration_id

    # ----------------------------------------------------------------------
    # LISTER LES LIVRABLES D’UN PROJET
    # ----------------------------------------------------------------------

    def list_iterations(self, project_id: str, template_name: str = None):
        data = []
        
        # 1. Charger livrables bruts depuis raw/
        raw_base = self._raw_output_dir(project_id)
        if raw_base.exists():
            for tmpl_dir in raw_base.iterdir():
                if tmpl_dir.is_dir():
                    tmpl = tmpl_dir.name
                    if template_name and tmpl != template_name:
                        continue
                    for meta_file in sorted(tmpl_dir.glob("*.json"), reverse=True):
                        with open(meta_file, "r", encoding="utf-8") as f:
                            data.append(json.load(f))
        
        # 2. Charger documents de travail depuis workdocs/
        workdocs_base = self._workdocs_dir(project_id)
        if workdocs_base.exists():
            for wd_dir in workdocs_base.iterdir():
                if wd_dir.is_dir():
                    meta_file = wd_dir / "meta.json"
                    if meta_file.exists():
                        with open(meta_file, "r", encoding="utf-8") as f:
                            data.append(json.load(f))

        data.sort(key=lambda x: x.get("generated_at") or x.get("created_at") or "", reverse=True)
        return data


    # ----------------------------------------------------------------------
    # CHANGEMENT D’ÉTAT / VERSIONNAGE
    # ----------------------------------------------------------------------

    def update_state(self, project_id: str, iteration_id: str, new_state: str):
        """Cherche dans raw/ et workdocs/"""
        state_map = {"WIP": "WIP", "FINAL": "FINALISE", "PRESENTED": "PRESENTE", "SENT": "ENVOYE"}
        
        # Chercher dans raw/
        raw_base = self._raw_output_dir(project_id)
        for tmpl_dir in raw_base.iterdir():
            if not tmpl_dir.is_dir():
                continue
            for meta_file in tmpl_dir.glob("*.json"):
                with open(meta_file, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                if meta["id"] == iteration_id:
                    meta["state"] = new_state
                    meta["version"] += 1
                    with open(meta_file, "w", encoding="utf-8") as f:
                        json.dump(meta, f, ensure_ascii=False, indent=2)
                    return meta
        
        # Chercher dans workdocs/
        workdocs_base = self._workdocs_dir(project_id)
        for wd_dir in workdocs_base.iterdir():
            if not wd_dir.is_dir():
                continue
            meta_file = wd_dir / "meta.json"
            if meta_file.exists():
                with open(meta_file, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                if meta["id"] == iteration_id:
                    meta["state"] = new_state
                    meta["version"] += 1
                    # Renommer fichier
                    if "label" in meta and "ppt_path" in meta:
                        old_file = Path(meta["ppt_path"])
                        if old_file.exists():
                            new_file = old_file.parent / f"V{meta['version']:03d}_{state_map.get(new_state, new_state)}_{meta['label']}.pptx"
                            import shutil
                            shutil.move(str(old_file), str(new_file))
                            meta["ppt_path"] = str(new_file)
                    with open(meta_file, "w", encoding="utf-8") as f:
                        json.dump(meta, f, ensure_ascii=False, indent=2)
                    return meta
        
        raise ValueError("Iteration introuvable")

    # ----------------------------------------------------------------------
    # CONSOLIDATION MULTI-LIVRABLES → WORKDOC
    # ----------------------------------------------------------------------

    def consolidate(self, project_id: str, iteration_ids: list, workdoc_name: str):
        """
        Concat PPT dans l’ordre donné.
        """
        all_iters = self.list_iterations(project_id)

        chosen = [x for x in all_iters if x["id"] in iteration_ids]
        if not chosen:
            raise ValueError("Aucune itération sélectionnée")

        output_dir = self._workdocs_dir(project_id) / workdoc_name
        output_dir.mkdir(parents=True, exist_ok=True)

        version = 1
        result_file = output_dir / f"V{version:03d}_WIP_{workdoc_name}.pptx"

        merge_presentations([Path(x["ppt_path"]) for x in chosen], result_file)

        meta = {
            "id": uuid.uuid4().hex,
            "project_id": project_id,
            "label": workdoc_name,
            "created_at": datetime.now().isoformat(),
            "state": "WIP",
            "version": version,
            "source_iterations": iteration_ids,
            "ppt_path": str(result_file),
        }

        with open(output_dir / "meta.json", "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        return meta

    def create_new_version(self, project_id: str, workdoc_id: str):
        """Crée une nouvelle version d'un workdoc."""
        base = self._workdocs_dir(project_id)
        state_map = {"WIP": "WIP", "FINAL": "FINALISE", "PRESENTED": "PRESENTE", "SENT": "ENVOYE"}
        
        for wd_dir in base.iterdir():
            if not wd_dir.is_dir():
                continue
            meta_file = wd_dir / "meta.json"
            if not meta_file.exists():
                continue
            with open(meta_file, "r", encoding="utf-8") as f:
                meta = json.load(f)
            if meta["id"] == workdoc_id:
                new_v = meta["version"] + 1
                label = meta["label"]
                state_str = state_map.get(meta.get("state", "WIP"), "WIP")
                new_file = wd_dir / f"V{new_v:03d}_{state_str}_{label}.pptx"
                import shutil
                shutil.copy2(Path(meta["ppt_path"]), new_file)
                meta["version"] = new_v
                meta["ppt_path"] = str(new_file)
                meta["created_at"] = datetime.now().isoformat()
                with open(meta_file, "w", encoding="utf-8") as f:
                    json.dump(meta, f, ensure_ascii=False, indent=2)
                return meta
        raise ValueError("Workdoc introuvable")

    def delete_workdoc(self, project_id: str, workdoc_id: str):
        """Supprime un workdoc (soft delete)."""
        base = self._workdocs_dir(project_id)
        trash = base.parent / "_trash_workdocs"
        trash.mkdir(exist_ok=True)
        for wd_dir in base.iterdir():
            if not wd_dir.is_dir() or wd_dir.name == "_trash":
                continue
            meta_file = wd_dir / "meta.json"
            if not meta_file.exists():
                continue
            with open(meta_file, "r", encoding="utf-8") as f:
                meta = json.load(f)
            if meta["id"] == workdoc_id:
                meta["deleted"] = True
                meta["deleted_at"] = datetime.now().isoformat()
                with open(meta_file, "w", encoding="utf-8") as f:
                    json.dump(meta, f, ensure_ascii=False, indent=2)
                import shutil
                shutil.move(str(wd_dir), str(trash / wd_dir.name))
                return True
        raise ValueError("Workdoc introuvable")

    def delete_iteration(self, project_id: str, iteration_id: str):
        """Supprime un livrable brut."""
        base = self._raw_output_dir(project_id)
        trash = base.parent / "_trash_iterations"
        trash.mkdir(exist_ok=True)
        for tmpl_dir in base.iterdir():
            if not tmpl_dir.is_dir():
                continue
            for meta_file in tmpl_dir.glob("*.json"):
                with open(meta_file, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                if meta["id"] == iteration_id:
                    import shutil
                    ppt = Path(meta.get("ppt_path", ""))
                    excel = Path(meta.get("excel_path", ""))
                    if ppt.exists():
                        shutil.move(str(ppt), str(trash / ppt.name))
                    if excel.exists():
                        shutil.move(str(excel), str(trash / excel.name))
                    shutil.move(str(meta_file), str(trash / meta_file.name))
                    return True
        raise ValueError("Iteration introuvable")