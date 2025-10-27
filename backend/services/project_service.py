# backend/services/project_service.py
from __future__ import annotations

from typing import Optional, Dict, Any, List, Tuple
from pathlib import Path
import json
import shutil
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from loguru import logger

# --- Dépendances internes ---
try:
    from backend.services.template_service import TemplateService
except Exception:
    from services.template_service import TemplateService

try:
    from backend.services.dataset_service import (
        align_df_to_expected_columns,
        _load_dataframe_from_source,
        _apply_source_python
    )
except Exception:
    def align_df_to_expected_columns(df: pd.DataFrame, expected_columns: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        expected = [c for c in (expected_columns or []) if isinstance(c, str) and c.strip()]
        cur_cols = list(df.columns)
        missing = [c for c in expected if c not in cur_cols]
        for c in missing:
            df[c] = pd.NA
        ordered = expected + [c for c in df.columns if c not in expected]
        return df[ordered], {"missing": missing, "extra": [c for c in cur_cols if c not in expected]}
    
    def _load_dataframe_from_source(source: dict) -> Optional[pd.DataFrame]:
        return None
    
    def _apply_source_python(df: pd.DataFrame, source: dict) -> pd.DataFrame:
        return df


try:
    from backend.services.gabarit_registry import get_default_source
except Exception:
    def get_default_source(gabarit_name: str, gabarit_version: str):
        return None
    

try:
    from backend.utils.file_utils import ensure_directories
except Exception:
    def ensure_directories(*paths: Path) -> None:
        for p in paths:
            Path(p).parent.mkdir(parents=True, exist_ok=True)


# ==================== CONFIGURATION CHEMINS ====================
PARIS = ZoneInfo("Europe/Paris")
CONF_DIR = Path("configuration")
PROJECTS_DIR = CONF_DIR / "projets"
TRASH_DIR = PROJECTS_DIR / "_trash"

# Assurer l'existence
PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
TRASH_DIR.mkdir(parents=True, exist_ok=True)


# ==================== HELPERS ====================

def _now_paris_iso() -> str:
    """Horodatage Europe/Paris en ISO (secondes)."""
    return datetime.now(PARIS).isoformat(timespec="seconds")


def _slugify(name: str) -> str:
    """Slug stable pour générer un project_id."""
    s = "".join(ch if ch.isalnum() or ch in "-_" else "-" for ch in name.strip())
    while "--" in s:
        s = s.replace("--", "-")
    return s.strip("-_").lower() or f"project-{datetime.now(PARIS).strftime('%Y%m%d%H%M%S')}"


def _project_dir(project_id: str) -> Path:
    """Chemin du dossier d'un projet actif."""
    return PROJECTS_DIR / project_id


def _project_config_path(project_id: str) -> Path:
    """Chemin du fichier config.json d'un projet actif."""
    return _project_dir(project_id) / "config.json"


def _trash_path_for_project(base_name: str) -> Path:
    """
    Génère un nom unique dans _trash en ajoutant un suffixe si collision.
    Format : {base_name}_Supr_n0001
    """
    prefix = f"{base_name}_Supr_n"
    
    # Trouver le prochain numéro disponible
    existing_nums = []
    for p in TRASH_DIR.iterdir():
        if p.is_dir() and p.name.startswith(prefix):
            try:
                num_str = p.name.split(prefix)[1]
                existing_nums.append(int(num_str))
            except (IndexError, ValueError):
                continue
    
    next_num = (max(existing_nums) + 1) if existing_nums else 1
    return TRASH_DIR / f"{prefix}{str(next_num).zfill(4)}"


# ==================== SERVICE PRINCIPAL ====================

class ProjectService:
    """
    Service de gestion des projets avec stockage file-based.
    
    Structure :
    configuration/
    └── projets/
        ├── {project_id}/
        │   ├── config.json
        │   ├── masters/
        │   │   └── {template_id}/
        │   │       ├── presentation.pptx
        │   │       └── donnees.xlsx
        │   └── cache/
        │       └── params_{template_id}.json
        └── _trash/
            └── {project_id}_Supr_n0001/
    """
    
    def __init__(self, db_session, template_service: Optional[TemplateService] = None):
        self.db = db_session
        self.ts = template_service or TemplateService(db_session)
    
    # ==================== CRUD PROJET ====================
    
    def create_project(self, name: str, description: str = "", client_name: str = "",
                       project_id: Optional[str] = None) -> Dict[str, Any]:
        """
        Crée un nouveau projet avec structure de dossiers.
        """
        pid = project_id or _slugify(name)
        
        # Créer la structure de dossiers
        project_dir = _project_dir(pid)
        project_dir.mkdir(parents=True, exist_ok=True)
        
        (project_dir / "masters").mkdir(exist_ok=True)
        (project_dir / "cache").mkdir(exist_ok=True)
        
        # Config initiale
        data: Dict[str, Any] = {
            "project_id": pid,
            "name": name,
            "client_name": client_name,
            "description": description,
            "status": "active",
            "deliverables": [],
            "data_sources": [],
            "created_at": _now_paris_iso(),
            "updated_at": _now_paris_iso(),
        }
        
        self.save_project(data)
        logger.success(f"Projet créé : {pid}")
        return data
    
    def load_project(self, project_id: str) -> Dict[str, Any]:
        """Charge la config d'un projet depuis config.json."""
        path = _project_config_path(project_id)
        if not path.exists():
            raise FileNotFoundError(f"Projet introuvable : {project_id}")
        
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    
    def save_project(self, project: Dict[str, Any]) -> None:
        """Sauvegarde la config d'un projet dans config.json."""
        project_id = project["project_id"]
        path = _project_config_path(project_id)
        
        project["updated_at"] = _now_paris_iso()
        
        ensure_directories(path)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(project, f, ensure_ascii=False, indent=2)
        
        logger.info(f"Projet sauvegardé : {path}")
    
    def list_projects(self) -> List[Dict[str, Any]]:
        """Liste uniquement les projets actifs (hors _trash)."""
        out: List[Dict[str, Any]] = []
        
        if not PROJECTS_DIR.exists():
            return out
        
        for project_dir in PROJECTS_DIR.iterdir():
            # Ignorer _trash et les fichiers
            if not project_dir.is_dir() or project_dir.name.startswith("_"):
                continue
            
            config_file = project_dir / "config.json"
            if not config_file.exists():
                continue
            
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    out.append(json.load(f))
            except Exception as e:
                logger.warning(f"Impossible de lire {config_file} : {e}")
                continue
        
        # Tri par date de mise à jour (plus récent en premier)
        return sorted(out, key=lambda x: x.get("updated_at", ""), reverse=True)
    

    def duplicate_project(self, src_project_id: str, new_name: str, new_client_name: str = "") -> Dict[str, Any]:
        """
        Duplique INTÉGRALEMENT un projet :
        - copie le dossier configuration/projets/{src}/ → {dst}/ (masters, cache, etc.)
        - met à jour config.json (project_id, name, client_name, status, dates)
        - renvoie le dict projet chargé (config.json)
        """
        new_name = (new_name or "").strip()
        if not new_name:
            raise ValueError("Nouveau nom de projet requis")

        # 1) Charger la source pour lire sa config (et vérifier existence)
        src_dir = _project_dir(src_project_id)
        if not src_dir.exists():
            raise FileNotFoundError(f"Projet source introuvable : {src_project_id}")

        # 2) Générer un project_id unique à partir du nom
        base_pid = _slugify(new_name)
        dst_pid = base_pid
        k = 1
        while (_project_dir(dst_pid)).exists():
            k += 1
            dst_pid = f"{base_pid}-{k}"

        dst_dir = _project_dir(dst_pid)

        # 3) Copie physique complète (masters, cache, etc.)
        shutil.copytree(src_dir, dst_dir)

        # 4) Mettre à jour config.json dans la copie
        cfg_path = dst_dir / "config.json"
        if not cfg_path.exists():
            raise FileNotFoundError(f"config.json manquant dans la copie : {cfg_path}")

        with open(cfg_path, "r", encoding="utf-8") as f:
            data = json.load(f) or {}

        # Champs à remettre à jour
        data["project_id"] = dst_pid
        data["name"] = new_name
        data["client_name"] = (new_client_name or "").strip()
        data["status"] = "active"
        data["created_at"] = _now_paris_iso()
        data["updated_at"] = _now_paris_iso()

        # (Option) si tu veux nettoyer des traces spécifiques, fais-le ici :
        # ex: data.pop("archived_at", None)

        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        logger.success(f"Projet dupliqué : {src_project_id} → {dst_pid}")

        # 5) Retourner l'objet projet (chargé)
        return self.load_project(dst_pid)


    # ==================== ARCHIVAGE ====================
    
    def soft_delete(self, project_id: str) -> str:
        """
        Archive un projet dans _trash avec renommage unique.
        Retourne le nouveau nom du projet archivé.
        """
        src_dir = _project_dir(project_id)
        
        if not src_dir.exists():
            raise FileNotFoundError(f"Projet introuvable : {project_id}")
        
        # Charger la config pour mettre à jour le nom
        config_file = src_dir / "config.json"
        if not config_file.exists():
            raise FileNotFoundError(f"Config introuvable : {config_file}")
        
        with open(config_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Générer le nouveau nom unique
        original_name = data.get("name", project_id)
        dest_dir = _trash_path_for_project(project_id)
        
        # Renommer dans la config
        new_display_name = dest_dir.name  # Ex: projet_Supr_n0001
        data["name"] = f"{original_name}_Supr_n{dest_dir.name.split('_Supr_n')[1]}"
        data["status"] = "archived"
        data["archived_at"] = _now_paris_iso()
        data["updated_at"] = _now_paris_iso()
        
        # Sauvegarder la config modifiée
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # Déplacer le dossier complet
        shutil.move(str(src_dir), str(dest_dir))
        
        logger.success(f"Projet archivé : {src_dir} → {dest_dir}")
        return new_display_name
    
    def restore(self, archived_name: str) -> str:
        """
        Restaure un projet depuis _trash.
        archived_name : nom du dossier dans _trash (ex: projet_Supr_n0001)
        Retourne le project_id restauré.
        """
        src_dir = TRASH_DIR / archived_name
        
        if not src_dir.exists():
            raise FileNotFoundError(f"Projet archivé introuvable : {archived_name}")
        
        # Charger la config pour récupérer le project_id original
        config_file = src_dir / "config.json"
        with open(config_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        original_id = data.get("project_id")
        if not original_id:
            raise ValueError("project_id manquant dans la config")
        
        # Vérifier qu'un projet avec cet ID n'existe pas déjà
        dest_dir = _project_dir(original_id)
        if dest_dir.exists():
            raise FileExistsError(f"Un projet '{original_id}' existe déjà")
        
        # Restaurer le nom original
        original_name = data["name"].rsplit("_Supr_n", 1)[0]
        data["name"] = original_name
        data["status"] = "active"
        data.pop("archived_at", None)
        data["updated_at"] = _now_paris_iso()
        
        # Sauvegarder
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # Déplacer
        shutil.move(str(src_dir), str(dest_dir))
        
        logger.success(f"Projet restauré : {archived_name} → {original_id}")
        return original_id
    
    def list_trash(self) -> List[Dict[str, Any]]:
        """Liste les projets archivés dans _trash."""
        out: List[Dict[str, Any]] = []
        
        if not TRASH_DIR.exists():
            return out
        
        for project_dir in TRASH_DIR.iterdir():
            if not project_dir.is_dir():
                continue
            
            config_file = project_dir / "config.json"
            if not config_file.exists():
                continue
            
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    data["_archived_folder_name"] = project_dir.name
                    out.append(data)
            except Exception as e:
                logger.warning(f"Impossible de lire {config_file} : {e}")
                continue
        
        return sorted(out, key=lambda x: x.get("archived_at", x.get("updated_at", "")), reverse=True)
    
    # ==================== LIVRABLES ====================
    
    def add_deliverable(self, project_id: str, template_id: int) -> Dict[str, Any]:
        """Ajoute un livrable au projet et duplique les masters."""
        proj = self.load_project(project_id)
        
        # Vérifier que le template existe
        tpl = self.ts.get_template(template_id)
        if not tpl:
            raise ValueError(f"Template {template_id} introuvable")
        
        # Vérifier si déjà présent
        if any(d.get("template_id") == template_id for d in proj.get("deliverables", [])):
            logger.warning(f"Template {template_id} déjà dans le projet")
            return proj
        
        # Dupliquer les masters
        masters_paths = self._duplicate_masters(project_id, template_id)
        
        # Créer l'entrée deliverable
        deliverable = {
            "template_id": template_id,
            "template_name": tpl.name,
            "template_version": tpl.version,
            "custom_masters": masters_paths,
            "custom_parameters": {},
            "is_functional": False,
            "completion_rate": 0.0,
            "data_sources_status": {"client": 0, "default": 0},
            "last_generated_at": None
        }
        
        deliverables = proj.get("deliverables", [])
        deliverables.append(deliverable)
        proj["deliverables"] = deliverables
        
        self.save_project(proj)
        logger.success(f"Livrable {tpl.name} ajouté au projet {project_id}")
        return deliverable
    
    def remove_deliverable(self, project_id: str, template_id: int) -> None:
        """Retire un livrable du projet (conserve les masters pour historique)."""
        proj = self.load_project(project_id)
        deliverables = [d for d in proj.get("deliverables", []) if d.get("template_id") != template_id]
        proj["deliverables"] = deliverables
        self.save_project(proj)
        logger.info(f"Livrable {template_id} retiré du projet {project_id}")
    
    def list_deliverables(self, project_id: str) -> List[Dict[str, Any]]:
        """Liste tous les livrables du projet avec leurs statuts."""
        proj = self.load_project(project_id)
        return proj.get("deliverables", [])
    
    def _duplicate_masters(self, project_id: str, template_id: int) -> Dict[str, str]:
        """
        Copie les masters PPT/Excel du template vers le dossier du projet.
        Retourne les chemins des fichiers dupliqués.
        """
        tpl = self.ts.get_template(template_id)
        if not tpl:
            raise ValueError(f"Template {template_id} introuvable")
        
        # Destination
        dest_dir = _project_dir(project_id) / "masters" / str(template_id)
        dest_dir.mkdir(parents=True, exist_ok=True)
        
        masters_paths = {}
        
        # Copier PPT
        if tpl.ppt_template_path and Path(tpl.ppt_template_path).exists():
            src_ppt = Path(tpl.ppt_template_path)
            dest_ppt = dest_dir / f"presentation{src_ppt.suffix}"
            shutil.copy2(src_ppt, dest_ppt)
            masters_paths["ppt_path"] = str(dest_ppt)
            logger.info(f"PPT dupliqué : {dest_ppt}")
        
        # Copier Excel
        if tpl.excel_template_path and Path(tpl.excel_template_path).exists():
            src_excel = Path(tpl.excel_template_path)
            dest_excel = dest_dir / f"donnees{src_excel.suffix}"
            shutil.copy2(src_excel, dest_excel)
            masters_paths["excel_path"] = str(dest_excel)
            logger.info(f"Excel dupliqué : {dest_excel}")
        
        return masters_paths
    
    def reset_master(self, project_id: str, template_id: int, file_type: str) -> str:
        """
        Réinitialise un master (PPT ou Excel) depuis le template original.
        file_type: "ppt" ou "excel"
        """
        tpl = self.ts.get_template(template_id)
        if not tpl:
            raise ValueError(f"Template {template_id} introuvable")
        
        dest_dir = _project_dir(project_id) / "masters" / str(template_id)
        
        if file_type == "ppt":
            src = Path(tpl.ppt_template_path)
            dest = dest_dir / f"presentation{src.suffix}"
        elif file_type == "excel":
            src = Path(tpl.excel_template_path)
            dest = dest_dir / f"donnees{src.suffix}"
        else:
            raise ValueError(f"file_type invalide: {file_type}")
        
        if not src.exists():
            raise FileNotFoundError(f"Fichier source introuvable: {src}")
        
        shutil.copy2(src, dest)
        logger.success(f"Master {file_type} réinitialisé depuis le template")
        return str(dest)
    
    # ==================== SOURCES DE DONNÉES ====================
    
    def set_data_source(self, project_id: str, gabarit_name: str, gabarit_version: str,
                       source_type: str, source_config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Configure la source de données pour un gabarit.
        source_type: "default" ou "client"
        source_config: {"type": "csv", "path": "...", "sep": ";", ...}
        """
        proj = self.load_project(project_id)
        sources = proj.get("data_sources", [])
        
        # Supprimer l'existant
        sources = [s for s in sources if not (
            s.get("gabarit_name") == gabarit_name and 
            s.get("gabarit_version") == gabarit_version
        )]
        
        # Ajouter la nouvelle source
        source = {
            "gabarit_name": gabarit_name,
            "gabarit_version": gabarit_version,
            "source_type": source_type,
            "source_config": source_config,
            "row_count": None,
            "columns_filled": {},
            "last_validated_at": None
        }
        
        sources.append(source)
        proj["data_sources"] = sources
        self.save_project(proj)
        
        return source
    
    def get_data_source(self, project_id: str, gabarit_name: str, gabarit_version: str) -> Optional[Dict[str, Any]]:
        """Récupère la source configurée pour un gabarit."""
        proj = self.load_project(project_id)
        for s in proj.get("data_sources", []):
            if s.get("gabarit_name") == gabarit_name and s.get("gabarit_version") == gabarit_version:
                return s
        return None
    
    def validate_data_source(self, project_id: str, gabarit_name: str, gabarit_version: str) -> Dict[str, Any]:
        """Valide une source de données (colonnes, lignes, complétude)."""
        source = self.get_data_source(project_id, gabarit_name, gabarit_version)
        if not source:
            raise ValueError("Source non configurée")
        
        # Charger un échantillon
        df, profile = self._load_and_profile_source(source["source_config"], head=1000)
        
        # Calculer complétude par colonne
        columns_filled = {
            col: round((1 - df[col].isna().mean()) * 100, 1)
            for col in df.columns
        }
        
        # Mettre à jour la source
        proj = self.load_project(project_id)
        for s in proj.get("data_sources", []):
            if s.get("gabarit_name") == gabarit_name and s.get("gabarit_version") == gabarit_version:
                s["row_count"] = len(df)
                s["columns_filled"] = columns_filled
                s["last_validated_at"] = _now_paris_iso()
                break
        
        self.save_project(proj)
        
        return {
            "row_count": len(df),
            "columns": list(df.columns),
            "columns_filled": columns_filled,
            "profile": profile
        }
    

    def build_dataframe(
        self,
        project_id: str,
        gabarit_name: Optional[str] = None,
        gabarit_version: str = "v1",
        expected_columns: Optional[List[str]] = None,
        *,
        usage: Optional[Dict[str, Any]] = None,
        template_id: Optional[int] = None,
        limit: Optional[int] = None
    ) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        """
        Construit le DataFrame en mode PROJET pour un gabarit donné.
        - Résout la source: PROJET (client) prioritaire, sinon source DÉFAUT du gabarit
        - Charge le DF via _load_dataframe_from_source
        - Applique le code Python éventuel (_apply_source_python)
        - Aligne/filtre les colonnes si expected_columns est fourni
        Retourne: (df, meta)
        """

        # 1) Support d'appel via `usage` (ReportService transmet souvent un dict usage)
        if usage:
            gabarit_name = usage.get("gabarit_name") or gabarit_name
            gabarit_version = usage.get("gabarit_version") or gabarit_version or "v1"
            cols_from_usage = usage.get("columns_enabled")
            if cols_from_usage and isinstance(cols_from_usage, list) and cols_from_usage:
                expected_columns = cols_from_usage

        if not gabarit_name:
            raise ValueError("build_dataframe: gabarit_name obligatoire")

        # 2) Résoudre la source: PROJET (client) > DÉFAUT GABARIT
        source_entry = self.get_data_source(project_id, gabarit_name, gabarit_version)
        source_cfg = None

        if source_entry and isinstance(source_entry.get("source_config"), dict):
            source_cfg = dict(source_entry["source_config"])  # shallow copy
            logger.debug(f"[ProjectService] Source PROJET utilisée pour {gabarit_name}:{gabarit_version}")
        else:
            default_src = get_default_source(gabarit_name, gabarit_version)
            if default_src:
                source_cfg = dict(default_src)
                logger.debug(f"[ProjectService] Source DÉFAUT utilisée pour {gabarit_name}:{gabarit_version}")

        if not source_cfg:
            raise RuntimeError(
                f"Aucune source disponible pour {gabarit_name}:{gabarit_version} "
                f"(ni côté projet, ni côté gabarit)."
            )

        # 3) Charger le DataFrame depuis la source
        df = _load_dataframe_from_source(source_cfg)
        if df is None:
            raise RuntimeError(
                f"Échec de chargement de la source pour {gabarit_name}:{gabarit_version}."
            )

        # 4) (désactivé) : le code Python de source est déjà appliqué au chargement
        # try:
        #     df = _apply_source_python(df, source_cfg)
        # except Exception as e:
        #     logger.warning(f"[ProjectService] _apply_source_python a échoué: {e}")


        # 5) Optionnel: limiter (utile pour des previews/tests)
        if isinstance(limit, int) and limit > 0:
            df = df.head(limit)

        # 6) Aligner/filtrer les colonnes si attendu fourni
        meta: Dict[str, Any] = {}
        if expected_columns and isinstance(expected_columns, list) and expected_columns:
            try:
                df, meta = align_df_to_expected_columns(df, expected_columns)
            except Exception as e:
                logger.warning(f"[ProjectService] align_df_to_expected_columns a échoué: {e}")
                # fallback: garder df brut pour ne pas bloquer

            # Filtrer strictement si la liste attendue est fournie
            try:
                keep = [c for c in expected_columns if c in df.columns]
                if keep:
                    df = df[keep]
            except Exception:
                pass

        # 7) Méta + retour
        meta.setdefault("rows", len(df))
        meta.setdefault("columns", list(df.columns))
        meta.setdefault("source_type", source_entry["source_type"] if source_entry else "default")
        try:
            df.attrs["kaivaa_meta"] = meta
        except Exception:
            pass
        return df

    
    def _load_and_profile_source(self, source_config: Dict[str, Any], head: int = 1000) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        """Charge une source et génère un profil."""
        df = _load_dataframe_from_source(source_config)
        if df is None:
            raise ValueError("Impossible de charger la source")
        
        df = df.head(head)
        
        profile = {
            "rows": len(df),
            "cols": len(df.columns),
            "dtypes": {c: str(df[c].dtype) for c in df.columns}
        }
        
        return df, profile
    
    # ==================== STATUTS LIVRABLES ====================
    
    def compute_deliverable_status(self, project_id: str, template_id: int) -> Dict[str, Any]:
        """
        Calcule le statut de complétude d'un livrable.
        """
        # 1. Récupérer les gabarits requis par ce template
        usages = self.ts.list_gabarit_usages(template_id)
        
        required_gabarits = set()
        for u in usages:
            gname = u.get("gabarit_name")
            gver = u.get("gabarit_version", "v1")
            if gname:
                required_gabarits.add((gname, gver))
        
        if not required_gabarits:
            return {
                "is_functional": True,
                "completion_rate": 100.0,
                "data_sources_status": {"client": 0, "default": 0},
                "missing_gabarits": []
            }
        
        # 2. Vérifier les sources configurées dans le projet
        proj = self.load_project(project_id)
        configured_sources = {}
        
        for s in proj.get("data_sources", []):
            key = (s.get("gabarit_name"), s.get("gabarit_version"))
            configured_sources[key] = s.get("source_type", "default")
        
        # 3. Calculer complétude
        client_count = 0
        default_count = 0
        missing = []
        
        for (gname, gver) in required_gabarits:
            source_type = configured_sources.get((gname, gver))
            
            if source_type == "client":
                client_count += 1
            elif source_type == "default":
                default_count += 1
            else:
                # Vérifier si une source par défaut existe dans le gabarit
                from backend.services.gabarit_registry import get_default_source
                default_src = get_default_source(gname, gver)
                if default_src:
                    default_count += 1
                else:
                    missing.append((gname, gver))
        
        total = len(required_gabarits)
        configured = client_count + default_count
        completion_rate = round((configured / total) * 100, 1) if total > 0 else 0
        
        return {
            "is_functional": len(missing) == 0,
            "completion_rate": completion_rate,
            "data_sources_status": {
                "client": client_count,
                "default": default_count
            },
            "missing_gabarits": [f"{g[0]} (v{g[1]})" for g in missing]
        }
    
    def update_deliverable_status(self, project_id: str, template_id: int) -> None:
        """Met à jour le statut d'un livrable dans le projet."""
        status = self.compute_deliverable_status(project_id, template_id)
        
        proj = self.load_project(project_id)
        for d in proj.get("deliverables", []):
            if d.get("template_id") == template_id:
                d["is_functional"] = status["is_functional"]
                d["completion_rate"] = status["completion_rate"]
                d["data_sources_status"] = status["data_sources_status"]
                break
        
        self.save_project(proj)