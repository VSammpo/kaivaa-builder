# backend/services/transformation_service.py
"""
Service de gestion des transformations réutilisables.
Stockage file-based dans configuration/transformations/
"""
from __future__ import annotations
from pathlib import Path
import json
from typing import List, Optional, Dict, Any
from datetime import datetime
from zoneinfo import ZoneInfo
from loguru import logger
import pandas as pd

# Dépendances internes
from backend.services.gabarit_registry import get_gabarit

# === CONFIGURATION CHEMINS ===
CONF_DIR = Path("configuration")
TRANSFO_DIR = CONF_DIR / "transformations"
TRASH_DIR = TRANSFO_DIR / "_trash"

def _ensure_dirs() -> None:
    """Crée les dossiers nécessaires"""
    TRANSFO_DIR.mkdir(parents=True, exist_ok=True)
    TRASH_DIR.mkdir(parents=True, exist_ok=True)

def _transfo_path(name: str, version: str) -> Path:
    """Chemin vers le fichier config.json d'une transformation"""
    return TRANSFO_DIR / name / version / "config.json"

def _now_paris_iso() -> str:
    """Horodatage Europe/Paris en ISO"""
    return datetime.now(ZoneInfo("Europe/Paris")).isoformat(timespec="seconds")

def _read_json(p: Path) -> dict:
    """Lit un fichier JSON"""
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        logger.error(f"Erreur lecture JSON {p}: {e}")
        return {}

def _write_json(p: Path, data: dict) -> None:
    """Écrit un fichier JSON de manière atomique"""
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)

def _normalize_transformation(data: dict) -> dict:
    """Normalise une transformation avec valeurs par défaut"""
    data = dict(data or {})
    data.setdefault("name", "")
    data.setdefault("version", "v1")
    data.setdefault("description", "")
    data.setdefault("gabarit_base", {"name": "", "version": "v1"})
    data.setdefault("columns_enabled", [])
    data.setdefault("enrichments", [])
    data.setdefault("methods", [])
    data.setdefault("overlay_python", "")
    data.setdefault("final_order", [])
    data.setdefault("final_excludes", [])
    data.setdefault("final_renames", {})
    data.setdefault("final_sort", [])
    data.setdefault("created_at", _now_paris_iso())
    data.setdefault("updated_at", _now_paris_iso())
    return data


# ==================== CRUD TRANSFORMATIONS ====================

def list_transformations() -> List[Dict[str, Any]]:
    """
    Liste toutes les transformations actives (hors _trash).
    Retourne une liste de dicts avec métadonnées.
    """
    _ensure_dirs()
    results: List[Dict[str, Any]] = []
    
    if not TRANSFO_DIR.exists():
        return results
    
    for transfo_dir in sorted(TRANSFO_DIR.iterdir()):
        if not transfo_dir.is_dir() or transfo_dir.name.startswith("_"):
            continue
        
        # Parcourir les versions
        for version_dir in sorted(transfo_dir.iterdir()):
            if not version_dir.is_dir():
                continue
            
            config_file = version_dir / "config.json"
            if not config_file.exists():
                continue
            
            try:
                data = _normalize_transformation(_read_json(config_file))
                results.append(data)
            except Exception as e:
                logger.warning(f"Impossible de lire {config_file}: {e}")
                continue
    
    # Tri par date de mise à jour (plus récent en premier)
    return sorted(results, key=lambda x: x.get("updated_at", ""), reverse=True)


def get_transformation(name: str, version: str = "v1") -> Optional[Dict[str, Any]]:
    """
    Récupère une transformation par nom et version.
    Retourne None si introuvable.
    """
    p = _transfo_path((name or "").strip(), (version or "v1").strip())
    if not p.exists():
        return None
    
    try:
        data = _normalize_transformation(_read_json(p))
        return data
    except Exception as e:
        logger.error(f"Erreur chargement transformation {name} v{version}: {e}")
        return None


def create_transformation(
    name: str,
    version: str = "v1",
    description: str = "",
    gabarit_base: Dict[str, str] = None,
    columns_enabled: List[str] = None,
    enrichments: List[Dict] = None,
    methods: List[str] = None,
    overlay_python: str = "",
    final_order: List[str] = None,
    final_excludes: List[str] = None,
    final_renames: Dict[str, str] = None,
    final_sort: List[Dict] = None,
) -> Dict[str, Any]:
    """
    Crée une nouvelle transformation.
    Lève ValueError si elle existe déjà.
    """
    _ensure_dirs()
    
    name = (name or "").strip()
    version = (version or "v1").strip()
    
    if not name:
        raise ValueError("Le nom de la transformation est obligatoire")
    
    # Vérifier unicité
    existing = get_transformation(name, version)
    if existing:
        raise ValueError(f"La transformation '{name}' v{version} existe déjà")
    
    # Construire la transformation
    data = {
        "name": name,
        "version": version,
        "description": description or "",
        "gabarit_base": gabarit_base or {"name": "", "version": "v1"},
        "columns_enabled": list(columns_enabled or []),
        "enrichments": list(enrichments or []),
        "methods": list(methods or []),
        "overlay_python": (overlay_python or "").strip(),
        "final_order": list(final_order or []),
        "final_excludes": list(final_excludes or []),
        "final_renames": dict(final_renames or {}),
        "final_sort": list(final_sort or []),
        "created_at": _now_paris_iso(),
        "updated_at": _now_paris_iso(),
    }
    
    # Sauvegarder
    p = _transfo_path(name, version)
    _write_json(p, data)
    
    logger.success(f"Transformation créée: {name} v{version} -> {p}")
    return data


def update_transformation(
    name: str,
    version: str = "v1",
    updates: Dict[str, Any] = None
) -> Dict[str, Any]:
    """
    Met à jour une transformation existante.
    Lève ValueError si introuvable.
    """
    existing = get_transformation(name, version)
    if not existing:
        raise ValueError(f"Transformation '{name}' v{version} introuvable")
    
    # Fusionner les updates
    merged = dict(existing)
    for key, value in (updates or {}).items():
        if key not in ["name", "version", "created_at"]:  # Champs protégés
            merged[key] = value
    
    merged["updated_at"] = _now_paris_iso()
    
    # Sauvegarder
    p = _transfo_path(name, version)
    _write_json(p, merged)
    
    logger.info(f"Transformation mise à jour: {name} v{version}")
    return merged


def delete_transformation(name: str, version: str = "v1", hard_delete: bool = False) -> Dict[str, Any]:
    """
    Supprime une transformation (soft delete vers _trash par défaut).
    
    Args:
        name: Nom de la transformation
        version: Version
        hard_delete: Si True, suppression définitive
    
    Returns:
        Dict avec {"old_name", "new_name", "archived": True/False}
    """
    name = (name or "").strip()
    version = (version or "v1").strip()
    
    p = _transfo_path(name, version)
    if not p.exists():
        raise FileNotFoundError(f"Transformation '{name}' v{version} introuvable")
    
    if hard_delete:
        # Suppression définitive
        logger.warning(f"Suppression DÉFINITIVE de la transformation '{name}' v{version}")
        import shutil
        version_dir = p.parent
        shutil.rmtree(version_dir)
        
        # Nettoyer le dossier parent si vide
        try:
            version_dir.parent.rmdir()
        except:
            pass
        
        logger.success(f"Transformation supprimée définitivement: {name} v{version}")
        return {"old_name": name, "new_name": None, "archived": False}
    
    else:
        # Soft delete vers _trash
        return soft_delete_transformation(name, version)


def soft_delete_transformation(name: str, version: str = "v1") -> Dict[str, Any]:
    """
    Archive une transformation vers _trash avec suffixe unique.
    
    Returns:
        Dict avec {"old_name", "new_name"}
    """
    _ensure_dirs()
    
    name = (name or "").strip()
    version = (version or "v1").strip()
    
    p = _transfo_path(name, version)
    if not p.exists():
        raise FileNotFoundError(f"Transformation '{name}' v{version} introuvable")
    
    # Charger la config pour mettre à jour le nom
    data = _read_json(p)
    
    # Générer le nouveau nom unique
    prefix = f"{name}_Supr_n"
    existing_nums = []
    for d in TRASH_DIR.iterdir():
        if d.is_dir() and d.name.startswith(prefix):
            try:
                num_str = d.name.split(prefix)[1]
                existing_nums.append(int(num_str))
            except:
                continue
    
    next_num = (max(existing_nums) + 1) if existing_nums else 1
    new_name = f"{prefix}{str(next_num).zfill(4)}"
    
    # Mettre à jour le nom dans la config
    data["name"] = new_name
    data["updated_at"] = _now_paris_iso()
    
    # Déplacer vers _trash
    dest_dir = TRASH_DIR / new_name / version
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_file = dest_dir / "config.json"
    
    _write_json(dest_file, data)
    
    # Supprimer l'original
    import shutil
    try:
        p.unlink()
        try:
            p.parent.rmdir()  # Version dir
            try:
                p.parent.parent.rmdir()  # Name dir
            except:
                pass
        except:
            pass
    except Exception as e:
        logger.warning(f"Impossible de supprimer l'original: {e}")
    
    logger.success(f"Transformation archivée: {name} -> {new_name}")
    return {"old_name": name, "new_name": new_name}


# ==================== UTILITAIRES ====================

def get_transformation_input_columns(name: str, version: str = "v1") -> List[str]:
    """
    Retourne les colonnes MINIMALES requises en entrée pour cette transformation.
    (colonnes du gabarit de base + left_keys des enrichissements + inputs des méthodes)
    """
    transfo = get_transformation(name, version)
    if not transfo:
        return []
    
    # Colonnes de base
    gabarit_base = transfo.get("gabarit_base", {})
    gab_name = gabarit_base.get("name", "")
    gab_ver = gabarit_base.get("version", "v1")
    
    if not gab_name:
        return []
    
    gab = get_gabarit(gab_name, gab_ver)
    if not gab:
        return []
    
    base_cols = [c.name for c in (gab.columns or [])]
    required = list(transfo.get("columns_enabled") or base_cols[:])
    
    # Ajouter les left_key des enrichissements (premier saut uniquement)
    for e in (transfo.get("enrichments") or []):
        path = e.get("path") or []
        if path:
            first_left_key = path[0][1]  # [from, left_key, to, right_key]
            if first_left_key and first_left_key not in required:
                required.append(first_left_key)
    
    # Ajouter les inputs des méthodes
    from backend.services.gabarit_registry import list_methods_for_gabarit
    sel_methods = set(transfo.get("methods") or [])
    if sel_methods:
        all_methods = list_methods_for_gabarit(gab_name, gab_ver) or []
        methods_iter = (all_methods.values() if isinstance(all_methods, dict) else all_methods)
        for m in methods_iter:
            if isinstance(m, dict) and m.get("name") in sel_methods:
                for col_input in (m.get("required_columns") or []):
                    if col_input and col_input not in required:
                        required.append(col_input)
    
    return list(dict.fromkeys(required))  # Déduplication en conservant l'ordre


def get_transformation_output_columns(name: str, version: str = "v1") -> List[str]:
    """
    Retourne les colonnes de SORTIE après application de la transformation.
    (ordre final, renommages appliqués, exclusions retirées)
    """
    transfo = get_transformation(name, version)
    if not transfo:
        return []
    
    # Ordre final
    final_order = list(transfo.get("final_order") or [])
    final_excludes = set(transfo.get("final_excludes") or [])
    final_renames = dict(transfo.get("final_renames") or {})
    
    # Appliquer exclusions
    output_cols = [c for c in final_order if c not in final_excludes]
    
    # Appliquer renommages
    renamed_cols = []
    for col in output_cols:
        new_name = final_renames.get(col, col)
        renamed_cols.append(new_name)
    
    return renamed_cols


def validate_transformation(name: str, version: str = "v1") -> Dict[str, Any]:
    """
    Valide qu'une transformation est correctement configurée.
    
    Returns:
        Dict avec {"valid": bool, "errors": List[str], "warnings": List[str]}
    """
    errors = []
    warnings = []
    
    transfo = get_transformation(name, version)
    if not transfo:
        return {"valid": False, "errors": ["Transformation introuvable"], "warnings": []}
    
    # Vérifier gabarit de base
    gabarit_base = transfo.get("gabarit_base", {})
    gab_name = gabarit_base.get("name", "")
    gab_ver = gabarit_base.get("version", "v1")
    
    if not gab_name:
        errors.append("Gabarit de base non défini")
    else:
        gab = get_gabarit(gab_name, gab_ver)
        if not gab:
            errors.append(f"Gabarit de base '{gab_name}' v{gab_ver} introuvable")
    
    # Vérifier enrichissements
    for idx, e in enumerate(transfo.get("enrichments", [])):
        path = e.get("path", [])
        if not path:
            warnings.append(f"Enrichissement #{idx+1} sans chemin")
    
    # Vérifier méthodes
    if gab_name and not errors:
        from backend.services.gabarit_registry import list_methods_for_gabarit
        available_methods = list_methods_for_gabarit(gab_name, gab_ver) or []
        method_names = [m.get("name") for m in available_methods if isinstance(m, dict)] if isinstance(available_methods, list) else list(available_methods.keys())
        
        for method in transfo.get("methods", []):
            if method not in method_names:
                warnings.append(f"Méthode '{method}' introuvable dans le gabarit")
    
    return {
        "valid": len(errors) == 0,
        "errors": errors,
        "warnings": warnings
    }