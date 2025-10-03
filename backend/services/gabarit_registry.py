# backend/services/gabarit_registry.py
from __future__ import annotations
from pathlib import Path
import json
from typing import List, Optional, Dict, Any
from loguru import logger

from backend.models.gabarits import TableGabarit

_REG_DIR = Path("assets/registry")
_REG_FILE = _REG_DIR / "gabarits.json"

def _ensure_storage() -> None:
    _REG_DIR.mkdir(parents=True, exist_ok=True)
    if not _REG_FILE.exists():
        _REG_FILE.write_text(json.dumps({"gabarits": []}, ensure_ascii=False, indent=2), encoding="utf-8")

def _load_raw() -> Dict[str, Any]:
    _ensure_storage()
    return json.loads(_REG_FILE.read_text(encoding="utf-8"))

def _save_raw(data: Dict[str, Any]) -> None:
    _ensure_storage()
    _REG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def list_gabarits() -> List[TableGabarit]:
    data = _load_raw()
    return [TableGabarit(**g) for g in data.get("gabarits", [])]

def get_gabarit(name: str, version: str = "v1") -> Optional[TableGabarit]:
    for g in list_gabarits():
        if g.name == name and g.version == version:
            return g
    return None

def upsert_gabarit(gabarit: TableGabarit) -> None:
    data = _load_raw()
    items = data.get("gabarits", [])
    items = [g for g in items if not (g.get("name") == gabarit.name and g.get("version") == gabarit.version)]
    items.append(gabarit.model_dump(mode="json"))
    data["gabarits"] = items
    _save_raw(data)
    logger.info(f"Gabarit upsert: {gabarit.name} v{gabarit.version}")

def delete_gabarit(name: str, version: str = "v1") -> bool:
    data = _load_raw()
    items = data.get("gabarits", [])
    new_items = [g for g in items if not (g.get("name") == name and g.get("version") == version)]
    if len(new_items) == len(items):
        return False
    data["gabarits"] = new_items
    _save_raw(data)
    logger.info(f"Gabarit supprimé: {name} v{version}")
    return True
