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

# === Méthodes & dépendances de colonnes (MVP) ================================

from typing import Iterable, Set, Dict, Any, List

def _safe_get(d: Dict[str, Any], *path, default=None):
    cur = d
    for p in path:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(p, {})
    return cur if cur else default

def _index_methods(meta: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """
    Convertit la liste de méthodes d'un gabarit en dict indexé par nom.
    Accepte au choix:
      - {"methods": [{"name":"m1","requires":["A","B"]}, ...]}
      - {"methods": {"m1":{"requires":["A","B"]}, ...}}
    """
    methods = meta.get("methods") or {}
    if isinstance(methods, dict):
        return {k: (v or {}) for k, v in methods.items()}
    if isinstance(methods, list):
        out = {}
        for m in methods:
            if isinstance(m, dict) and m.get("name"):
                out[str(m["name"])] = m
        return out
    return {}

def get_method_requirements(gabarit_name: str, gabarit_version: str, methods: Iterable[str]) -> List[str]:
    """
    Retourne la liste triée des colonnes requises par l'ensemble des 'methods' demandées
    pour un gabarit/version. Tolérant si la méthode n'existe pas dans le registre.
    """
    reg = load_registry()
    g_meta = (reg.get(gabarit_name) or {})
    # on cherche d'abord la version demandée, sinon v1 par défaut
    v_meta = _safe_get(g_meta, "versions", gabarit_version, default=g_meta.get("versions", {}).get("v1", {})) or {}
    idx = _index_methods(v_meta)

    req: Set[str] = set()
    for m in (methods or []):
        m = (m or "").strip()
        if not m:
            continue
        mi = idx.get(m, {})
        cols = mi.get("requires") or []
        for c in cols:
            if isinstance(c, str) and c.strip():
                req.add(c.strip())
    return sorted(req)

# === Index global du registre (accès par nom/version) ========================

def load_registry() -> Dict[str, Dict[str, Any]]:
    """
    Construit un index mémoire des gabarits:
    {
      "<gabarit_name>": {
        "versions": {
          "<version>": { ... métadonnées du gabarit (model_dump) ... }
        }
      },
      ...
    }
    """
    reg: Dict[str, Dict[str, Any]] = {}
    try:
        items = list_gabarits()  # -> List[TableGabarit]
        for g in items:
            try:
                g_name = getattr(g, "name", None) or ""
                g_ver  = getattr(g, "version", None) or "v1"
                if not g_name:
                    continue
                if g_name not in reg:
                    reg[g_name] = {"versions": {}}
                meta = g.model_dump(mode="json") if hasattr(g, "model_dump") else {}
                reg[g_name]["versions"][g_ver] = meta or {}
            except Exception:
                continue
    except Exception as e:
        logger.warning(f"Impossible de charger le registre des gabarits: {e}")
    return reg

# Exports explicites (évitent des surprises avec des imports partiels)
__all__ = [
    "list_gabarits", "get_gabarit", "upsert_gabarit", "delete_gabarit",
    "get_method_requirements", "load_registry"
]
