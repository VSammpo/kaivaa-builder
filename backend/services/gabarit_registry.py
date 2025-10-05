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

def _normalize(data: dict) -> dict:
    data.setdefault("gabarits", [])
    data.setdefault("roles", [])
    data.setdefault("relations", [])
    data.setdefault("defaults", [])
    data.setdefault("trash", {"gabarits": []})
    return data

def _load_raw() -> dict:
    _ensure_storage()
    data = json.loads(_REG_FILE.read_text(encoding="utf-8"))
    return _normalize(data)

def _save_raw(data: dict) -> None:
    _ensure_storage()
    data = _normalize(data)
    _REG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def set_default_source(gabarit_name: str, gabarit_version: str, source: dict | None) -> None:
    """
    Enregistre (ou retire) la source par défaut pour un gabarit.
    Si source is None -> supprime aussi l'éventuel preview mémorisé.
    """
    data = _load_raw()
    defs = data.get("defaults", [])
    key = (gabarit_name, gabarit_version or "v1")

    # Filtrer l'entrée existante
    new_defs = []
    existing_preview = None
    for d in defs:
        if d.get("gabarit_name") == key[0] and (d.get("gabarit_version") or "v1") == key[1]:
            existing_preview = d.get("preview")  # on le garde si on réécrit la source
            continue
        new_defs.append(d)

    # Si on retire la source -> on retire aussi le preview
    if source is None:
        data["defaults"] = new_defs
        _save_raw(data)
        return

    # On réinsère l'entrée, en conservant le preview existant si présent
    entry = {
        "gabarit_name": key[0],
        "gabarit_version": key[1],
        "source": source,
    }
    if existing_preview:
        entry["preview"] = existing_preview

    new_defs.append(entry)
    data["defaults"] = new_defs
    _save_raw(data)

def set_default_preview(gabarit_name: str, gabarit_version: str, rows: list[dict], columns: list[str]) -> None:
    """
    Mémorise un mini-apercu (rows max 20) pour la donnée par défaut du gabarit.
    Crée l'entrée si elle n'existe pas encore (avec source vide).
    """
    data = _load_raw()
    defs = data.get("defaults", [])
    v = (gabarit_version or "v1")

    found = False
    for d in defs:
        if d.get("gabarit_name") == gabarit_name and (d.get("gabarit_version") or "v1") == v:
            d["preview"] = {
                "columns": list(columns or []),
                "rows": list(rows or [])[:20],
            }
            found = True
            break

    if not found:
        defs.append({
            "gabarit_name": gabarit_name,
            "gabarit_version": v,
            "source": {},
            "preview": {
                "columns": list(columns or []),
                "rows": list(rows or [])[:20],
            }
        })

    data["defaults"] = defs
    _save_raw(data)


def get_default_preview(gabarit_name: str, gabarit_version: str) -> dict | None:
    """
    Retourne un dict {"columns": [...], "rows": [...]} ou None si absent.
    """
    data = _load_raw()
    v = (gabarit_version or "v1")
    for d in data.get("defaults", []):
        if d.get("gabarit_name") == gabarit_name and (d.get("gabarit_version") or "v1") == v:
            return d.get("preview") or None
    return None


def get_default_source(gabarit_name: str, gabarit_version: str) -> dict | None:
    data = _load_raw()
    for d in data.get("defaults", []):
        if d.get("gabarit_name")==gabarit_name and (d.get("gabarit_version") or "v1")==(gabarit_version or "v1"):
            return d.get("source") or None
    return None

def clear_default_source(gabarit_name: str, gabarit_version: str) -> None:
    set_default_source(gabarit_name, gabarit_version, None)


def _next_supr_suffix(existing_names: list[str], base: str) -> str:
    # renvoie "Nom_Supr_n0001" (ou n0002, etc.)
    base_prefix = f"{base}_Supr_n"
    nums = []
    for n in existing_names:
        if n.startswith(base_prefix):
            try:
                nums.append(int(n.split(base_prefix, 1)[1]))
            except Exception:
                pass
    k = (max(nums) + 1) if nums else 1
    return f"{base}_Supr_n{str(k).zfill(4)}"

def count_links(gabarit_name: str, gabarit_version: str) -> int:
    data = _load_raw()
    v = (gabarit_version or "v1").strip()
    rels = data.get("relations", [])
    n = 0
    for r in rels:
        if (r.get("from_gabarit")==gabarit_name and (r.get("from_version") or "v1")==v) or \
           (r.get("to_gabarit")==gabarit_name and (r.get("to_version") or "v1")==v):
            n += 1
    return n

def soft_delete_gabarit(gabarit_name: str, gabarit_version: str) -> dict:
    """
    Supprime 'logiquement' un gabarit:
      - retire le gabarit actif
      - enlève son rôle
      - supprime toutes les relations (from/to) qui le mentionnent
      - déplace une copie dans trash.gabarits avec un nom renommé 'Nom_Supr_nXXXX'
    Retourne {"old_name":..., "new_name":..., "removed_relations": N}
    """
    data = _load_raw()
    v = (gabarit_version or "v1").strip()

    # 1) récupérer l'objet gabarit
    gabs = data.get("gabarits", [])
    idx = None
    for i, g in enumerate(gabs):
        if g.get("name")==gabarit_name and (g.get("version") or "v1")==v:
            idx = i
            break
    if idx is None:
        raise FileNotFoundError("Gabarit introuvable")

    gab = gabs.pop(idx)  # retirer de la liste active

    # 2) retirer le rôle
    roles = [r for r in data.get("roles", []) if not (r.get("gabarit_name")==gabarit_name and (r.get("gabarit_version") or "v1")==v)]
    data["roles"] = roles

    # 3) retirer toutes les relations (from/to) qui le mentionnent
    rels = data.get("relations", [])
    before = len(rels)
    rels = [r for r in rels if not (
        (r.get("from_gabarit")==gabarit_name and (r.get("from_version") or "v1")==v) or
        (r.get("to_gabarit")==gabarit_name and (r.get("to_version") or "v1")==v)
    )]
    removed = before - len(rels)
    data["relations"] = rels

    # 4) renommer et pousser dans la trash
    trash_list = data.setdefault("trash", {}).setdefault("gabarits", [])
    active_names = [x.get("name") for x in trash_list if x.get("name", "").startswith(f"{gabarit_name}_Supr_n")]
    new_name = _next_supr_suffix(active_names, gabarit_name)

    gab["name"] = new_name
    trash_list.append(gab)

    data["gabarits"] = gabs
    _save_raw(data)

    return {"old_name": gabarit_name, "new_name": new_name, "removed_relations": removed}

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

def set_role(gabarit_name: str, gabarit_version: str, role: str) -> None:
    """
    role ∈ {'fact','dimension','mixed'} – stocké dans le registre (clé: name+version)
    """
    role = (role or "").strip().lower()
    assert role in {"fact", "dimension", "mixed"}, "role invalide"

    data = _load_raw()
    roles = data.get("roles", [])
    # on remplace l'existant pour (name, version)
    roles = [r for r in roles
             if not (r.get("gabarit_name") == gabarit_name and (r.get("gabarit_version") or "v1") == (gabarit_version or "v1"))]
    roles.append({
        "gabarit_name": gabarit_name,
        "gabarit_version": gabarit_version or "v1",
        "role": role,
    })
    data["roles"] = roles
    _save_raw(data)


def get_role(gabarit_name: str, gabarit_version: str) -> str | None:
    data = _load_raw()
    for r in data.get("roles", []):
        if r.get("gabarit_name") == gabarit_name and (r.get("gabarit_version") or "v1") == (gabarit_version or "v1"):
            return r.get("role")
    return None


def add_relation(from_gabarit: str, from_version: str,
                 to_gabarit: str, to_version: str,
                 left_key: str, right_key: str) -> dict:
    """
    Enregistre une relation (clé↔clé) au niveau 'catalogue'.
    Retourne l'item créé, avec 'relation_id' stable.
    """
    item = {
        "from_gabarit": (from_gabarit or "").strip(),
        "from_version": (from_version or "v1").strip(),
        "to_gabarit": (to_gabarit or "").strip(),
        "to_version": (to_version or "v1").strip(),
        "left_key": (left_key or "").strip(),
        "right_key": (right_key or "").strip(),
    }
    # id stable
    item["relation_id"] = (
        f"{item['from_gabarit']}|{item['from_version']}->"
        f"{item['to_gabarit']}|{item['to_version']}::"
        f"{item['left_key']}={item['right_key']}"
    )

    data = _load_raw()
    rels = data.get("relations", [])

    # anti-dup EXACT
    if not any(r == item for r in rels):
        rels.append(item)
        data["relations"] = rels
        _save_raw(data)
    return item


def delete_relation(from_gabarit: str, from_version: str,
                    to_gabarit: str, to_version: str,
                    left_key: str, right_key: str) -> bool:
    data = _load_raw()
    rels = data.get("relations", [])
    before = len(rels)
    rels = [r for r in rels if not (
        r.get("from_gabarit") == (from_gabarit or "").strip()
        and (r.get("from_version") or "v1") == (from_version or "v1").strip()
        and r.get("to_gabarit") == (to_gabarit or "").strip()
        and (r.get("to_version") or "v1") == (to_version or "v1").strip()
        and r.get("left_key") == (left_key or "").strip()
        and r.get("right_key") == (right_key or "").strip()
    )]
    changed = len(rels) != before
    if changed:
        data["relations"] = rels
        _save_raw(data)
    return changed


def get_relations(gabarit_name: str, gabarit_version: str) -> list[dict]:
    """
    Relations SORTANTES (FROM = ce gabarit).
    """
    data = _load_raw()
    rels = data.get("relations", [])
    return [
        r for r in rels
        if r.get("from_gabarit") == (gabarit_name or "").strip()
        and (r.get("from_version") or "v1") == (gabarit_version or "v1").strip()
    ]
