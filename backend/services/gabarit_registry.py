# backend/services/gabarit_registry.py
from __future__ import annotations
from pathlib import Path
import json
from typing import List, Optional, Dict, Any
from loguru import logger
import pandas as pd

from backend.models.gabarits import TableGabarit

# --- NOUVEAU STOCKAGE FILE-BASED ---
CONF_DIR = Path("configuration")
GAB_DIR = CONF_DIR / "gabarits"

def _gab_path(name: str, version: str) -> Path:
    return GAB_DIR / name / f"{(version or 'v1')}.json"

def _ensure_dirs() -> None:
    GAB_DIR.mkdir(parents=True, exist_ok=True)

def _read_json(p: Path) -> dict:
    return json.loads(p.read_text(encoding="utf-8"))

def _write_json(p: Path, data: dict) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)

def _normalize_gabarit_payload(d: dict) -> dict:
    d = dict(d or {})
    d.setdefault("name", "")
    d.setdefault("version", "v1")
    d.setdefault("title", d.get("name", ""))
    d.setdefault("description", "")
    d.setdefault("columns", [])
    d.setdefault("methods", [])        # méthodes propres au gabarit
    d.setdefault("defaults", {})       # {source, preview}
    d.setdefault("relations", [])      # relations sortantes
    d.setdefault("role", None)         # "fact"/"dimension"/"mixed"/None
    return d


def set_default_source(gabarit_name: str, gabarit_version: str, source: dict | None) -> None:
    p = _gab_path(gabarit_name, gabarit_version)
    data = _normalize_gabarit_payload(_read_json(p) if p.exists() else {"name": gabarit_name, "version": gabarit_version})
    if source is None:
        data.setdefault("defaults", {}).pop("source", None)
        data.setdefault("defaults", {}).pop("preview", None)  # on retire aussi l'aperçu si on retire la source
    else:
        d = data.setdefault("defaults", {})
        d["source"] = source
    _write_json(p, data)

def set_default_preview(gabarit_name: str, gabarit_version: str, rows: list[dict], columns: list[str]) -> None:
    p = _gab_path(gabarit_name, gabarit_version)
    data = _normalize_gabarit_payload(_read_json(p) if p.exists() else {"name": gabarit_name, "version": gabarit_version})
    d = data.setdefault("defaults", {})
    d["preview"] = {"columns": list(columns or []), "rows": list(rows or [])[:20]}
    _write_json(p, data)

def get_default_preview(gabarit_name: str, gabarit_version: str) -> dict | None:
    """
    Lit directement le JSON du gabarit et renvoie defaults.preview.
    """
    p = _gab_path(gabarit_name, gabarit_version)
    if not p.exists():
        return None
    try:
        data = _read_json(p)
        d = data.get("defaults") or {}
        return d.get("preview")
    except Exception:
        return None


def get_default_source(gabarit_name: str, gabarit_version: str) -> dict | None:
    """
    Lit directement le JSON du gabarit et renvoie defaults.source
    (on n'utilise pas TableGabarit pour éviter les attributs absents).
    """
    p = _gab_path(gabarit_name, gabarit_version)
    if not p.exists():
        return None
    try:
        data = _read_json(p)
        d = data.get("defaults") or {}
        return d.get("source")
    except Exception:
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
    """
    Version file-based : compte les relations 'from' et 'to' en scannant tous les fichiers
    configuration/gabarits/<Nom>/<version>.json.
    """
    _ensure_dirs()
    v = (gabarit_version or "v1").strip()
    total = 0
    if not GAB_DIR.exists():
        return 0
    for gdir in GAB_DIR.iterdir():
        if not gdir.is_dir():
            continue
        for jf in gdir.glob("*.json"):
            try:
                data = _normalize_gabarit_payload(_read_json(jf))
                for r in data.get("relations", []) or []:
                    if (
                        (r.get("from_gabarit") == gabarit_name and (r.get("from_version") or "v1") == v)
                        or
                        (r.get("to_gabarit") == gabarit_name and (r.get("to_version") or "v1") == v)
                    ):
                        total += 1
            except Exception:
                continue
    return total

def soft_delete_gabarit(gabarit_name: str, gabarit_version: str) -> dict:
    """
    Archive le gabarit sous configuration/gabarits/_trash/<Nom_Supr_nXXXX>/<version>.json
    et supprime toutes les relations (in/out) qui le mentionnent. Retourne un dict:
    { old_name, new_name, removed_relations } où removed_relations est un int.
    """
    v = (gabarit_version or "v1").strip()
    p = _gab_path(gabarit_name, v)
    if not p.exists():
        raise FileNotFoundError("Gabarit introuvable")

    # charger le gabarit
    data = _normalize_gabarit_payload(_read_json(p))
    base = gabarit_name

    # calcul du suffixe nXXXX
    trash_root = GAB_DIR / "_trash"
    trash_root.mkdir(parents=True, exist_ok=True)
    prefix = f"{base}_Supr_n"
    # chercher le prochain index
    k = 1
    existing = [d.name for d in trash_root.iterdir() if d.is_dir() and d.name.startswith(prefix)]
    if existing:
        try:
            k = max(int(x.split(prefix, 1)[-1]) for x in existing) + 1
        except Exception:
            k = len(existing) + 1
    new_name = f"{base}_Supr_n{str(k).zfill(4)}"

    # MAJ du nom dans le JSON
    data["name"] = new_name

    # écrire dans la corbeille
    dest = trash_root / new_name / f"{v}.json"
    dest.parent.mkdir(parents=True, exist_ok=True)
    _write_json(dest, data)

    # supprimer l'ancien fichier + dossier si vide
    try:
        p.unlink()
        try:
            p.parent.rmdir()
        except Exception:
            pass
    except Exception:
        pass

    # Purger les relations (in/out) dans TOUS les gabarits actifs (hors _trash)
    removed = 0
    if GAB_DIR.exists():
        for gdir in GAB_DIR.iterdir():
            if not gdir.is_dir():
                continue
            if gdir.name.startswith("_"):  # ignore corbeille
                continue
            for jf in gdir.glob("*.json"):
                try:
                    d = _normalize_gabarit_payload(_read_json(jf))
                    rels = d.get("relations", []) or []
                    before = len(rels)
                    rels = [r for r in rels if not (
                        (r.get("from_gabarit") == gabarit_name and (r.get("from_version") or "v1") == v)
                        or
                        (r.get("to_gabarit") == gabarit_name and (r.get("to_version") or "v1") == v)
                    )]
                    if len(rels) != before:
                        removed += (before - len(rels))
                        d["relations"] = rels
                        _write_json(jf, d)
                except Exception:
                    continue

    return {"old_name": gabarit_name, "new_name": new_name, "removed_relations": removed}

def list_gabarits() -> List[TableGabarit]:
    _ensure_dirs()
    out: List[TableGabarit] = []
    if not GAB_DIR.exists():
        return out
    for gdir in sorted([p for p in GAB_DIR.iterdir() if p.is_dir()]):
        if gdir.name.startswith("_"):      # ignore _trash, _index, etc.
            continue
        for jf in sorted(gdir.glob("*.json")):
            try:
                data = _normalize_gabarit_payload(_read_json(jf))
                out.append(TableGabarit(**data))
            except Exception:
                continue
    return out


def get_gabarit(name: str, version: str = "v1") -> Optional[TableGabarit]:
    p = _gab_path((name or "").strip(), (version or "v1").strip())
    if not p.exists():
        return None
    try:
        data = _normalize_gabarit_payload(_read_json(p))
        return TableGabarit(**data)
    except Exception:
        return None

def upsert_gabarit(gabarit: TableGabarit) -> None:
    _ensure_dirs()
    payload = _normalize_gabarit_payload(gabarit.model_dump(mode="json"))
    p = _gab_path(payload["name"], payload["version"])
    _write_json(p, payload)
    logger.info(f"Gabarit upsert: {payload['name']} {payload['version']} -> {p}")

def delete_gabarit(name: str, version: str = "v1") -> bool:
    p = _gab_path((name or "").strip(), (version or "v1").strip())
    if not p.exists():
        return False
    p.unlink()
    # si le dossier est vide, on le supprime
    try:
        p.parent.rmdir()
    except Exception:
        pass
    logger.info(f"Gabarit supprimé: {name} {version}")
    return True

# === Méthodes & dépendances de colonnes (MVP) ================================

from typing import Iterable, Set

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
    role = (role or "").strip().lower()
    assert role in {"fact", "dimension", "mixed"}, "role invalide"
    p = _gab_path(gabarit_name, gabarit_version)
    data = _normalize_gabarit_payload(_read_json(p) if p.exists() else {"name": gabarit_name, "version": gabarit_version})
    data["role"] = role
    _write_json(p, data)

def get_role(gabarit_name: str, gabarit_version: str) -> str | None:
    """
    Lit directement le JSON du gabarit et renvoie le champ 'role' (si présent).
    On ne passe pas par TableGabarit pour éviter les attributs absents du modèle.
    """
    p = _gab_path(gabarit_name, gabarit_version)
    if not p.exists():
        return None
    try:
        data = _read_json(p)
        role = (data or {}).get("role")
        return (role or None)
    except Exception:
        return None

def add_relation(from_gabarit: str, from_version: str,
                 to_gabarit: str, to_version: str,
                 left_key: str, right_key: str) -> dict:
    p = _gab_path(from_gabarit, from_version)
    data = _normalize_gabarit_payload(_read_json(p) if p.exists() else {"name": from_gabarit, "version": from_version})
    rels = data.setdefault("relations", [])
    item = {
        "from_gabarit": (from_gabarit or "").strip(),
        "from_version": (from_version or "v1").strip(),
        "to_gabarit": (to_gabarit or "").strip(),
        "to_version": (to_version or "v1").strip(),
        "left_key": (left_key or "").strip(),
        "right_key": (right_key or "").strip(),
    }
    # id stable pour éviter les doublons exacts
    item["relation_id"] = (
        f"{item['from_gabarit']}|{item['from_version']}->"
        f"{item['to_gabarit']}|{item['to_version']}::"
        f"{item['left_key']}={item['right_key']}"
    )
    if not any(r.get("relation_id") == item["relation_id"] for r in rels if isinstance(r, dict)):
        rels.append(item)
        _write_json(p, data)
    return item

def delete_relation(from_gabarit: str, from_version: str,
                    to_gabarit: str, to_version: str,
                    left_key: str, right_key: str) -> bool:
    p = _gab_path(from_gabarit, from_version)
    if not p.exists():
        return False
    data = _read_json(p)
    rels = data.get("relations") or []
    before = len(rels)
    rels = [
        r for r in rels if not (
            (r.get("from_gabarit")==from_gabarit and (r.get("from_version") or "v1")==from_version) and
            (r.get("to_gabarit")==to_gabarit and (r.get("to_version") or "v1")==to_version) and
            (r.get("left_key")==left_key and r.get("right_key")==right_key)
        )
    ]
    if len(rels) != before:
        data["relations"] = rels
        _write_json(p, data)
        return True
    return False

def get_relations(gabarit_name: str, gabarit_version: str) -> list[dict]:
    """
    Lit directement le JSON du gabarit et renvoie la liste 'relations'.
    Évite de passer par TableGabarit (qui ne porte pas ce champ).
    """
    p = _gab_path(gabarit_name, gabarit_version)
    if not p.exists():
        return []
    try:
        data = _read_json(p)
        rels = data.get("relations") or []
        # on garantit une liste de dicts
        return [dict(r) for r in rels if isinstance(r, dict)]
    except Exception:
        return []


# ====== MÉTHODES PAR GABARIT =================================================
# Stockage dans le registre JSON sous clé:
# "gabarit_methods": {
#    "<name>|<version>": [
#       { "id", "name", "description", "output_column",
#         "param_schema": [...], "required_columns": [...],
#         "code", "order": int }
#    ],
#    ...
# }

import uuid

def _ensure_gab_methods_key(data: dict) -> dict:
    if "gabarit_methods" not in data:
        data["gabarit_methods"] = {}
    return data

def _gab_key(gabarit_name: str, gabarit_version: str) -> str:
    return f"{gabarit_name}|{gabarit_version or 'v1'}"


def list_methods_for_gabarit(gabarit_name: str, gabarit_version: str) -> list[dict]:
    """
    Lit directement le fichier configuration/gabarits/<name>/<version>.json
    pour récupérer la liste 'methods'. On évite TableGabarit (qui ne porte pas ce champ).
    """
    p = _gab_path(gabarit_name, gabarit_version)
    if not p.exists():
        return []
    try:
        data = _read_json(p)
        arr = data.get("methods") or []
        arr = [dict(m) for m in arr if isinstance(m, dict)]
        return sorted(arr, key=lambda m: m.get("order", 1))
    except Exception:
        return []


def get_method_for_gabarit(gabarit_name: str, gabarit_version: str, method_id: str) -> dict | None:
    methods = list_methods_for_gabarit(gabarit_name, gabarit_version)
    for m in methods:
        if m.get("id") == method_id:
            return m
    return None


def upsert_method_for_gabarit(
    gabarit_name: str,
    gabarit_version: str,
    *,
    name: str,
    description: str,
    output_column: str,
    param_schema: list[dict],
    required_columns: list[str],
    code: str,
    order: int | None = None,
    method_id: str | None = None,
) -> dict:
    p = _gab_path(gabarit_name, gabarit_version)
    data = _normalize_gabarit_payload(_read_json(p) if p.exists() else {"name": gabarit_name, "version": gabarit_version})
    ms = list(data.get("methods") or [])

    if method_id:
        # update
        for i, m in enumerate(ms):
            if m.get("id") == method_id:
                ms[i] = {
                    "id": method_id,
                    "name": name,
                    "description": description,
                    "output_column": (output_column or "").strip(),
                    "param_schema": param_schema or [],
                    "required_columns": required_columns or [],
                    "code": code or "",
                    "order": int(order if order is not None else m.get("order", 1)),
                }
                break
        else:
            raise ValueError(f"Method not found in gabarit: {method_id}")
        data["methods"] = ms
        _write_json(p, data)
        return ms[i]
    else:
        import uuid
        mid = str(uuid.uuid4())
        new_order = int(order) if order is not None else (ms[-1].get("order", 0) + 1 if ms else 1)
        entry = {
            "id": mid,
            "name": name,
            "description": description,
            "output_column": (output_column or "").strip(),
            "param_schema": param_schema or [],
            "required_columns": required_columns or [],
            "code": code or "",
            "order": new_order,
        }
        ms.append(entry)
        data["methods"] = ms
        _write_json(p, data)
        return entry

def delete_method_for_gabarit(gabarit_name: str, gabarit_version: str, method_id: str) -> None:
    p = _gab_path(gabarit_name, gabarit_version)
    if not p.exists():
        return
    data = _normalize_gabarit_payload(_read_json(p))
    data["methods"] = [m for m in (data.get("methods") or []) if m.get("id") != method_id]
    _write_json(p, data)

def reorder_methods_for_gabarit(gabarit_name: str, gabarit_version: str, ordered_ids: list[str]) -> None:
    p = _gab_path(gabarit_name, gabarit_version)
    if not p.exists():
        return
    data = _normalize_gabarit_payload(_read_json(p))
    arr = list(data.get("methods") or [])
    idx = {m["id"]: m for m in arr}
    new_arr = []
    for pos, mid in enumerate(ordered_ids, start=1):
        if mid in idx:
            m = idx[mid]
            m["order"] = pos
            new_arr.append(m)
    for m in arr:
        if m["id"] not in ordered_ids:
            m["order"] = len(new_arr) + 1
            new_arr.append(m)
    data["methods"] = new_arr
    _write_json(p, data)


# --- AJOUT : façade full dataframe --------------------------------------------


def get_default_dataframe(gabarit_name: str, gabarit_version: str = "v1") -> Optional[pd.DataFrame]:
    """
    Charge la donnée par défaut complète d'un gabarit (si une 'source' est définie).
    Retourne un DataFrame ou None si absent/erreur.
    """
    try:
        src = get_default_source(gabarit_name, gabarit_version)
        if not src:
            return None
        from backend.services.dataset_service import _load_dataframe_from_source
        return _load_dataframe_from_source(src)
    except Exception:
        return None
