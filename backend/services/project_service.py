# backend/services/project_service.py
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional, Dict, Any, List, Tuple
from pathlib import Path
import json
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from loguru import logger

# --- dépendances internes (sans casser l’existant)
try:
    from backend.services.template_service import TemplateService
except Exception:
    from services.template_service import TemplateService  # fallback si ton PYTHONPATH diffère

try:
    from backend.services.dataset_service import align_df_to_expected_columns
except Exception:
    def align_df_to_expected_columns(df: pd.DataFrame, expected_columns: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        expected = [c for c in (expected_columns or []) if isinstance(c, str) and c.strip()]
        cur_cols = list(df.columns)
        missing = [c for c in expected if c not in cur_cols]
        for c in missing:
            df[c] = pd.NA
        ordered = expected + [c for c in df.columns if c not in expected]
        return df[ordered], {"missing": missing, "extra": [c for c in cur_cols if c not in expected]}

try:
    # util pour créer dossiers si dispo chez toi
    from backend.utils.file_utils import ensure_directories
except Exception:
    def ensure_directories(*paths):
        for p in paths:
            Path(p).parent.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------

PARIS = ZoneInfo("Europe/Paris")
PROJECTS_DIR = Path("assets") / "projects"
PROJECTS_DIR.mkdir(parents=True, exist_ok=True)

@dataclass
class ProjectId:
    value: str

def _now_paris_iso() -> str:
    return datetime.now(PARIS).isoformat(timespec="seconds")

def _project_path(project_id: str) -> Path:
    return PROJECTS_DIR / f"{project_id}.json"

def _slugify(name: str) -> str:
    s = "".join(ch if ch.isalnum() or ch in "-_" else "-" for ch in name.strip())
    while "--" in s:
        s = s.replace("--", "-")
    return s.strip("-_").lower() or f"project-{datetime.now(PARIS).strftime('%Y%m%d%H%M%S')}"

# ---------------------------------------------------------------------------

class ProjectService:
    """
    Service 'Projet' (MVP JSON).
    Schéma JSON projet :
    {
      "project_id": "mon-projet",
      "name": "Client X",
      "description": "texte",
      "template_ids": [12, 34],
      "parameters": {},

      "gabarit_union": [
        {"gabarit_name":"...", "gabarit_version":"v1", "columns_required":["...","..."]}
      ],

      "gabarit_pipelines": [
        {
          "gabarit_name":"...",
          "gabarit_version":"v1",
          "source": {"type":"csv", "path":"...", "sep":";", "encoding":"utf-8-sig"},
          "sql": "/* optionnel */",
          "python": "# optionnel: df = df.rename(...)\n",
          "last_validation_result": {"errors":0, "warnings":3}
        }
      ],

      "created_at": "2025-10-03T22:10:00+02:00",
      "updated_at": "2025-10-03T22:12:00+02:00"
    }
    """

    def __init__(self, db_session, template_service: Optional[TemplateService] = None):
        self.db = db_session
        self.ts = template_service or TemplateService(db_session)

    # -------------------- CRUD Projet --------------------

    def create_project(self, name: str, description: str = "", project_id: Optional[str] = None,
                       parameters: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        pid = project_id or _slugify(name)
        data = {
            "project_id": pid,
            "name": name,
            "description": description,
            "template_ids": [],
            "parameters": parameters or {},
            "gabarit_union": [],
            "gabarit_pipelines": [],
            "created_at": _now_paris_iso(),
            "updated_at": _now_paris_iso(),
        }
        self.save_project(data)
        logger.success(f"Projet créé: {pid}")
        return data

    def load_project(self, project_id: str) -> Dict[str, Any]:
        path = _project_path(project_id)
        if not path.exists():
            raise FileNotFoundError(f"Projet introuvable: {project_id}")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def save_project(self, project: Dict[str, Any]) -> None:
        path = _project_path(project["project_id"])
        project["updated_at"] = _now_paris_iso()
        ensure_directories(path)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(project, f, ensure_ascii=False, indent=2)
        logger.info(f"Projet sauvegardé: {path}")

    def list_projects(self) -> List[Dict[str, Any]]:
        out = []
        for p in PROJECTS_DIR.glob("*.json"):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    out.append(json.load(f))
            except Exception:
                continue
        return sorted(out, key=lambda x: x.get("updated_at",""), reverse=True)

    # -------------------- Attacher templates --------------------

    def attach_template(self, project_id: str, template_id: int) -> Dict[str, Any]:
        proj = self.load_project(project_id)
        ids = set(proj.get("template_ids", []))
        ids.add(int(template_id))
        proj["template_ids"] = sorted(ids)
        self.save_project(proj)
        return proj

    def detach_template(self, project_id: str, template_id: int) -> Dict[str, Any]:
        proj = self.load_project(project_id)
        ids = [tid for tid in proj.get("template_ids", []) if int(tid) != int(template_id)]
        proj["template_ids"] = ids
        self.save_project(proj)
        return proj

    def list_templates(self, project_id: str) -> List[int]:
        return self.load_project(project_id).get("template_ids", [])

    # -------------------- Union par gabarit --------------------

    def compute_union(self, project_id: str) -> List[Dict[str, Any]]:
        """
        À partir des templates attachés :
        - lit les *tables demandées* (gabarit_usages)
        - calcule les colonnes attendues pour chaque usage (colonnes cochées ∪ dépendances méthodes)
        - regroupe par (gabarit_name, gabarit_version) -> union des colonnes
        """
        proj = self.load_project(project_id)
        template_ids = proj.get("template_ids", [])

        union_map: Dict[Tuple[str,str], List[str]] = {}

        for tid in template_ids:
            usages = self.ts.list_gabarit_usages(int(tid))
            for u in usages:
                gname = u.get("gabarit_name","").strip()
                gver  = u.get("gabarit_version","v1").strip()
                if not gname:
                    continue
                cols = self.ts.resolve_usage_expected_columns(int(tid), gname, gver)
                key = (gname, gver)
                if key not in union_map:
                    union_map[key] = []
                # union ordonnée
                for c in cols:
                    if c not in union_map[key]:
                        union_map[key].append(c)

        union_list = [
            {"gabarit_name": k[0], "gabarit_version": k[1], "columns_required": v}
            for k, v in union_map.items()
        ]

        proj["gabarit_union"] = union_list
        self.save_project(proj)
        logger.success(f"Union colonnes recalculée pour projet '{project_id}' ({len(union_list)} gabarit(s))")
        return union_list

    # -------------------- Pipelines par gabarit --------------------

    def set_pipeline(self, project_id: str, gabarit_name: str, gabarit_version: str,
                     source: Dict[str, Any],
                     sql: Optional[str] = None,
                     python_code: Optional[str] = None) -> Dict[str, Any]:
        """
        Déclare/MAJ le pipeline d'alimentation pour un gabarit {source, sql?, python?}.
        source (MVP): {"type":"csv", "path":"...", "sep":";", "encoding":"utf-8-sig"}
        """
        proj = self.load_project(project_id)
        pipes = proj.get("gabarit_pipelines", [])
        if not isinstance(pipes, list):
            pipes = []

        gname = (gabarit_name or "").strip()
        gver  = (gabarit_version or "v1").strip()

        # supprime l'existant
        pipes = [p for p in pipes if not (p.get("gabarit_name")==gname and p.get("gabarit_version")==gver)]

        item = {
            "gabarit_name": gname,
            "gabarit_version": gver,
            "source": source or {},
            "sql": sql or None,
            "python": python_code or None,
            "last_validation_result": None
        }
        pipes.append(item)

        proj["gabarit_pipelines"] = pipes
        self.save_project(proj)
        return item

    def get_pipeline(self, project_id: str, gabarit_name: str, gabarit_version: str) -> Optional[Dict[str, Any]]:
        proj = self.load_project(project_id)
        for p in proj.get("gabarit_pipelines", []):
            if p.get("gabarit_name")==gabarit_name and p.get("gabarit_version")==gabarit_version:
                return p
        return None

    # -------------------- Preview / Validation (MVP) --------------------

    def _load_source_df(self, source: Dict[str, Any], head: int = 1000) -> pd.DataFrame:
        stype = (source or {}).get("type")
        if stype == "csv":
            path = source.get("path")
            if not path or not Path(path).exists():
                raise FileNotFoundError(f"CSV introuvable: {path}")
            sep = source.get("sep") or ";"
            enc = source.get("encoding") or "utf-8-sig"
            df = pd.read_csv(path, sep=sep, encoding=enc)
            return df.head(head)
        # place-holders pour extensions futures (excel_table/postgres)
        raise NotImplementedError(f"Source non supportée (MVP): {stype}")

    @staticmethod
    def _simplify_dtype(series: pd.Series) -> str:
        dt = str(series.dtype)
        if "datetime" in dt:
            return "date"
        if "int" in dt or "float" in dt or "decimal" in dt:
            return "num"
        return "text"

    def preview(self, project_id: str, gabarit_name: str, gabarit_version: str, head: int = 1000) -> Tuple[pd.DataFrame, Dict[str, Any]]:
        """
        Charge la source, applique SQL/Python si fournis (MVP: SQL ignoré, Python basique),
        retourne df.head(head) + stats de profiling simples.
        """
        pipe = self.get_pipeline(project_id, gabarit_name, gabarit_version)
        if not pipe:
            raise RuntimeError("Pipeline non défini pour ce gabarit")

        # 1) lecture source
        df = self._load_source_df(pipe.get("source"), head=head)

        # 2) (MVP) SQL ignoré (prévu P2: sqlite in-memory)
        # sql = pipe.get("sql")

        # 3) Python (pandas) optionnel – sandbox minimal
        code = pipe.get("python")
        if code and isinstance(code, str) and code.strip():
            loc: Dict[str, Any] = {"df": df, "pd": pd}
            try:
                exec(code, {}, loc)
                if isinstance(loc.get("df"), pd.DataFrame):
                    df = loc["df"]
            except Exception as e:
                logger.warning(f"Erreur script Python pipeline: {e}")

        # Profiling simple
        profile = {
            "rows": int(df.shape[0]),
            "cols": int(df.shape[1]),
            "completeness": {c: float(1 - df[c].isna().mean()) for c in df.columns},
            "dtypes": {c: self._simplify_dtype(df[c]) for c in df.columns}
        }
        return df.head(head), profile

    def validate(self, project_id: str, gabarit_name: str, gabarit_version: str) -> Dict[str, Any]:
        """
        Validation non bloquante :
        - colonnes manquantes par rapport à l'union du projet pour ce gabarit
        - colonnes extra (info)
        - types grossiers (info)
        """
        proj = self.load_project(project_id)
        union = proj.get("gabarit_union", [])
        expected = []
        for u in union:
            if u.get("gabarit_name")==gabarit_name and u.get("gabarit_version")==gabarit_version:
                expected = u.get("columns_required") or []
                break

        pipe = self.get_pipeline(project_id, gabarit_name, gabarit_version)
        if not pipe:
            raise RuntimeError("Pipeline non défini pour ce gabarit")

        df, profile = self.preview(project_id, gabarit_name, gabarit_version, head=1000)

        # alignement (non bloquant)
        _, warnings = align_df_to_expected_columns(df.copy(), expected)

        result = {
            "errors": 0,  # MVP: jamais bloquant
            "warnings": len(warnings.get("missing", [])),
            "missing_columns": warnings.get("missing", []),
            "extra_columns": warnings.get("extra", []),
            "profile": profile
        }

        # persist
        proj = self.load_project(project_id)
        pipes = proj.get("gabarit_pipelines", [])
        for p in pipes:
            if p.get("gabarit_name")==gabarit_name and p.get("gabarit_version")==gabarit_version:
                p["last_validation_result"] = result
                break
        proj["gabarit_pipelines"] = pipes
        self.save_project(proj)

        return result

    # === Lecture FULL & construction DataFrame pour injection =====================

    def _load_source_df_full(self, source: Dict[str, Any]) -> pd.DataFrame:
        stype = (source or {}).get("type")
        if stype == "csv":
            path = source.get("path")
            if not path or not Path(path).exists():
                raise FileNotFoundError(f"CSV introuvable: {path}")
            sep = source.get("sep") or ";"
            enc = source.get("encoding") or "utf-8-sig"
            return pd.read_csv(path, sep=sep, encoding=enc)
        raise NotImplementedError(f"Source non supportée (MVP): {stype}")

    def build_dataframe(self, project_id: str, gabarit_name: str, gabarit_version: str) -> pd.DataFrame:
        """
        Construit le DataFrame complet pour injection (pas de limite de lignes),
        en appliquant le pipeline Python si présent.
        """
        pipe = self.get_pipeline(project_id, gabarit_name, gabarit_version)
        if not pipe:
            raise RuntimeError("Pipeline non défini pour ce gabarit")

        df = self._load_source_df_full(pipe.get("source"))

        code = pipe.get("python")
        if code and isinstance(code, str) and code.strip():
            loc: Dict[str, Any] = {"df": df, "pd": pd}
            try:
                exec(code, {}, loc)
                if isinstance(loc.get("df"), pd.DataFrame):
                    df = loc["df"]
            except Exception as e:
                logger.warning(f"Erreur script Python pipeline (full): {e}")

        return df
