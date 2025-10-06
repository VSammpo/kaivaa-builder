# backend/services/method_executor.py
import pandas as pd
from typing import Any

class MethodExecutionError(Exception):
    pass

def _build_params(param_schema: list[dict], param_values: dict[str, Any], df: pd.DataFrame) -> dict:
    """
    Valide/normalise les paramètres côté Python, en appliquant les defaults.
    Gère types: text, number, boolean, select, select_from_column.
    """
    out = {}
    schema_idx = {p["name"]: p for p in (param_schema or []) if p.get("name")}
    for name, spec in schema_idx.items():
        t = spec.get("type", "text")
        default = spec.get("default", None)
        val = param_values.get(name, default)

        if t == "number" and val is not None:
            try:
                val = float(val)
            except Exception:
                raise MethodExecutionError(f"Paramètre '{name}' doit être un nombre.")

        if t == "boolean":
            val = True if str(val).lower() in ("1", "true", "yes", "y", "on") else False

        if t == "select":
            options = spec.get("options") or []
            if val is None and options:
                val = options[0]
            if options and val not in options:
                raise MethodExecutionError(f"Paramètre '{name}' doit appartenir à {options}.")

        if t == "select_from_column":
            col = spec.get("source_column")
            if not col or col not in df.columns:
                raise MethodExecutionError(f"Paramètre '{name}': colonne source invalide '{col}'.")

        out[name] = val
    return out

def apply_method(df: pd.DataFrame, method: dict, params: dict[str, Any]) -> pd.DataFrame:
    """
    Exécute une méthode 'colonne calculée':
      - method['output_column'] = nom de la colonne
      - method['code'] doit définir une variable 'value' (Series/array/scalar)
      - affectation: df[output_column] = value
    """
    code = (method or {}).get("code", "")
    required_cols = (method or {}).get("required_columns", []) or []
    param_schema = (method or {}).get("param_schema", []) or []
    out_col = (method or {}).get("output_column", "") or ""

    if not out_col:
        raise MethodExecutionError("La méthode n'a pas de 'output_column' défini.")

    for c in required_cols:
        if c not in df.columns:
            raise MethodExecutionError(f"Colonne requise manquante: '{c}'.")

    # Paramètres normalisés
    norm_params = _build_params(param_schema, params or {}, df)

    df2 = df.copy()
    # Exec sandbox: on attend que 'value' apparaisse
    loc = {"df": df2, "pd": pd, "params": norm_params, "value": None}
    try:
        exec(code, {}, loc)
    except Exception as e:
        raise MethodExecutionError(f"Erreur exécution méthode: {e}")

    value = loc.get("value", None)
    if value is None:
        raise MethodExecutionError("Le code de la méthode doit définir une variable 'value' (Series/array/scalar).")

    # Assignation
    try:
        df2[out_col] = value
    except Exception as e:
        raise MethodExecutionError(f"Impossible d'assigner la valeur à la colonne '{out_col}': {e}")

    return df2
