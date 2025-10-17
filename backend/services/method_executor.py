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
    Exécute une méthode 'colonne calculée'. On supporte trois styles de code utilisateur :
      A) Affectation directe : df[out] = ...
      B) Variable 'value'   : value = ...
      C) Expression seule   : (ex: df['x'] * 3)  → df[out] = <résultat>
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

    # --- 1) Essayer d'évaluer comme une simple expression (mode C)
    #     Si le code est une expression valide, on affecte son résultat à df[out_col].
    try:
        compiled = compile(code, "<method>", "eval")
    except SyntaxError:
        compiled = None

    if compiled is not None:
        try:
            env = {"df": df2, "pd": pd, "params": norm_params, "__builtins__": __builtins__}
            result = eval(compiled, env, env)

            df2[out_col] = result
            return df2
        except Exception:
            # on retombe sur le mode 'exec' (A/B)
            pass

    # --- 2) Exécuter comme script (modes A et B)
    env = {"df": df2, "pd": pd, "params": norm_params, "value": None, "__builtins__": __builtins__}
    exec(code, env, env)
    loc = env


    # 2a) Style B : 'value' fourni → on affecte df[out_col] = value
    value = loc.get("value", None)
    if value is not None:
        try:
            df2[out_col] = value
            return df2
        except Exception as e:
            raise MethodExecutionError(f"Impossible d'assigner la valeur à '{out_col}': {e}")

    # 2b) Style A : l'utilisateur a déjà écrit df[out_col] = ... dans le code
    if out_col in df2.columns:
        return df2

    # Rien n'a produit la colonne
    raise MethodExecutionError(
        "Le code doit soit écrire df[out] = ..., soit définir 'value', soit être une expression."
    )
