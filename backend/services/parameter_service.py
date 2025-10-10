# backend/services/parameter_service.py
"""
Service de gestion des paramètres de templates
"""

from typing import List, Dict, Optional, Any
import pandas as pd
from loguru import logger

from backend.models.template_config import ParameterConfig


class ParameterService:
    """Service pour résoudre et valider les paramètres de templates"""
    
    @staticmethod
    def resolve_parameter_options(param: ParameterConfig) -> List[str]:
        """
        Résout les options disponibles pour un paramètre.
        PRIORITÉ au cache 'options_cache' s'il est présent (zéro recalcul).
        Sinon, fallback sur le comportement existant.
        """
        cache = getattr(param, "options_cache", None)
        if isinstance(cache, dict):
            vals = cache.get("values")
            if isinstance(vals, list) and vals:
                return vals

        mode = getattr(param, "options_mode", "none")
        if mode == "none":
            return []
        elif mode == "manual":
            return param.options_manual or []
        elif mode == "from_column":
            # NOTE: calcul potentiellement lourd -> ne sera appelé que si pas de cache
            return ParameterService._resolve_options_from_gabarit(param)

        return []


    @staticmethod
    def _resolve_options_from_gabarit(param: ParameterConfig) -> List[str]:
        """
        Charge les valeurs uniques depuis une colonne de gabarit.
        TOUJOURS sur le FULL dataframe, jamais sur preview.
        
        Args:
            param: Configuration du paramètre
            
        Returns:
            Liste des valeurs uniques (limitée à 500)
        """
        try:
            from backend.services.gabarit_registry import get_default_dataframe
            
            source = param.options_source or {}
            gabarit_name = source.get("gabarit", "")
            gabarit_version = source.get("version", "v1")
            column_name = source.get("column", "")
            
            if not gabarit_name or not column_name:
                logger.warning(f"Source incomplète pour paramètre '{param.name}'")
                return []
            
            # ✅ TOUJOURS charger le FULL (pas de preview ici)
            try:
                df_full = get_default_dataframe(gabarit_name, gabarit_version)
                if isinstance(df_full, pd.DataFrame) and column_name in df_full.columns:
                    values = sorted(df_full[column_name].dropna().astype(str).unique().tolist())
                    logger.debug(f"Options pour '{param.name}' depuis FULL: {len(values)} valeurs")
                    return values
            except Exception as e:
                logger.warning(f"Impossible de charger données FULL pour '{gabarit_name}': {e}")
            
            return []
        
        except Exception as e:
            logger.error(f"Erreur résolution options pour '{param.name}': {e}")
            return []

    @staticmethod
    def validate_parameter_value(param: ParameterConfig, value: Any) -> tuple[bool, Optional[str]]:
        """
        Valide une valeur de paramètre.
        
        Args:
            param: Configuration du paramètre
            value: Valeur à valider
            
        Returns:
            (is_valid, error_message)
        """
        # Vérifier required
        if param.required and (value is None or str(value).strip() == ""):
            return False, f"Le paramètre '{param.name}' est obligatoire"
        
        # Vérifier type
        if value is not None and str(value).strip() != "":
            if param.type == "integer":
                try:
                    int(value)
                except (ValueError, TypeError):
                    return False, f"'{param.name}' doit être un entier"
            
            elif param.type == "date":
                try:
                    from datetime import datetime
                    if isinstance(value, str):
                        datetime.fromisoformat(value)
                except (ValueError, TypeError):
                    return False, f"'{param.name}' doit être une date valide (YYYY-MM-DD)"
        
        # Vérifier options (si liste)
        if param.type == "liste" and param.options_mode != "none":  # ✅ Remplacer 'select' par 'liste'
            options = ParameterService.resolve_parameter_options(param)
            if options and str(value) not in options:
                return False, f"'{value}' n'est pas une option valide pour '{param.name}'"
        
        return True, None
    
    @staticmethod
    def get_default_value(param: ParameterConfig) -> Any:
        """
        Retourne la valeur par défaut d'un paramètre.
        
        Args:
            param: Configuration du paramètre
            
        Returns:
            Valeur par défaut (ou première option si select)
        """
        if param.default is not None:
            return param.default
        
        # Pour les listes, prendre la première option
        if param.type == "liste":  # ✅ Remplacer 'select' par 'liste'
            options = ParameterService.resolve_parameter_options(param)
            if options:
                return options[0]
        
        # Valeurs par défaut selon type
        if param.type == "integer":
            return 0
        elif param.type == "date":
            from datetime import datetime
            return datetime.now().strftime("%Y-%m-%d")
        else:
            return ""
    
    @staticmethod
    def prepare_parameters_for_execution(
        params_config: List[ParameterConfig],
        user_values: Dict[str, Any]
    ) -> tuple[Dict[str, Any], List[str]]:
        """
        Prépare les paramètres pour l'exécution avec validation.
        
        Args:
            params_config: Liste des configurations de paramètres
            user_values: Valeurs fournies par l'utilisateur
            
        Returns:
            (parameters_dict, errors_list)
        """
        result = {}
        errors = []
        
        for param in params_config:
            value = user_values.get(param.name)
            
            # Utiliser valeur par défaut si non fournie
            if value is None or str(value).strip() == "":
                value = ParameterService.get_default_value(param)
            
            # Valider
            is_valid, error_msg = ParameterService.validate_parameter_value(param, value)
            if not is_valid:
                errors.append(error_msg)
                continue
            
            result[param.name] = value
        
        return result, errors