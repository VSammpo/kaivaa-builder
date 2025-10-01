"""
Modèles Pydantic pour validation des configurations
"""

from .template_config import (
    TemplateConfig,
    ParameterConfig,
    DataSourceConfig,
    SlideMapping,
    ImageInjection,
    LoopConfig
)
from .custom_table import CustomTableConfig

__all__ = [
    "TemplateConfig",
    "ParameterConfig", 
    "DataSourceConfig",
    "SlideMapping",
    "ImageInjection",
    "LoopConfig",
    "CustomTableConfig"
]