"""
Modèles Pydantic pour la configuration des templates
"""

from typing import List, Dict, Optional, Any
from pydantic import BaseModel, Field, validator
from datetime import datetime


class ParameterConfig(BaseModel):
    """Configuration d'un paramètre de template"""
    name: str = Field(..., description="Nom du paramètre (ex: sous_marque)")
    type: str = Field(..., description="Type: string, integer, date, list")
    required: bool = Field(True, description="Paramètre obligatoire")
    default: Optional[Any] = Field(None, description="Valeur par défaut")
    allowed_values: Optional[List[str]] = Field(None, description="Valeurs autorisées (pour type list)")
    description: Optional[str] = Field(None, description="Description du paramètre")
    balise_ppt: str = Field(..., description="Balise dans PowerPoint (ex: [Sous_Marque])")
    
    @validator('type')
    def validate_type(cls, v):
        allowed_types = ['string', 'integer', 'date', 'list']
        if v not in allowed_types:
            raise ValueError(f"Type doit être parmi {allowed_types}")
        return v


class DataSourceConfig(BaseModel):
    """Configuration de la source de données"""
    type: str = Field(..., description="Type: postgresql, mysql, excel, csv, api")
    connection_string: Optional[str] = Field(None, description="Chaîne de connexion")
    required_tables: List[str] = Field(default_factory=list, description="Tables nécessaires")
    
    @validator('type')
    def validate_type(cls, v):
        allowed_types = ['postgresql', 'mysql', 'sqlserver', 'excel', 'csv', 'api']
        if v not in allowed_types:
            raise ValueError(f"Type doit être parmi {allowed_types}")
        return v


class SlideMapping(BaseModel):
    """Configuration du mapping d'une slide"""
    slide_id: str = Field(..., description="ID de la slide (ex: A003)")
    sheet_name: str = Field(..., description="Nom de la feuille Excel source")
    excel_range: str = Field(..., description="Plage Excel (ex: A1:D10)")
    has_header: bool = Field(True, description="Première ligne = en-tête")


class ImageInjection(BaseModel):
    """Configuration d'une injection d'image"""
    type: str = Field(..., description="Type d'image (ex: product_image, logo)")
    pattern: str = Field(..., description="Pattern du chemin (ex: assets/{Marque}/{Produit}.png)")
    default_path: Optional[str] = Field(None, description="Image par défaut")
    position: Dict[str, float] = Field(..., description="Position {left, top}")
    size: Dict[str, float] = Field(..., description="Taille {max_width, max_height}")
    background: bool = Field(False, description="Placer en arrière-plan")
    loop_dependent: bool = Field(False, description="Image dépendante d'une boucle")  

class LoopConfig(BaseModel):
    """Configuration d'une boucle sur slides"""
    loop_id: str = Field(..., description="ID de la boucle (ex: Produits, Concurrents)")
    slides: List[str] = Field(..., description="Liste des slide IDs concernés")
    sheet_name: str = Field("Boucles", description="Feuille contenant le tableau Loop")


class TemplateConfig(BaseModel):
    """Configuration complète d'un template"""
    
    # Métadonnées
    name: str = Field(..., description="Nom du template")
    version: str = Field("1.0", description="Version du template")
    description: Optional[str] = Field(None, description="Description")
    created_by: Optional[str] = Field(None, description="Créateur")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    
    # Paramètres
    parameters: List[ParameterConfig] = Field(..., description="Paramètres du template")
    
    # Source de données
    data_source: DataSourceConfig = Field(..., description="Configuration source de données")
    
    # Mappings
    slide_mappings: List[SlideMapping] = Field(default_factory=list, description="Mappings slides → Excel")
    
    # Boucles
    loops: List[LoopConfig] = Field(default_factory=list, description="Configurations des boucles")
    
    # Images
    image_injections: Dict[str, List[ImageInjection]] = Field(
        default_factory=dict, 
        description="Images par slide_id"
    )
    
    # Chemins des fichiers
    ppt_template_path: Optional[str] = Field(None, description="Chemin du template PowerPoint")
    excel_template_path: Optional[str] = Field(None, description="Chemin du template Excel")
    
    # Statistiques
    estimated_generation_time: Optional[int] = Field(None, description="Temps estimé en secondes")
    
    class Config:
        json_schema_extra = {
            "example": {
                "name": "Suivi Commercial",
                "version": "1.0",
                "description": "Template de suivi commercial pour spiritueux",
                "parameters": [
                    {
                        "name": "sous_marque",
                        "type": "string",
                        "required": True,
                        "balise_ppt": "[Sous_Marque]"
                    }
                ],
                "data_source": {
                    "type": "postgresql",
                    "required_tables": ["observations", "dim_produits"]
                }
            }
        }
    
    def to_yaml(self) -> str:
        """Exporte la config en YAML"""
        import yaml
        return yaml.dump(self.model_dump(), default_flow_style=False, allow_unicode=True)
    
    @classmethod
    def from_yaml(cls, yaml_path: str) -> "TemplateConfig":
        """Charge une config depuis YAML"""
        import yaml
        with open(yaml_path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        return cls(**data)