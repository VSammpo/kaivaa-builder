"""
Modèles pour les custom tables (SQL + Python)
"""

from typing import Optional
from pydantic import BaseModel, Field, validator


class CustomTableConfig(BaseModel):
    """Configuration d'une custom table"""
    
    table_name: str = Field(..., description="Nom de la table (ex: D009_custom)")
    description: Optional[str] = Field(None, description="Description de la table")
    
    # Code SQL
    sql_query: str = Field(..., description="Requête SQL de base")
    
    # Code Python (optionnel)
    python_code: Optional[str] = Field(None, description="Code Python pour post-traitement")
    
    # Paramètres
    parameters: list[str] = Field(default_factory=list, description="Paramètres utilisés (ex: ['sous_marque', 'segment'])")
    
    # Validation
    is_validated: bool = Field(False, description="La table a été testée")
    last_validation_result: Optional[str] = Field(None, description="Résultat dernier test")
    
    @validator('table_name')
    def validate_table_name(cls, v):
        """Valide le format du nom de table"""
        import re
        if not re.match(r'^[A-Z]\d{3}_[a-zA-Z_]+$', v):
            raise ValueError("Format attendu : D009_custom (lettre + 3 chiffres + _ + nom)")
        return v
    
    @validator('sql_query')
    def validate_sql(cls, v):
        """Validation basique du SQL"""
        v_upper = v.upper()
        if 'DROP' in v_upper or 'DELETE' in v_upper or 'TRUNCATE' in v_upper:
            raise ValueError("Opérations destructives interdites (DROP, DELETE, TRUNCATE)")
        if 'SELECT' not in v_upper:
            raise ValueError("La requête doit contenir SELECT")
        return v
    
    def get_python_function_name(self) -> str:
        """Retourne le nom de la fonction Python attendue"""
        return f"process_{self.table_name.lower()}"
    
    class Config:
        json_schema_extra = {
            "example": {
                "table_name": "D009_custom_analysis",
                "description": "Analyse personnalisée des prix",
                "sql_query": "SELECT * FROM observations WHERE sous_marque = '{sous_marque}'",
                "python_code": "def process_d009_custom_analysis(df):\n    return df.groupby('region').agg({'prix': 'mean'})",
                "parameters": ["sous_marque"]
            }
        }