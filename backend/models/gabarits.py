# backend/models/gabarits.py
from typing import List, Literal, Optional
from pydantic import BaseModel, field_validator

ColumnType = Literal["text", "number", "integer", "date", "boolean"]

class GabaritColumn(BaseModel):
    name: str
    type: ColumnType
    is_key: bool = False  # coche = fait partie de la clé composite

    @field_validator("name")
    @classmethod
    def _normalize_name(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("column.name vide")
        # norme simple MVP : snake_case ASCII
        v = v.replace(" ", "_")
        return v

class TableGabarit(BaseModel):
    name: str
    version: str = "v1"
    description: Optional[str] = None
    columns: List[GabaritColumn] = []

    @field_validator("name")
    @classmethod
    def _normalize_gab_name(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("gabarit.name vide")
        return v
