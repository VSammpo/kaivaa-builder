# Crée un fichier temporaire : add_image_column.py
from sqlalchemy import create_engine, text
from backend.config import DatabaseConfig

engine = create_engine(DatabaseConfig.get_connection_string())

with engine.connect() as conn:
    conn.execute(text("ALTER TABLE templates ADD COLUMN card_image_path VARCHAR(500)"))
    conn.commit()
    print("✅ Colonne card_image_path ajoutée")