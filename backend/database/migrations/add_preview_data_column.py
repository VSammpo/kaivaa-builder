# backend/database/migrations/add_preview_data_column.py
"""
Migration : Fusionner preview_columns + preview_rows en preview_data
"""

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import sqlite

def upgrade():
    # 1) Ajouter nouvelle colonne
    op.add_column('gabarit_default_sources', 
                  sa.Column('preview_data', sa.JSON(), nullable=True))
    
    # 2) Migrer les données
    conn = op.get_bind()
    sources = conn.execute(
        sa.text("SELECT id, preview_columns, preview_rows FROM gabarit_default_sources")
    ).fetchall()
    
    for source_id, cols, rows in sources:
        if cols and rows:
            preview_data = {"columns": cols, "rows": rows}
            conn.execute(
                sa.text("UPDATE gabarit_default_sources SET preview_data = :data WHERE id = :id"),
                {"data": preview_data, "id": source_id}
            )
    
    # 3) Supprimer anciennes colonnes
    with op.batch_alter_table('gabarit_default_sources') as batch_op:
        batch_op.drop_column('preview_columns')
        batch_op.drop_column('preview_rows')


def downgrade():
    # Rollback
    with op.batch_alter_table('gabarit_default_sources') as batch_op:
        batch_op.add_column(sa.Column('preview_columns', sa.JSON(), nullable=True))
        batch_op.add_column(sa.Column('preview_rows', sa.JSON(), nullable=True))
    
    # Migrer données inverse
    conn = op.get_bind()
    sources = conn.execute(
        sa.text("SELECT id, preview_data FROM gabarit_default_sources")
    ).fetchall()
    
    for source_id, preview_data in sources:
        if preview_data:
            cols = preview_data.get("columns")
            rows = preview_data.get("rows")
            conn.execute(
                sa.text("""
                    UPDATE gabarit_default_sources 
                    SET preview_columns = :cols, preview_rows = :rows 
                    WHERE id = :id
                """),
                {"cols": cols, "rows": rows, "id": source_id}
            )
    
    op.drop_column('gabarit_default_sources', 'preview_data')