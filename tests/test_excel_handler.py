"""Tests pour excel_handler"""

import sys
from pathlib import Path

# Ajouter le projet au path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from backend.core.excel_handler import (
    excel_app_context,
    copy_template_excel,
    inject_filter_values
)


def test_excel_app_context():
    """Test du context manager Excel"""
    # Créer un fichier Excel de test
    import openpyxl
    
    test_file = project_root / "tests" / "test_data.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Test"
    wb.save(test_file)
    
    # Tester l'ouverture
    try:
        with excel_app_context(str(test_file)) as (app, wb):
            sheet = wb.sheets[0]
            value = sheet.range("A1").value
            assert value == "Test", f"Expected 'Test', got {value}"
            print("✅ Test excel_app_context OK")
    finally:
        test_file.unlink()


if __name__ == "__main__":
    print("🧪 Tests excel_handler")
    test_excel_app_context()
    print("✅ Tous les tests passés")