"""Tests pour les utilitaires"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from backend.utils.file_utils import clean_filename, generate_batch_id, get_output_paths


def test_clean_filename():
    """Test du nettoyage de noms de fichiers"""
    assert clean_filename("Mon Fichier/Test") == "Mon_Fichier_Test"
    assert clean_filename("Test:2025") == "Test_2025"
    assert clean_filename("Normal") == "Normal"
    print("✅ Test clean_filename OK")


def test_generate_batch_id():
    """Test de génération de batch ID"""
    batch_id = generate_batch_id()
    assert len(batch_id) == 13  # YYYYMMDD_HHmm
    
    batch_with_prefix = generate_batch_id("test")
    assert batch_with_prefix.startswith("test_")
    print("✅ Test generate_batch_id OK")


def test_get_output_paths():
    """Test de génération des chemins de sortie"""
    paths = get_output_paths(
        study_name="Test Study",
        category="Gin",
        brand="BOMBAY",
        batch="20251001",
        distributor="Leclerc",
        template_name="suivi_commercial"
    )
    
    assert "excel_path" in paths
    assert "pptx_path" in paths
    assert "BOMBAY" in paths["excel_path"]
    assert "suivi_commercial" in paths["excel_path"]
    print("✅ Test get_output_paths OK")


if __name__ == "__main__":
    print("🧪 Tests utilitaires")
    test_clean_filename()
    test_generate_batch_id()
    test_get_output_paths()
    print("✅ Tous les tests passés")