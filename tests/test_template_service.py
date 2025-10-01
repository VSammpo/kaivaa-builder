"""Tests du service de templates"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.models.template_config import (
    TemplateConfig,
    ParameterConfig,
    DataSourceConfig
)


def test_template_service():
    """Test basique du service de templates"""
    
    # Initialiser la base
    DatabaseService.create_tables()
    
    # Créer une config de test
    config = TemplateConfig(
        name="test_template",
        version="1.0",
        description="Template de test",
        parameters=[
            ParameterConfig(
                name="marque",
                type="string",
                required=True,
                balise_ppt="[Marque]"
            )
        ],
        data_source=DataSourceConfig(
            type="postgresql",
            required_tables=["observations"]
        )
    )
    
    with DatabaseService.get_session() as db:
        service = TemplateService(db)
        
        # Créer un template
        template = service.create_template(
            config=config,
            user_id=1
        )
        
        print(f"✅ Template créé : {template.name} (ID: {template.id})")
        
        # Lister les templates
        templates = service.list_templates()
        print(f"✅ {len(templates)} templates trouvés")
        
        # Récupérer les stats
        stats = service.get_template_stats(template.id)
        print(f"✅ Stats : {stats}")


if __name__ == "__main__":
    print("🧪 Test du service de templates")
    test_template_service()
    print("✅ Tests terminés")