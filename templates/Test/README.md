# Test

**Version:** 1.0
**Créé le:** 2025-10-01

## Description

Pas de description

## Paramètres

- **Test_1** (string): Pas de description - Obligatoire

## Source de données

- Type: excel
- Tables requises: 

## Structure des fichiers

```
Test/
├── config.yaml           # Configuration du template
├── master.pptx          # Template PowerPoint
├── master.xlsx          # Template Excel
├── queries/             # Requêtes SQL
│   ├── table1.sql
│   └── table2.sql
└── README.md            # Ce fichier
```

## Utilisation

Pour générer un rapport avec ce template :

```python
from backend.services.report_service import ReportService

service = ReportService()
result = service.generate_report(
    template_name="Test",
    parameters={
        "Test_1": "valeur",
    }
)
```