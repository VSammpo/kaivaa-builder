# Test vide

**Version:** 1.0
**Créé le:** 2025-10-03

## Description

Fichier vide

## Paramètres



## Source de données

- Type: excel
- Tables requises: Performance

## Structure des fichiers

```
Test vide/
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
    template_name="Test vide",
    parameters={

    }
)
```