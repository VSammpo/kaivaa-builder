# BCE_INPI

**Version:** 1.0
**Créé le:** 2025-10-03

## Description

Version de test

## Paramètres



## Source de données

- Type: excel
- Tables requises: Performance

## Structure des fichiers

```
BCE_INPI/
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
    template_name="BCE_INPI",
    parameters={

    }
)
```