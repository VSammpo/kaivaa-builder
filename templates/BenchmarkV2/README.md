# BenchmarkV2

**Version:** 1.0
**Créé le:** 2025-10-03

## Description

Evaluation de la performance

## Paramètres



## Source de données

- Type: excel
- Tables requises: 

## Structure des fichiers

```
BenchmarkV2/
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
    template_name="BenchmarkV2",
    parameters={

    }
)
```