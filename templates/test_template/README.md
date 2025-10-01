# test_template

**Version:** 1.0
**Créé le:** 2025-10-01

## Description

Template de test

## Paramètres

- **marque** (string): Pas de description - Obligatoire

## Source de données

- Type: postgresql
- Tables requises: observations

## Structure des fichiers

```
test_template/
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
    template_name="test_template",
    parameters={
        "marque": "valeur",
    }
)
```