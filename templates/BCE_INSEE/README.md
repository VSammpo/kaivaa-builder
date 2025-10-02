# BCE_INSEE

**Version:** 1.0
**Créé le:** 2025-10-02

## Description

Test

## Paramètres

- **Test** (string): Pas de description - Obligatoire

## Source de données

- Type: excel
- Tables requises: Performance

## Structure des fichiers

```
BCE_INSEE/
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
    template_name="BCE_INSEE",
    parameters={
        "Test": "valeur",
    }
)
```