# TESTV4

**Version:** 1.0
**Créé le:** 2025-10-02

## Description

Pas de description

## Paramètres

- **Entreprise** (string): Pas de description - Obligatoire
- **Background** (string): Pas de description - Obligatoire

## Source de données

- Type: excel
- Tables requises: Performance

## Structure des fichiers

```
TESTV4/
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
    template_name="TESTV4",
    parameters={
        "Entreprise": "valeur",
        "Background": "valeur",
    }
)
```