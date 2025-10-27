# Analyse de la politique tariaire

**Version:** 1.0
**Créé le:** 2025-10-19

## Description

Pas de description

## Paramètres



## Tables demandées

- Type: excel
- Tables demandées (par gabarit): 

## Structure des fichiers

```
Analyse de la politique tariaire/
├── config.yaml           # Configuration du template livrable (tables demandées)
├── master.pptx          # Master PPT (facultatif)
├── master.xlsx          # Master Excel (obligatoire)
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
    template_name="Analyse de la politique tariaire",
    parameters={

    }
)
```