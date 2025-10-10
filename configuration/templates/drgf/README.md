# drgf

**Version:** 1.0
**Créé le:** 2025-10-09

## Description

Pas de description

## Paramètres



## Tables demandées

- Type: excel
- Tables demandées (par gabarit): 

## Structure des fichiers

```
drgf/
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
    template_name="drgf",
    parameters={

    }
)
```