# Test

**Version:** 1.0
**Créé le:** 2025-10-27

## Description

Analyse stratégique du marché reposant exclusivement sur du sell-out.

## Paramètres

- **Marque** (string): Choisir la marque que l'on souhaite mettre en avant. - Obligatoire

## Tables demandées

- Type: excel
- Tables demandées (par gabarit): 

## Structure des fichiers

```
Test/
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
    template_name="Test",
    parameters={
        "Marque": "valeur",
    }
)
```