# 01a - Promotion (méthode)

**Version:** 1.0
**Créé le:** 2025-10-10

## Description

Première étape PROMO : 
* Calcul de la baseline et de l'uplift promotionnel.
* Présentation de la méthode et définition des paramètres.
* Premières analyses.

## Paramètres



## Tables demandées

- Type: excel
- Tables demandées (par gabarit): 

## Structure des fichiers

```
01a - Promotion (méthode)/
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
    template_name="01a - Promotion (méthode)",
    parameters={

    }
)
```