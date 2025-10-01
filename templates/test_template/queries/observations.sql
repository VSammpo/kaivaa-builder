-- Template SQL pour observations
-- Généré automatiquement par KAIVAA Builder

-- Paramètres disponibles :
-- - {marque} : string

SELECT 
    *
FROM observations
WHERE 1=1
  AND marque = '{marque}'
ORDER BY created_at DESC;
