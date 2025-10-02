-- Template SQL pour Performance
-- Généré automatiquement par KAIVAA Builder

-- Paramètres disponibles :
-- - {Entreprise} : string

SELECT 
    *
FROM Performance
WHERE 1=1
  AND Entreprise = '{Entreprise}'
ORDER BY created_at DESC;
