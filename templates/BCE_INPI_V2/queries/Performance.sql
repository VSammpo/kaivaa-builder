-- Template SQL pour Performance
-- Généré automatiquement par KAIVAA Builder

-- Paramètres disponibles :
-- - {Entreprise} : string
-- - {Background} : string

SELECT 
    *
FROM Performance
WHERE 1=1
  AND Entreprise = '{Entreprise}'
  AND Background = '{Background}'
ORDER BY created_at DESC;
