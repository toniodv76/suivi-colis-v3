# Suivi Colis V3

Application Flask pour suivi colis avec écran atelier.

## Nouveautés V3

- Statuts : A FAIRE / EN COURS / FAIT
- Couleurs : A FAIRE rouge, EN COURS bleu, FAIT vert
- Modification complète d'un colis
- Description longue
- Historique
- Export Excel mensuel
- Export Excel complet
- Compatible Render + PostgreSQL

## Variables Render

ACCESS_CODE=h2otech  
PYTHON_VERSION=3.12.7  
DATABASE_URL=Internal Database URL PostgreSQL

## Commandes Render

Build Command:
pip install -r requirements.txt

Start Command:
gunicorn app:app

## Routes

- / : gestion
- /ecran : affichage atelier
- /historique : historique
- /export/monthly : export du mois
- /export/all : export complet
