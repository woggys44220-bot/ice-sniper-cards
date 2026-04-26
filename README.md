# ice-sniper-cards

Application Python (version simple) pour analyser des photos de cartes de hockey et générer un fichier Excel prêt pour la recherche de prix.

## Objectif de cette V1 (sans API payante)

- Déposer des images dans `photos/`
- Extraire des informations **basiques** depuis le nom de fichier (heuristiques)
- Générer une requête de prix propre
- Générer des liens de recherche:
  - eBay sold listings
  - 130point
- Créer `resultats_cartes_hockey.xlsx`
- **Ne jamais inventer de prix**: les colonnes de prix restent à `à vérifier`

## Fichiers générés

- `main.py`
- `requirements.txt`
- `photos/` (à remplir par l'utilisateur)
- `resultats_cartes_hockey.xlsx` (créé à l'exécution)

## Installation (Windows)

```powershell
cd chemin\vers\ice-sniper-cards
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## Utilisation

1. Copiez vos photos dans le dossier `photos/`.
2. Lancez le script:

```powershell
python main.py
```

3. Ouvrez le fichier `resultats_cartes_hockey.xlsx`.

## Format recommandé des noms de fichiers

La V1 lit le nom du fichier (pas encore d'OCR). Plus le nom est structuré, meilleurs seront les résultats.

Exemples:

- `Connor_McDavid_2015_UpperDeck_YoungGuns_YG201_Oilers_recto.jpg`
- `Sidney_Crosby_2005_Panini_Rookie_RC87_Penguins_verso.png`

## Colonnes du fichier Excel

- `fichier_photo`
- `joueur`
- `annee`
- `marque`
- `serie`
- `insert`
- `numero_carte`
- `equipe`
- `requete_prix`
- `lien_ebay`
- `lien_130point`
- `prix_bas`
- `prix_moyen`
- `prix_haut`
- `statut`

## Limites de la V1

- Pas d'OCR ni de vision IA
- Extraction basée sur des règles simples
- Certaines cartes resteront en `à vérifier`

## Feuille de route V2 (avancée)

Pour une version avancée:

1. Ajouter OCR local (`pytesseract` ou `easyocr`)
2. Détection texte recto/verso et fusion des infos
3. Normalisation intelligente (joueurs, équipes, marques)
4. Option de scraping/API pour récupérer des prix réels (toujours sans inventer)
5. Score de confiance par champ

## Rappel important

- Les prix ne sont **jamais** inventés.
- Si un prix n'est pas trouvé automatiquement, garder `à vérifier`.
