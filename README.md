# ice-sniper-cards

Application Python pour analyser des photos de cartes de hockey, extraire le texte visible via OCR, puis générer un fichier Excel prêt pour la recherche de prix.

## Objectif de la V2

- Lire les photos dans `photos/`
- Utiliser **PIL + OpenCV + pytesseract** pour lire le contenu réel de l'image
- Remplir les colonnes : `joueur`, `annee`, `marque`, `serie`, `insert`, `numero_carte`, `equipe`
- Ajouter `texte_ocr` pour audit du texte détecté
- Générer une requête de prix propre
- Générer des liens:
  - eBay sold listings
  - 130point
- Créer `resultats_cartes_hockey.xlsx`
- Ne jamais inventer les prix (`prix_bas`, `prix_moyen`, `prix_haut` restent `à vérifier`)

## Installation (Windows)

```powershell
pip install -r requirements.txt
```

## Utilisation

```powershell
python main.py
```

## Pré-requis OCR Windows

Le package Python `pytesseract` nécessite l'installation de **Tesseract OCR** sur Windows.

1. Installer Tesseract OCR (application système)
2. Ajouter son dossier d'installation au `PATH`
3. Relancer le terminal

## Colonnes du fichier Excel

- `fichier_photo`
- `joueur`
- `annee`
- `marque`
- `serie`
- `insert`
- `numero_carte`
- `equipe`
- `texte_ocr`
- `requete_prix`
- `lien_ebay`
- `lien_130point`
- `prix_bas`
- `prix_moyen`
- `prix_haut`
- `statut`

## Règles importantes

- Le script lit d'abord le **texte de l'image** (pas seulement le nom du fichier).
- Si l'OCR ne lit pas assez bien, les champs restent `à vérifier`.
- Aucun prix n'est inventé.
