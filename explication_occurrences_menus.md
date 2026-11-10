# Guide de `occurrences_menus.py`

## Objectif

Ce script compte les occurrences de chaque **intitulé de plat** dans les menus nettoyés de **2025** et **2026**, toutes catégories confondues. Il ne corrige pas l'orthographe : cette étape appartient à `clean_menu.py`.

## Commande

```bash
python occurrences_menus.py --entree data/processed/nettoyage --sortie data/processed/nettoyage/occurrences_plats.csv --ecraser
```

Aucune dépendance externe n'est nécessaire : ce script utilise la bibliothèque standard Python.

## Entrées

Le script recherche :

- `historique_menus_2025_nettoye.csv`
- `historique_menus_2026_nettoye.csv`

Il lit les colonnes `entree`, `proteine`, `feculent`, `legumes`, `dessert_base`, `dessert_resto` lorsqu'elles sont présentes. Un fichier annuel absent est ignoré ; si les deux sont absents, le script échoue.

## Fonctionnement

1. **`nettoyer_pour_compter`** harmonise uniquement les majuscules et espaces, et ignore les marqueurs vides (`NaN`, etc.). Il ne découpe pas les intitulés contenant des virgules.
2. **`compter`** parcourt chaque cellule de plat non vide et incrémente un `Counter` par année. Si le même plat figure deux fois, il est compté deux fois.
3. **`main`** réunit les plats des deux années et les trie par **nombre total d'occurrences décroissant**, puis par nom alphabétique.
4. Le résultat est écrit dans `occurrences_plats.csv`, avec le séparateur `;` et l'encodage UTF-8 avec BOM.

## Sortie

| plat | occurrences_2025 | occurrences_2026 |
|---|---:|---:|
| PURÉE DE CAROTTES | 3 | 5 |
| POULET RÔTI | 2 | 1 |

Ces nombres sont **des exemples**, pas des résultats calculés sur vos fichiers.

## Points de vigilance

- Deux orthographes différentes sont comptées comme deux plats différents : enrichissez `corrections.csv`, relancez `clean_menu.py`, puis relancez ce script.
- Le script ne produit **ni catégorie**, ni colonne `total`.
- `--ecraser` est nécessaire pour remplacer un CSV déjà présent.
