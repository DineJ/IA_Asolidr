# Guide de `clean_menu.py`

## Objectif

Ce script nettoie les historiques de menus **2025 et 2026** sans modifier les fichiers sources. Les commentaires et docstrings du code sont en **anglais** ; ce guide est en français.

## Installation

```bash
pip install pandas rapidfuzz
```

Placez `clean_menu.py` et `corrections.csv` dans le même dossier. Le fichier de règles est requis, sauf si vous indiquez un autre chemin avec `--corrections`.

## Commande

```bash
python clean_menu.py --entree data/processed --sortie data/processed/nettoyage --ecraser
```

Le seuil de suggestion est modifiable avec `--seuil 88` (de 0 à 100). `--ecraser` autorise le remplacement des résultats existants.

## Fichiers attendus

- `data/processed/historique_menus_2025.csv`
- `data/processed/historique_menus_2026.csv`
- `corrections.csv`, avec les colonnes `type;incorrect;correct`

Les fichiers annuels doivent contenir `chemin_source`, `ligne_excel` et les six champs : `entree`, `proteine`, `feculent`, `legumes`, `dessert_base`, `dessert_resto`.

## Étapes du nettoyage

1. **`charger_corrections`** lit les règles `mot` et `expression`. Deux corrections différentes pour la même entrée et le même type provoquent une erreur explicite.
2. **`nettoyer_avec_regles`** transforme le texte en majuscules, simplifie les espaces et reconnaît les valeurs vides. Ces changements de présentation ne sont pas journalisés.
3. Le script supprime la mention `(HALAL)` et trace cette suppression. Il harmonise ensuite les apostrophes et certains caractères Unicode.
4. Il applique les **expressions** validées (les plus longues en premier), puis les **mots** validés. Chaque remplacement est enregistré avec ses valeurs avant/après.
5. **`suggestions_rapidfuzz`** compare les mots restants au vocabulaire des corrections validées. Il crée des **suggestions seulement** : aucune correction approximative n'est appliquée automatiquement.
6. **`motifs_suspects`** signale notamment les quantités, consignes et libellés de service inhabituels. Les mentions `MENU`, les jours, `MIDI` et `SOIR` ne sont pas supprimés.
7. **`traiter_fichier`** conserve les références aux fichiers et lignes Excel et collecte les alertes et modifications.
8. **`main`** écrit les fichiers de sortie.

## Résultats

| Fichier | Contenu |
|---|---|
| `historique_menus_2025_nettoye.csv` | Menus 2025 nettoyés |
| `historique_menus_2026_nettoye.csv` | Menus 2026 nettoyés |
| `modifications.csv` | Historique détaillé des règles appliquées, avec provenance |
| `alert.csv` | Anomalies et suggestions RapidFuzz à examiner |

Les valeurs manquantes sont exportées sous la forme littérale `NaN`. Le CSV utilise `;` et UTF-8 avec BOM.

## Ajouter une correction

Dans `corrections.csv` :

```csv
type;incorrect;correct
mot;PUREE;PURÉE
expression;TARTE POMMES;TARTE AUX POMMES
```

`mot` cible un mot entier ; `expression` cible une suite précise de caractères, éventuellement au milieu d'un plat. Validez les suggestions de `alert.csv` avant de les ajouter au fichier de règles.

## Points de vigilance

- RapidFuzz mesure une **ressemblance**, pas la justesse linguistique.
- La version fournie recherche `(HALAL)` exactement (avec des espaces facultatifs à l'intérieur des parenthèses). Si vos sources contiennent aussi `(HALLAL)`, il faut étendre cette règle.
- Une règle de correction contradictoire dans `corrections.csv` bloque volontairement le traitement.
- Relancez ensuite `occurrences_menus.py` pour recalculer les statistiques.
