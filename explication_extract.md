# Explication du script `extract.py`

Ce document décrit **la version actuelle du script `extract.py`**. Il distingue les données des grilles 2022–2024 de celles des menus 2025–2026, et tient compte des colonnes simplifiées, de l'export optionnel des annexes et de la suppression du fichier `erreurs_par_fichier.csv`.

## 1. Objectif et organisation

Le programme parcourt les classeurs `.xls` et `.xlsx` présents dans les sous-dossiers annuels de `data/raw`, choisit une feuille principale par classeur, puis écrit un fichier `historique_menus_YYYY.csv` par année contenant des données. Il conserve les noms des plats tels que saisis, à l'exception de l'harmonisation des espaces, et ne modifie pas les Excel originaux.

| Années | Format interne | Onglet principal | Unité d'extraction |
|---|---|---|---|
| 2022–2024 | `GRILLE` | Onglet correspondant à la semaine, avec replis contrôlés | Une proposition par cellule |
| 2025 | `RECAP` | `Recap` prioritaire, sinon `Menus` | Un repas par ligne |
| 2026 | `MENUS` | `Menus` | Un repas par ligne |

Le numéro de semaine des fichiers 2022–2024 sert à choisir l'onglet, **pas à vérifier les dates**.

## 2. Installation et commandes

```bash
python -m pip install openpyxl xlrd
python extract.py --input data/raw --output data/processed
```

Pour régénérer les fichiers déjà présents :

```bash
python extract.py --input data/raw --output data/processed --ecraser
```

Pour inclure les annexes `Livraison` et `Médian` des années 2022–2023 :

```bash
python extract.py --input data/raw --output data/processed --ecraser --inclure-annexes
```

`--input` et `--output` sont facultatifs (valeurs par défaut : `data/raw` et `data/processed`). `--ecraser` autorise le remplacement des sorties existantes. `--inclure-annexes` est désactivé par défaut.

## 3. Fichiers générés et colonnes

Les fichiers CSV utilisent le séparateur **`;`** et l'encodage **UTF-8 avec BOM** (`utf-8-sig`). Ils n'ont **pas tous les mêmes colonnes** :

### Historique 2022–2024 : six colonnes

| Colonne | Signification |
|---|---|
| `chemin_source` | Chemin relatif du fichier Excel, avec son année |
| `ligne_excel` | Ligne de la proposition dans le classeur |
| `colonne_excel` | Colonne Excel de la proposition (numéro, pas lettre) |
| `categorie` | Catégorie déduite de la ligne de grille |
| `sous_categorie` | Sous-catégorie de grille (`ENTREE 1`, `PLAT 2`, etc.) |
| `proposition` | Texte de la cellule alimentaire |

Une ligne CSV = **une proposition alimentaire dans une cellule Excel**.

### Historique 2025–2026 : huit colonnes

| Colonne | Signification |
|---|---|
| `chemin_source` | Chemin relatif du classeur |
| `ligne_excel` | Ligne du repas dans Excel |
| `entree` | Entrée |
| `proteine` | Protéine / plat principal |
| `feculent` | Féculent |
| `legumes` | Légumes |
| `dessert_base` | Dessert de base |
| `dessert_resto` | Dessert restaurant |

Une ligne CSV = **un repas**, dont les six champs alimentaires peuvent être vides individuellement.

### Autres sorties

- **`alert.csv`** : toutes les erreurs et tous les avertissements techniques, avec les colonnes `fichier_source`, `niveau`, `code`, `onglet`, `nombre_feuilles`, `message`. Filtrer `niveau = ERREUR` pour voir uniquement les erreurs.
- **`annexes.csv`** : créé **uniquement** avec `--inclure-annexes`. Colonnes : `chemin_source`, `onglet`, `ligne_excel`, `colonne_excel`, `categorie`, `sous_categorie`, `proposition`.
- **`erreurs_par_fichier.csv` n'est plus généré.** Les colonnes `type_enregistrement`, `type_menu`, `enfant`, `type_feuille` et `onglet` ont été retirées des **historiques**. L'onglet reste présent dans les alertes et les annexes.

Un historique annuel n'est créé que si l'année comporte au moins un enregistrement extrait. Le script ne supprime pas automatiquement les anciens fichiers de sortie devenus inutiles.

## 4. Constantes et détection des formats

`MENU_FIELDS` définit les six champs des repas récents. `GRID_OUTPUT_COLUMNS`, `MEAL_OUTPUT_COLUMNS` et `ANNEX_OUTPUT_COLUMNS` fixent les colonnes des trois types d'exports.

`expected_sheet(path, root)` lit le dossier année. Pour 2022–2024, il récupère également le numéro de semaine dans le nom de fichier. `sheet_matches()` teste la conformité d'un nom d'onglet. `choose_sheet()` sélectionne une seule feuille :

- en 2025, `Recap` est préféré à `Menus` ;
- en 2026, l'onglet `Menus` est attendu ;
- en 2022–2024, l'onglet correspondant à la semaine est choisi, avec priorité à `midi` lorsqu'il existe plusieurs correspondances ; à défaut, l'unique onglet `midi` ou l'unique feuille peut être accepté avec avertissement ; les situations ambiguës ne sont pas résolues au hasard.

`read_selected_sheet()` lit les `.xls` avec `xlrd` et les `.xlsx` avec `openpyxl` (`read_only=True`, `data_only=True`). Il copie **A1:V65** de la feuille principale et fournit un accesseur `cell(ligne, colonne)` avec coordonnées Excel commençant à 1.

## 5. Extraction 2022–2024 : `extract_grid()`

Les sept colonnes candidates sont **D, F, H, J, L, N et P** (`GRID_COLUMNS = (4, 6, 8, 10, 12, 14, 16)`). Les catégories dépendent des lignes définies par `GRID_ROWS` :

| Lignes Excel | Catégorie | Sous-catégories |
|---|---|---|
| 7–10 | `entree` | `ENTREE 1` à `ENTREE 4` |
| 11–13 | `plat` | `PLAT 1`, `PLAT 2`, `ASE SOIR` (C13 remplace le libellé de la ligne 13 s'il est renseigné) |
| 14 | `poisson` | `POISSON` |
| 15–16 | `legume` | `LEGUME 1`, `LEGUME 2` |
| 17 | `feculent` | `FECULENTS` |
| 18–20 | `laitage` | `LAITAGE 1` à `LAITAGE 3` |
| 21–22 | `fromage` | `FROMAGE 1`, `FROMAGE 2` |
| 23, 25 | `dessert` | `DESSERT 1`, `DESSERT 2` |
| 24, 26 | `fruit` | `FRUIT 1`, `FRUIT 2` |

Le script examine **20 lignes × 7 colonnes = 140 positions** par grille. Les cellules vides et les étiquettes de jour seules sont ignorées. Il ne reconstitue pas des repas complets à partir des propositions historiques.

## 6. Extraction 2025–2026 : `find_header()` et `extract_meals()`

`find_header()` recherche les en-têtes dans les **lignes 3 à 6**, colonnes **A à Q**. Il reconnaît `entree`, `proteine` ou `plat principal`, `feculent`, `legumes` et les deux types de desserts. Il retient la ligne présentant le plus des quatre champs essentiels (`entree`, `proteine`, `feculent`, `legumes`) ; **au moins trois** sont nécessaires.

`extract_meals()` parcourt les lignes de `max(5, header_row + 1)` à **18 incluse**. Chaque champ est lu dans **la colonne détectée par son en-tête**. Le mapping reste **fixe sur toutes les lignes**.

**Correction importante :** si une cellule d'entrée contient `mercredi soir`, cette valeur est ignorée comme étiquette de jour, **sans décaler les autres colonnes**. L'ancien décalage `+1` pouvait placer `RIZ` dans `proteine` alors que la source Excel le mettait dans `feculent`. Le paramètre `recap` a été supprimé de la fonction.

Une ligne sans aucune valeur alimentaire est ignorée. Les autres deviennent des repas dans le CSV annuel. Aucun champ de menu enfant n'est exporté.

## 7. Traitement du texte et traçabilité

- `cell_text()` convertit les cellules en texte, remplace les espaces insécables et réduit les espaces répétés. Les cellules vides et les booléens deviennent des chaînes vides.
- `match_text()` supprime les différences de casse et d'accents **uniquement pour comparer les éléments de structure**.
- `is_day_label()` reconnaît un jour seul, éventuellement suivi de `midi` ou `soir` ; une description qui contient seulement le mot « lundi » parmi d'autres mots n'est pas automatiquement éliminée.
- `source_info()` conserve en interne `fichier_source`, `chemin_source`, `onglet` et `ligne_excel`. L'onglet et le nom seul du fichier ne sont plus écrits dans les historiques simplifiés, mais servent encore au suivi interne et au dédoublonnage.

Le script ne corrige ni l'orthographe des plats, ni les quantités, ni les descriptions coupées entre cellules. Il ne détecte pas automatiquement qu'une boisson a été saisie dans la colonne `legumes`.

## 8. Annexes facultatives

`annex_kind()` reconnaît les onglets dont le nom contient `livraison` ou `median` (avec ou sans accent). Si `--inclure-annexes` est actif **et** que le classeur appartient à 2022 ou 2023, `read_annex_cells()` lit toutes leurs cellules non vides, sans la limite A1:V65 ; `extract_annex()` conserve leur texte et leurs coordonnées.

Ces données peuvent être logistiques. Elles sont exportées dans **`annexes.csv` séparément**, jamais dans les historiques alimentaires. Sans option, les annexes ne sont pas extraites.

## 9. Déroulement de `main()`

1. Lire les options, vérifier le dossier d'entrée et refuser l'écrasement non autorisé des fichiers de sortie.
2. Chercher récursivement les `.xls` et `.xlsx`, sauf les fichiers temporaires commençant par `~$`.
3. Pour chaque fichier, déterminer son format et sa feuille principale ; enregistrer les problèmes techniques dans une liste d'alertes, sans arrêter le traitement des autres classeurs.
4. Extraire les propositions historiques ou les repas récents ; éventuellement extraire les annexes 2022–2023.
5. Dédoublonner par **coordonnées source**, jamais par nom de plat : `(chemin_source, onglet, ligne_excel)` pour les repas, et la même clé complétée par `colonne_excel` pour les propositions et annexes.
6. Regrouper par dossier année et écrire les historiques avec leurs colonnes respectives ; écrire éventuellement `annexes.csv`, puis **`alert.csv` uniquement** pour les diagnostics.
7. Afficher un bilan avec fichiers trouvés, traités, non traités, erreurs, avertissements, repas, propositions et annexes.

Parmi les codes d'alerte figurent `FEUILLE_INTROUVABLE`, `NOM_FEUILLE_DIFFERENT`, `ENTETES_INVALIDES`, `FEUILLES_IGNOREES`, `ANNEXE_VIDE`, `LECTURE_ANNEXE`, `AUCUNE_DONNEE`, `ANNEE_INTROUVABLE` et `EXCEPTION`.

**Attention : zéro erreur technique ne garantit pas l'absence d'erreurs dans les plats ou leurs catégories.**

## 10. Limites et précautions

- Les grilles 2022–2024 reposent sur des coordonnées fixes ; un changement de modèle peut nécessiter une adaptation.
- Les menus 2025–2026 nécessitent des en-têtes reconnaissables entre A3 et Q6 et ne lisent les repas que jusqu'à la ligne 18.
- La fenêtre de lecture des feuilles principales est limitée à A1:V65.
- L'année est celle du dossier source, non celle d'une date inscrite dans Excel.
- Le dédoublonnage préserve deux plats identiques présents à des coordonnées différentes.
- Le CSV d'alertes traite les **problèmes techniques** et non les erreurs culinaires ou orthographiques.
- Le script ne supprime pas automatiquement un ancien `erreurs_par_fichier.csv` présent dans le dossier de sortie : il cesse simplement de le produire.