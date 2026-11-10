# Explication des tests unitaires — IA_Asolidr

## Objectif

Vérifier que `extract.py` identifie les différents formats Excel (2022 à 2026), sélectionne les bons onglets et extrait les données sans modifier les fichiers sources.

## Organisation

```text
IA_Asolidr/
├── extract.py
├── pytest.ini
├── data/
│   └── raw/
│       ├── 2022/
│       ├── 2023/
│       ├── 2024/
│       ├── 2025/
│       └── 2026/
└── tests/
    ├── conftest.py
    └── test_extract.py
```

## `conftest.py`

Pytest charge automatiquement ce fichier. Il rend disponibles deux *fixtures* :

- `project_root` : racine du dépôt `IA_Asolidr`.
- `raw_dir` : dossier réel `IA_Asolidr/data/raw`.

Il permet aussi d'importer `extract.py` depuis le dossier `tests/`.

## `test_extract.py`

### Tests unitaires

Les tests de détection vérifient :

- Le format `GRILLE` et le numéro de semaine pour 2022, 2023 et 2024.
- Le format `RECAP` pour 2025.
- Le format `MENUS` pour 2026.
- Le rejet d'un fichier hors dossier d'année.
- Le rejet d'un fichier 2022 sans numéro de semaine.
- La sélection de l'onglet `RECAP` ou `MENUS`.

Ces tests sont indépendants du contenu réel de `data/raw` : ils vérifient uniquement la logique de reconnaissance.

### Test d'intégration avec `data/raw`

`test_extract_from_real_raw_folder` lance réellement `extract.py` sur les fichiers `.xls` et `.xlsx` trouvés dans `data/raw`, y compris dans les sous-dossiers des années.

**Sécurité :** la sortie est créée dans un dossier temporaire `tmp_path` fourni par Pytest. Le test ne remplace donc pas `data/processed` et ne modifie pas les fichiers Excel sources.

Le test vérifie que la commande termine sans erreur, que `alert.csv` existe, que des historiques annuels sont produits et que leurs colonnes correspondent au format attendu.

Si `data/raw` n'existe pas ou ne contient aucun fichier Excel, le test d'intégration est **ignoré** (`SKIPPED`), tandis que les tests unitaires restent exécutables.

**Limite importante :** ce test ne garantit pas que chaque fichier a été traité correctement. L'extracteur peut signaler des problèmes dans `alert.csv` tout en terminant normalement. Il faut donc également consulter ce fichier après une extraction sur les données réelles. Les tests ne supposent pas que toutes les années sont présentes ni qu'elles contiennent un nombre fixe de lignes.

## Installation

Depuis la racine du dépôt :

```bash
python -m pip install -r requirement.txt
python -m pip install pytest
```

Conserver le `requirement.txt` existant ; `pytest` peut être installé séparément.

## Exécution

```bash
# Tous les tests du projet
python -m pytest -v

# Uniquement les tests d'extraction
python -m pytest tests/test_extract.py -v

# Un test précis
python -m pytest tests/test_extract.py::test_extract_from_real_raw_folder -v
```

## Différence entre les fichiers

| Fichier | Rôle |
| --- | --- |
| `pytest.ini` | Configure la découverte et l'affichage des tests. |
| `conftest.py` | Fournit les chemins et fixtures communs. |
| `test_extract.py` | Contient les assertions de détection et d'extraction. |
| `extract.py` | Script métier testé, non modifié. |

## Bonnes pratiques

- Exécuter les tests avant un `git push`.
- Garder des fichiers fictifs dans `tests/fixtures` pour les tests reproductibles à long terme.
- Utiliser `data/raw` surtout pour les tests d'intégration, car les données réelles évoluent.
- Ne jamais écrire les sorties de test directement dans `data/processed`.
