# IA_Asolidr

**IA_Asolidr** est un projet de gestion de cuisine centrale visant à faciliter le traitement des menus et, à terme, la gestion des commandes et des documents de livraison.

Le pipeline Python documenté ici permet actuellement de **récupérer des menus depuis des classeurs Excel**, de **normaliser les intitulés des plats** et de **calculer leurs fréquences d'apparition en 2025 et 2026**. La description initiale du projet mentionne également un tableau de bord web, la planification des menus et les bons de livraison ; leur installation et leur exécution ne sont pas documentées ici, faute de configuration vérifiée dans les fichiers fournis.

## Sommaire

1. [Prérequis](#prérequis)
2. [Récupérer le projet](#récupérer-le-projet)
3. [Installation et démarrage automatiques](#installation-et-démarrage-automatiques)
4. [Préparer les données](#préparer-les-données)
5. [Exécuter les tests automatiques](#exécuter-les-tests-automatiques)
6. [Exécuter le pipeline complet](#exécuter-le-pipeline-complet)
7. [Comprendre les fichiers générés](#comprendre-les-fichiers-générés)
8. [Gérer les corrections orthographiques](#gérer-les-corrections-orthographiques)
9. [Options des scripts](#options-des-scripts)
10. [Résoudre les problèmes courants](#résoudre-les-problèmes-courants)
11. [Documentation technique](#documentation-technique)

## Prérequis

- **Git** pour cloner le dépôt.
- **Python 3** et `pip` ; utiliser une version de Python compatible avec les versions de dépendances indiquées dans `requirement.txt`.
- Les **fichiers Excel sources** (`.xls` ou `.xlsx`) pour exécuter l'extraction.
- Un terminal (exemples ci-dessous pour **Linux/macOS** ; sous Windows, les commandes d'activation du venv diffèrent).

Les principales bibliothèques utilisées par le pipeline sont `openpyxl`, `xlrd`, `pandas` et `rapidfuzz`. Le script de comptage utilise uniquement la bibliothèque standard de Python.

## Récupérer le projet

```bash
git clone -b main https://github.com/DineJ/IA_Asolidr.git
cd IA_Asolidr
```

Si le dépôt est déjà cloné :

```bash
git switch main
git pull origin main
```

> **Accès au dépôt :** le clonage nécessite les autorisations GitHub appropriées si le dépôt est privé.

## Installation et démarrage automatiques

La méthode recommandée sous **Linux/macOS** consiste à utiliser `setup_and_run.sh` depuis la racine du dépôt. Ce script :

1. Crée l'environnement virtuel `.env/` s'il n'existe pas (ou réutilise celui déjà présent).
2. Met à jour `pip` dans cet environnement.
3. Installe les dépendances de **`requirement.txt`** (nom au singulier, conservé tel quel).
4. Installe `pytest` et, si présent, `requirements-dev.txt`.
5. Lance l'ensemble des tests Pytest ; en cas d'échec, le pipeline n'est pas exécuté.
6. Lance `main.py`, qui exécute l'extraction, le nettoyage et le comptage dans cet ordre.

**Premier démarrage :** placez d'abord les classeurs Excel dans `data/raw/YYYY/` (voir section suivante), puis exécutez :

```bash
chmod +x setup_and_run.sh
./setup_and_run.sh
```

Pour régénérer les CSV existants :

```bash
./setup_and_run.sh --ecraser
```

| Commande | Fonction |
| --- | --- |
| `./setup_and_run.sh` | Installation, tous les tests, puis pipeline. |
| `./setup_and_run.sh --ecraser` | Même chose en autorisant l'écrasement des CSV générés. |
| `./setup_and_run.sh --tests-only` | Installation et tests, sans lancer le pipeline. |
| `./setup_and_run.sh --skip-tests` | Installation et pipeline, sans exécuter les tests. |
| `./setup_and_run.sh --help` | Affiche les options disponibles. |

`--tests-only` et `--skip-tests` ne peuvent pas être utilisés ensemble. Le script ne supprime pas `.env/` et ne modifie pas les classeurs Excel sources. **Attention :** il met à jour `pip` et réinstalle/vérifie les dépendances à chaque exécution ; une connexion réseau peut être nécessaire.

### Installation manuelle (alternative)

Si vous préférez contrôler chaque commande :

```bash
python3 -m venv .env
source .env/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirement.txt
python -m pip install pytest
python -m pytest -v
python main.py --ecraser
```

Sous Windows, le script Bash nécessite un environnement compatible (par exemple WSL). Pour l'activation manuelle avec PowerShell : ` .\.env\Scripts\Activate.ps1 `. Pour quitter l'environnement virtuel : `deactivate`.

Ajoutez `.env/` à `.gitignore` : un environnement virtuel ne doit pas être versionné.

## Préparer les données

Placez les classeurs Excel dans **`data/raw`**, organisés par année. L'extracteur recherche les fichiers `.xls` et `.xlsx` dans cette arborescence et exploite les formats propres aux années concernées.

```text
IA_Asolidr/
├── setup_and_run.sh
├── main.py
├── pytest.ini
├── tests/
│   ├── conftest.py
│   ├── test_extract.py
│   ├── test_clean_menu.py
│   ├── test_occurrences.py
│   └── test_main.py
├── extract.py
├── clean_menu.py
├── occurrences_menus.py
├── corrections.csv
├── requirement.txt
├── README.md
├── explication_extract.md
├── explication_clean_menu.md
├── explication_occurrences_menus.md
└── data/
    ├── raw/
    │   ├── 2022/
    │   ├── 2023/
    │   ├── 2024/
    │   ├── 2025/
    │   └── 2026/
    └── processed/
        └── nettoyage/
```

Cette arborescence est **indicative** : les dossiers et fichiers générés peuvent ne pas encore exister après le clonage. Créez au besoin les dossiers des données :

```bash
mkdir -p data/raw data/processed
```

**Important :** les classeurs Excel ne sont pas fournis automatiquement par ce README. S'ils ne sont pas versionnés dans Git, demandez à l'équipe où récupérer les données autorisées, puis placez-les dans les dossiers annuels correspondants. Ne publiez pas de données confidentielles dans le dépôt.

## Exécuter les tests automatiques

Les tests se trouvent dans `tests/`. `pytest.ini` configure leur découverte et `tests/conftest.py` fournit des éléments partagés (fixtures). Les tests couvrent notamment l'extraction, le nettoyage, les occurrences et l'orchestration par `main.py`.

Pour exécuter **tous les tests**, sans lancer le pipeline :

```bash
./setup_and_run.sh --tests-only
```

Si l'environnement est déjà configuré :

```bash
.env/bin/python -m pytest -v
```

Pour un fichier précis :

```bash
.env/bin/python -m pytest tests/test_extract.py -v
```

**Données de test :** certains tests utilisent des fichiers fictifs dans `tests/fixtures/`, tandis qu'un test d'intégration peut lire les classeurs présents dans `data/raw`. Ce dernier peut être ignoré si les données requises sont absentes. Les résultats de tests sont écrits dans des dossiers temporaires afin de préserver `data/processed`.

Si vous obtenez `No module named pytest`, lancez `./setup_and_run.sh --tests-only` ou installez `pytest` dans le bon environnement : `.env/bin/python -m pip install pytest`.

## Exécuter le pipeline complet

La méthode recommandée est `./setup_and_run.sh` (ou `./setup_and_run.sh --ecraser`). `main.py` lance les trois étapes ci-dessous dans le bon ordre. Les commandes individuelles restent disponibles pour le diagnostic et les exécutions partielles.

### Étape 1 — Extraire les données Excel

```bash
python extract.py --input data/raw --output data/processed --ecraser
```

Le script parcourt les fichiers Excel et produit les fichiers `historique_menus_YYYY.csv` pour les années disposant de données. Il traite les formats 2022–2024 et 2025–2026 différemment. Les classeurs originaux ne sont pas modifiés.

L'extraction crée aussi un **`data/processed/alert.csv`** contenant ses avertissements et erreurs techniques.

### Étape 2 — Nettoyer les menus 2025 et 2026

```bash
python clean_menu.py --entree data/processed --sortie data/processed/nettoyage --ecraser
```

Ce script :

- lit uniquement les historiques **2025 et 2026** ;
- uniformise la casse et les espaces, et écrit les valeurs manquantes sous la forme `NaN` ;
- applique les corrections validées dans **`corrections.csv`** ;
- retire la mention **`(HALAL)`** dans la version du script documentée ;
- conserve les mentions `MENU`, les jours de la semaine, `MIDI` et `SOIR` ;
- consigne les changements significatifs dans **`modifications.csv`** ;
- utilise **RapidFuzz** pour proposer des corrections possibles dans **`alert.csv`**, **sans les appliquer automatiquement**.

> La règle `(HALAL)` du script examiné ne couvre pas automatiquement la variante `(HALLAL)` à deux `L`. Vérifiez le comportement si cette graphie apparaît dans vos sources.

### Étape 3 — Compter les occurrences des plats

```bash
python occurrences_menus.py --entree data/processed/nettoyage --sortie data/processed/nettoyage/occurrences_plats.csv --ecraser
```

Le comptage lit les fichiers **déjà nettoyés**, compte chaque apparition d'un plat dans les six catégories alimentaires et produit une ligne par intitulé distinct, avec les fréquences de 2025 et de 2026.

**Ordre à respecter :** `extract.py` → `clean_menu.py` → `occurrences_menus.py`. Si vous modifiez `corrections.csv`, relancez **au minimum les étapes 2 et 3**.

### Commandes rapides (à copier-coller)

Pour lancer uniquement le pipeline, sans installation ni tests :

```bash
.env/bin/python main.py --ecraser
```

Pour exécuter chaque étape manuellement :

```bash
source .env/bin/activate
python extract.py --input data/raw --output data/processed --ecraser
python clean_menu.py --entree data/processed --sortie data/processed/nettoyage --ecraser
python occurrences_menus.py --entree data/processed/nettoyage --sortie data/processed/nettoyage/occurrences_plats.csv --ecraser
```

## Comprendre les fichiers générés

| Fichier | Rôle |
| --- | --- |
| `data/processed/historique_menus_2022.csv` à `historique_menus_2026.csv` | Historique annuel extrait, si des données sont disponibles pour l'année. |
| `data/processed/alert.csv` | Avertissements et erreurs **d'extraction**. |
| `data/processed/nettoyage/historique_menus_2025_nettoye.csv` | Menus 2025 nettoyés. |
| `data/processed/nettoyage/historique_menus_2026_nettoye.csv` | Menus 2026 nettoyés. |
| `data/processed/nettoyage/alert.csv` | Anomalies et suggestions **de nettoyage**. |
| `data/processed/nettoyage/modifications.csv` | Journal des corrections appliquées, avec provenance et valeurs avant/après. |
| `data/processed/nettoyage/occurrences_plats.csv` | Nombre d'apparitions par plat en 2025 et en 2026. |

Les fichiers CSV sont séparés par **`;`** et écrits en **UTF-8 avec BOM**, pour faciliter leur ouverture dans un tableur.

Le fichier final d'occurrences contient exactement :

```csv
plat;occurrences_2025;occurrences_2026
PURÉE DE CAROTTES;3;5
POULET RÔTI;2;1
```

*Les valeurs de cet exemple sont fictives.* Les occurrences sont triées par fréquence totale décroissante, puis par intitulé alphabétique. Il n'y a ni colonne `categorie` ni colonne `total`.

## Gérer les corrections orthographiques

Le fichier **`corrections.csv`** doit être placé à côté de `clean_menu.py` (sauf utilisation de l'option `--corrections`). Il utilise les colonnes suivantes :

```csv
type;incorrect;correct
mot;PUREE;PURÉE
mot;BETTRAVE;BETTERAVE
expression;TARTE POMMES;TARTE AUX POMMES
```

- **`mot`** : corrige un mot entier dans un intitulé.
- **`expression`** : corrige une suite de mots ou une formulation précise.
- **RapidFuzz** : détecte des ressemblances avec les graphies validées et ajoute des suggestions à `data/processed/nettoyage/alert.csv`. Ces suggestions doivent être **vérifiées par une personne** avant leur ajout à `corrections.csv`.

Si deux lignes du CSV définissent des corrections différentes pour la même entrée et le même type, le script s'arrête avec une erreur, par exemple `Corrections contradictoires pour COQUILETTE`. Corrigez le conflit dans `corrections.csv`, puis relancez le nettoyage.

## Options des scripts

### `main.py`

```bash
python main.py --help
```

- `--input` : répertoire des classeurs sources (défaut : `data/raw`).
- `--output` : répertoire des CSV (défaut : `data/processed`).
- `--corrections` : fichier des règles orthographiques.
- `--seuil` : seuil RapidFuzz (défaut : `88`).
- `--ecraser` : autorise le remplacement des sorties générées.
- `--inclure-annexes` : transmet l'option d'extraction des annexes.


### `extract.py`

```bash
python extract.py --help
```

- `--input` : dossier des fichiers Excel (défaut : `data/raw`).
- `--output` : dossier des CSV extraits (défaut : `data/processed`).
- `--ecraser` : autorise le remplacement des sorties existantes.
- `--inclure-annexes` : inclut les annexes prises en charge par l'extracteur.

### `clean_menu.py`

```bash
python clean_menu.py --help
```

- `--entree` : dossier des historiques extraits (défaut : `data/processed`).
- `--sortie` : dossier des fichiers nettoyés (défaut : `data/processed/nettoyage`).
- `--corrections` : chemin vers le CSV des corrections validées.
- `--seuil` : seuil de ressemblance RapidFuzz, entre 0 et 100 (défaut : `88`). Un seuil plus faible augmente les suggestions, y compris les faux positifs.
- `--ecraser` : autorise le remplacement des sorties existantes.

### `occurrences_menus.py`

```bash
python occurrences_menus.py --help
```

- `--entree` : dossier des historiques nettoyés (défaut : `data/processed/nettoyage`).
- `--sortie` : chemin du fichier d'occurrences (défaut : `data/processed/nettoyage/occurrences_plats.csv`).
- `--ecraser` : autorise le remplacement du CSV de sortie.

## Résoudre les problèmes courants

| Problème | Solution |
| --- | --- |
| `ModuleNotFoundError: No module named 'pandas'` ou `'rapidfuzz'` | Activez `.env`, puis relancez `python -m pip install -r requirement.txt`. |
| `No module named pytest` | Exécutez `./setup_and_run.sh --tests-only` ou `.env/bin/python -m pip install pytest`. |
| Les tests échouent et le pipeline ne démarre pas | Corrigez les tests en échec, ou utilisez explicitement `--skip-tests` pour un lancement sans tests. |
| `Fichier de corrections introuvable` | Vérifiez que `corrections.csv` est à côté de `clean_menu.py`, ou utilisez `--corrections chemin/vers/corrections.csv`. |
| `Corrections contradictoires pour ...` | Cherchez les règles en double dans `corrections.csv` et gardez une correction cohérente. |
| Aucun historique 2025/2026 trouvé | Exécutez d'abord `extract.py` et vérifiez `data/processed`. |
| Aucun fichier Excel traité | Vérifiez la présence des `.xls`/`.xlsx` dans `data/raw` et ses dossiers annuels. |
| Un fichier de sortie existe déjà | Ajoutez `--ecraser` pour régénérer les résultats. |
| Les occurrences ne reflètent pas une nouvelle correction | Relancez `clean_menu.py`, puis `occurrences_menus.py`. |
| Trop de suggestions RapidFuzz | Augmentez le seuil, par exemple `--seuil 93`, puis vérifiez les alertes. |
| Les résultats diffèrent d'une ancienne exécution | Vérifiez les classeurs sources, les règles de `corrections.csv`, les versions des dépendances et les alertes. |

## Documentation technique

Pour comprendre les fonctions et les choix de traitement en détail, consultez :

- [`explication_extract.md`](explication_extract.md) : lecture des classeurs Excel, détection des formats et export des historiques.
- [`explication_clean_menu.md`](explication_clean_menu.md) : règles de nettoyage, RapidFuzz, journal des modifications et alertes.
- [`explication_occurrences_menus.md`](explication_occurrences_menus.md) : méthode de comptage et structure des résultats.
- [`explication_unitest.md`](explication_unitest.md) : présentation de Pytest, des fixtures et des tests (si ce document est ajouté au dépôt).

## Bonnes pratiques de contribution

Avant de proposer une modification :

1. Travaillez sur une branche Git dédiée et documentez les changements.
2. N'ajoutez pas les fichiers de l'environnement virtuel `.env/` ni les données sensibles au dépôt.
3. Vérifiez les nouvelles règles dans `corrections.csv` : les corrections automatiques influencent directement les statistiques.
4. Testez les commandes du pipeline sur des données autorisées, et contrôlez `alert.csv` et `modifications.csv`.
5. Mettez à jour ce README et les guides Markdown si le comportement des scripts change.

---

**Périmètre de ce guide :** les scripts Python d'extraction, de nettoyage, de comptage, leur orchestration avec `main.py` et l'automatisation par `setup_and_run.sh`. Les étapes de déploiement d'un éventuel tableau de bord web ou de génération des bons de livraison devront être ajoutées lorsqu'elles seront disponibles et vérifiées.