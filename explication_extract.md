# Guide de compréhension du script d’extraction des menus Excel

## 1. Vue d’ensemble : ce qui entre et ce qui sort

Le programme reçoit un dossier d’entrée (`data/raw` par défaut) contenant des sous-dossiers d’années, par exemple `data/raw/2023/Menu sem 15.xlsx`. Il parcourt récursivement les fichiers `.xls` et `.xlsx`, identifie leur format selon l’année, sélectionne la feuille principale, extrait ses données, et enregistre les résultats dans `historique_menus.csv`. Les anomalies sont consignées dans `alert.csv`.

| Année | Format déterminé par le programme | Feuille principale attendue | Données extraites |
|---|---|---|---|
| 2022–2024 | `GRILLE` | Feuille correspondant au numéro de semaine, ou repli contrôlé | Propositions individuelles selon des coordonnées fixes |
| 2025 | `RECAP` | `Recap` en priorité, sinon `Menus` | Repas structurés en colonnes alimentaires |
| 2026 | `MENUS` | `Menus` | Repas structurés en colonnes alimentaires |

**Exception importante :** pour 2022 et 2023, le programme cherche également les onglets supplémentaires contenant « livraison » ou « Médian ». Il conserve alors **chaque cellule non vide**, sans prétendre qu’il s’agit nécessairement d’un plat.

**Chaîne d’exécution :** `main()` → recherche des fichiers → `expected_sheet()` → `read_selected_sheet()` / `choose_sheet()` → éventuellement `extract_annex()` → `find_header()` puis `extract_meals()` **ou** `extract_grid()` → dédoublonnage → `write_csv()`.

## 2. Variables de configuration : quelles positions Excel sont réellement utilisées ?

| Variable | Contenu exact | Rôle dans l’extraction |
|---|---|---|
| `MENU_FIELDS` | `entree`, `proteine`, `feculent`, `legumes`, `dessert_base`, `dessert_resto` | Six champs que `extract_meals()` cherche à renseigner pour chaque repas récent |
| `MENU_COLUMNS` | Provenance + `type_menu` + six champs + `enfant` | Colonnes attendues pour les enregistrements de repas |
| `PROPOSAL_COLUMNS` | Provenance + colonne, catégorie, sous-catégorie, proposition, type de feuille | Constante présente mais **non utilisée** dans l’assemblage final |
| `GRID_ROWS` | Lignes Excel **7 à 18 et 21 à 26**, avec des trous | Associe chaque ligne historique à une catégorie et une sous-catégorie. Par exemple, ligne 7 = `entree / ENTREE 1`, ligne 14 = `poisson / POISSON` |
| `GRID_COLUMNS` | `(4, 6, 8, 10, 12, 14, 16)` | Colonnes Excel **D, F, H, J, L, N, P** de la grille historique. Il y a sept colonnes parcourues |
| `EXPECTED_HEADERS` | `entree`, `proteine`, `feculent`, `legumes` | Quatre en-têtes utilisés pour reconnaître un tableau récent ; au moins **trois** doivent être détectés |
| `DAY_LABEL` | Jours de lundi à dimanche, éventuellement suivis de `MIDI` ou `SOIR` | Évite de prendre une étiquette de jour pour un nom de plat |

Les coordonnées Excel utilisées dans le guide sont **numérotées à partir de 1** : `cell(7, 4)` correspond à **D7**. En revanche, les listes Python commencent à l’indice 0 : c’est pour cela que l’accesseur `cell()` utilise `row - 1` et `column - 1`.

## 3. Bibliothèques : leur fonction dans **ce** programme

- **`argparse`** : récupère `--input`, `--output`, `--ecraser` depuis la commande. C’est lui qui détermine où chercher les Excel et où enregistrer les CSV.
- **`csv`** : écrit les dictionnaires de résultats dans les deux CSV, en respectant l’ordre des colonnes et le séparateur `;`.
- **`re`** : reconnaît les dossiers d’année (`20xx`), les numéros de semaine dans les noms de fichiers, les noms d’onglets et les étiquettes de jour.
- **`unicodedata`** : retire les différences d’accents pour comparer les en-têtes et noms d’onglets, **sans modifier les intitulés de plats enregistrés**.
- **`Counter`** (`collections`) : compte les fichiers traités par format, les annexes lues et certaines erreurs.
- **`Path`** (`pathlib`) : parcourt les dossiers, examine les extensions, calcule les chemins relatifs et crée le dossier de sortie.
- **`openpyxl.load_workbook`** : lit les `.xlsx` sans les modifier, en récupérant les valeurs de cellules.
- **`xlrd`** : lit les anciens `.xls` ; il est importé uniquement lorsqu’un tel fichier doit être ouvert.

## 4. Explication de la logique, fonction par fonction

Chaque section précise les **paramètres reçus**, les **variables calculées**, les **plages de parcours** et le **résultat renvoyé**. Les extraits de code sont tirés du script source.

### 4.0 Constantes et bornes dans le code source (lignes 20–43)

```python
020  MENU_FIELDS = ('entree', 'proteine', 'feculent', 'legumes', 'dessert_base', 'dessert_resto')
021  MENU_COLUMNS = ('fichier_source', 'chemin_source', 'onglet', 'ligne_excel',
022                  'type_menu', *MENU_FIELDS, 'enfant')
023  PROPOSAL_COLUMNS = ('fichier_source', 'chemin_source', 'onglet', 'ligne_excel',
024                      'colonne_excel', 'categorie', 'sous_categorie', 'proposition', 'type_feuille')
025  
026  # The historical layout uses fixed rows for categories and fixed columns for days.
027  GRID_ROWS = {
028      7: ('entree', 'ENTREE 1'), 8: ('entree', 'ENTREE 2'),
029      9: ('entree', 'ENTREE 3'), 10: ('entree', 'ENTREE 4'),
030      11: ('plat', 'PLAT 1'), 12: ('plat', 'PLAT 2'),
031      13: ('plat', 'ASE SOIR'), 14: ('poisson', 'POISSON'),
032      15: ('legume', 'LEGUME 1'), 16: ('legume', 'LEGUME 2'),
033      17: ('feculent', 'FECULENTS'), 18: ('laitage', 'LAITAGES'),
034      21: ('fromage', 'FROMAGE 1'), 22: ('fromage', 'FROMAGE 2'),
035      23: ('dessert', 'DESSERT 1'), 24: ('fruit', 'FRUIT 1'),
036      25: ('dessert', 'DESSERT 2'), 26: ('fruit', 'FRUIT 2'),
037  }
038  GRID_COLUMNS = (4, 6, 8, 10, 12, 14, 16)
039  EXPECTED_HEADERS = ('entree', 'proteine', 'feculent', 'legumes')
040  DAY_LABEL = re.compile(
041      r'(?:LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)'
042      r'(?:\s+(?:MIDI|SOIR))?', re.IGNORECASE
043  )
```


### 4.1 `cell_text(value)` — lignes 46–59

**Mission :** Nettoyer une valeur Excel sans réécrire les noms de plats.

- **`value`** est la valeur brute fournie par `openpyxl` ou `xlrd` ; elle peut être vide (`None`), booléenne, numérique ou textuelle.
- Si `value is None`, la fonction renvoie `''` : une cellule absente devient un texte vide. Si c’est un booléen, elle renvoie aussi `''`, pour ne pas transformer `True`/`False` en proposition culinaire.
- Pour les autres valeurs, `str(value)` produit du texte ; `replace('\u00a0', ' ')` remplace les espaces insécables ; `split()` puis `' '.join(...)` éliminent les espaces en trop.
- **Conséquence :** `"  Riz   pilaf "` devient `"Riz pilaf"`, mais le nom du plat n’est ni traduit ni uniformisé en catégorie.


**Code source correspondant :**

```python
046  def cell_text(value):
047      """Return a readable cell string without changing its culinary meaning.
048  
049      Args:
050          value: Raw Excel cell value, possibly None, a number or a string.
051      Returns:
052          str: Original text with whitespace standardized, or '' for empty cells.
053      """
054      if value is None:
055          return ''
056      # Boolean cells are metadata rather than food descriptions.
057      if isinstance(value, bool):
058          return ''
059      return ' '.join(str(value).replace('\u00a0', ' ').split())
```


### 4.2 `match_text(value)` — lignes 62–71

**Mission :** Comparer les libellés de structure indépendamment de la casse et des accents.

- **`value`** est généralement un nom d’onglet ou un en-tête de colonne, et non une proposition à enregistrer.
- **`normalized`** reçoit la version de `cell_text(value)` mise en minuscules de manière robuste (`casefold`) puis décomposée avec `NFKD`.
- La compréhension `char for char in normalized if not unicodedata.combining(char)` enlève les marques d’accent ; ainsi `"Entrée"` peut être comparé à `"entree"` et `"Médian"` à `"median"`.
- **Point crucial :** cette normalisation sert aux **comparaisons**. Le texte sauvegardé dans les CSV provient de `cell_text()`, pas de `match_text()`.


**Code source correspondant :**

```python
062  def match_text(value):
063      """Normalize text only for detecting workbook structure, not CSV values.
064  
065      Args:
066          value: Cell value or worksheet name.
067      Returns:
068          str: Accent-insensitive, case-folded text.
069      """
070      normalized = unicodedata.normalize('NFKD', cell_text(value).casefold())
071      return ''.join(char for char in normalized if not unicodedata.combining(char))
```


### 4.3 `is_day_label(value)` — lignes 74–82

**Mission :** Distinguer un jour de semaine d’un contenu alimentaire.

- **`value`** est le contenu d’une cellule candidate.
- `match_text(value).strip()` permet de reconnaître par exemple `"Lundi"`, `"MARDI MIDI"` ou `"dimanche soir"`.
- `DAY_LABEL.fullmatch(...)` exige que **toute** la cellule corresponde à une étiquette de jour ; `"Poulet lundi"` n’est donc pas éliminé simplement parce qu’il contient le mot « lundi ».
- **Résultat :** `True` signifie « étiquette structurelle à ne pas exporter comme plat », `False` signifie « ce filtre ne l’exclut pas ».


**Code source correspondant :**

```python
074  def is_day_label(value):
075      """Detect a weekday or weekday/service label in a structural cell.
076  
077      Args:
078          value: Cell content.
079      Returns:
080          bool: Whether the entire value is a recognized day label.
081      """
082      return bool(DAY_LABEL.fullmatch(match_text(value).strip()))
```


### 4.4 `expected_sheet(path, root)` — lignes 85–112

**Mission :** Déterminer le format attendu à partir du chemin du fichier.

- **`path`** est le chemin du classeur courant, **`root`** le dossier d’entrée (`args.input`).
- **`relative_parts`** découpe le chemin relatif. Exemple : `2023/Menu sem 15.xlsx` donne les parties `('2023', 'Menu sem 15.xlsx')`.
- **`years`** collecte les noms de dossiers à quatre chiffres commençant par `20`, convertis en entiers. `relative_parts[:-1]` exclut le nom du fichier. **`year = years[-1]`** choisit le dernier dossier année trouvé.
- Pour **2026**, le résultat est `('MENUS', None)` ; pour **2025**, `('RECAP', None)` : aucun numéro de semaine n’est nécessaire à la sélection de l’onglet.
- Pour **2022 à 2024**, `re.search(...)` cherche `sem` ou `semaine`, éventuellement séparé par des espaces, points, tirets ou underscores, suivi de **1 ou 2 chiffres**. **`match.group(1)`** fournit le numéro ; `int(...)` le convertit en entier. Exemple : `Menu sem 15.xlsx` → `('GRILLE', 15)`.
- Si le dossier année manque, si la semaine historique est introuvable, ou si l’année est hors plage, une `ValueError` remonte vers `main()` et sera consignée en alerte.


**Code source correspondant :**

```python
085  def expected_sheet(path, root):
086      """Return the exact expected sheet name type and week for a workbook.
087  
088      Args:
089          path (Path): Source Excel file.
090          root (Path): Root directory containing year folders.
091      Returns:
092          tuple[str, int | None]: Expected format and week number.
093      Raises:
094          ValueError: If the year or historical week cannot be identified.
095      """
096      relative_parts = path.relative_to(root).parts
097      years = [int(part) for part in relative_parts[:-1]
098               if re.fullmatch(r"20\d{2}", part)]
099      if not years:
100          raise ValueError("Year folder missing (expected data/raw/YYYY/)")
101      year = years[-1]
102      if year == 2026:
103          return 'MENUS', None
104      if year == 2025:
105          return 'RECAP', None
106      if 2022 <= year <= 2024:
107          match = re.search(r"(?:sem(?:aine)?[\s_.-]*)(\d{1,2})(?!\d)",
108                            match_text(path.stem))
109          if not match:
110              raise ValueError("Week number not found in file name")
111          return 'GRILLE', int(match.group(1))
112      raise ValueError(f"Unsupported year: {year}")
```


### 4.5 `sheet_matches(name, fmt, week)` — lignes 115–126

**Mission :** Vérifier si un nom d’onglet correspond au format et à la semaine attendus.

- **`name`** est un nom d’onglet ; **`fmt`** vaut `MENUS`, `RECAP` ou `GRILLE` ; **`week`** est un entier pour les grilles, sinon `None`.
- **`normalized`** est le nom d’onglet rendu comparable : accents et casse neutralisés.
- Si `fmt == 'MENUS'`, seul l’onglet nommé `menus` est accepté. Si `fmt == 'RECAP'`, `recap` ou `menus` sont acceptés.
- Pour `GRILLE`, l’expression régulière accepte le numéro de semaine avec un préfixe `sem`/`semaine` facultatif, des zéros initiaux facultatifs et le suffixe `midi` facultatif. **Le numéro doit correspondre exactement à `week`** : la semaine 15 n’est pas la semaine 16.
- **Résultat :** booléen utilisé par `choose_sheet()` pour constituer la liste des onglets admissibles.


**Code source correspondant :**

```python
115  def sheet_matches(name, fmt, week):
116      """Check whether a sheet is suitable for the selected year's layout."""
117      normalized = match_text(name).strip()
118      if fmt == 'MENUS':
119          return normalized == 'menus'
120      if fmt == 'RECAP':
121          return normalized in ('recap', 'menus')
122      # Accept week labels with an optional 'midi' suffix.
123      return bool(re.fullmatch(
124          r"(?:sem(?:aine)?[\s_.-]*)?0*" + str(week) + r"(?:[\s_.-]*midi)?",
125          normalized,
126      ))
```


### 4.6 `choose_sheet(names, fmt, week)` — lignes 129–158

**Mission :** Choisir un onglet principal en appliquant un ordre de préférence explicite.

- **`names`** est la liste des onglets réellement présents dans le classeur ; **`fmt`** et **`week`** viennent de `expected_sheet()`.
- En format **`RECAP`**, la boucle externe parcourt **deux préférences, dans cet ordre** : `('recap', 'menus')`. Pour chaque préférence, la boucle interne examine **tous les noms de `names`**. Un onglet `Recap` l’emporte donc sur `Menus`, même si `Menus` apparaît avant dans le classeur.
- En format **`MENUS`**, `next(...)` retourne le **premier** onglet qui satisfait `sheet_matches()`, ou `None` si aucun n’existe.
- En format **`GRILLE`**, **`matches`** contient les onglets correspondant au numéro de semaine. Le tri donne la priorité à un nom contenant `midi` ; à égalité, il préserve l’ordre initial via `names.index(name)`.
- Si aucune correspondance de semaine n’existe, **`midi`** contient les onglets dont le nom contient le mot autonome `midi`. S’il n’y en a **qu’un**, il est utilisé **avec avertissement**. Sinon, si le classeur ne possède **qu’un onglet**, cet onglet est utilisé **avec avertissement**.
- Si plusieurs onglets restent ambigus, la fonction renvoie `(None, None)` : elle **n’invente pas** un choix.
- **Résultat :** `(nom_onglet_choisi, avertissement_éventuel)`. `main()` transforme l’avertissement en ligne d’`alert.csv`.


**Code source correspondant :**

```python
129  def choose_sheet(names, fmt, week):
130      """Choose exactly one sheet; never guess among ambiguous multi-sheet grids.
131  
132      Returns:
133          tuple[str | None, str | None]: Selected sheet and optional warning.
134      """
135      if fmt == 'RECAP':
136          # Older 2025 workbooks use RECAP; later ones use MENUS.
137          for preferred in ('recap', 'menus'):
138              for name in names:
139                  if match_text(name).strip() == preferred:
140                      return name, None
141          return None, None
142      if fmt == 'MENUS':
143          return next((name for name in names if sheet_matches(name, fmt, week)), None), None
144  
145      matches = [name for name in names if sheet_matches(name, fmt, week)]
146      if matches:
147          # Prefer the recap grid over other sheets of the same week.
148          matches.sort(key=lambda name: ('midi' not in match_text(name), names.index(name)))
149          return matches[0], None
150      midi = [name for name in names if re.search(r'(?<![a-z])midi(?![a-z])', match_text(name))]
151      if len(midi) == 1:
152          return midi[0], (f"sheet name {midi[0]!r} does not match week {week}; "
153                           "using the only midi sheet")
154      if len(names) == 1:
155          # Many historical workbooks were copied without renaming their only tab.
156          return names[0], (f"sheet name {names[0]!r} does not match week {week}; "
157                            "using the only sheet")
158      return None, None
```


### 4.7 `read_selected_sheet(path, fmt, week)` — lignes 161–208

**Mission :** Ouvrir le bon onglet et créer un accès homogène aux cellules.

- **`path`** identifie le classeur ; **`fmt`**, **`week`** servent à sélectionner l’onglet.
- Si l’extension est **`.xls`**, le script utilise `xlrd` ; sinon, pour les fichiers retenus (`.xlsx`), il utilise `openpyxl` en lecture seule avec `data_only=True`.
- **`names`** contient tous les noms d’onglets ; **`selected`** est le nom retenu par `choose_sheet()` ; **`selection_warning`** est un éventuel motif d’avertissement ; **`skipped`** contient les autres noms.
- **`rows`** mémorise uniquement une fenêtre de **65 lignes × 22 colonnes** pour la feuille principale, soit les lignes Excel **1–65** et colonnes **A–V**. En `.xls`, `range(65)` parcourt les indices internes **0–64**, et `range(22)` les indices **0–21** ; les cellules hors dimensions existantes deviennent `None`.
- En `.xlsx`, `iter_rows(min_row=1, max_row=65, min_col=1, max_col=22, values_only=True)` lit la même fenêtre avec des coordonnées Excel directement numérotées à partir de 1.
- La fonction interne **`cell(row, column, data=rows)`** convertit les coordonnées Excel en indices Python : `data[row - 1][column - 1]`. Exemple : `cell(7, 4)` lit **D7**. Hors de la fenêtre chargée, elle renvoie `None`.
- Le classeur est fermé **avant** le retour ; `cell` continue de fonctionner car il lit la copie `rows` déjà mémorisée.
- **Résultat :** cinq valeurs `(selected, cell, names, skipped, selection_warning)`. Le deuxième élément est **une fonction utilisable** par `find_header()`, `extract_meals()` et `extract_grid()`, pas une cellule unique.


**Code source correspondant :**

```python
161  def read_selected_sheet(path, fmt, week):
162      """Read only the expected worksheet and report the other worksheet names.
163  
164      Args:
165          path (Path): Excel workbook.
166          fmt (str): MENUS, RECAP or GRILLE.
167          week (int | None): Week for historical grids.
168      Returns:
169          tuple[str | None, callable | None, list[str], list[str], str | None]:
170              Selected sheet, accessor, all names, skipped names and selection warning.
171      """
172      if path.suffix.lower() == '.xls':
173          import xlrd
174          workbook = xlrd.open_workbook(str(path), on_demand=True)
175          try:
176              names = workbook.sheet_names()
177              selected, selection_warning = choose_sheet(names, fmt, week)
178              skipped = [name for name in names if name != selected]
179              if selected is None:
180                  return None, None, names, skipped, None
181              sheet = workbook.sheet_by_name(selected)
182              # Cache only the required region before closing the workbook.
183              rows = [[sheet.cell_value(r, c) if r < sheet.nrows and c < sheet.ncols
184                       else None for c in range(22)] for r in range(65)]
185          finally:
186              workbook.release_resources()
187      else:
188          workbook = load_workbook(path, read_only=True, data_only=True)
189          try:
190              names = workbook.sheetnames
191              selected, selection_warning = choose_sheet(names, fmt, week)
192              skipped = [name for name in names if name != selected]
193              if selected is None:
194                  return None, None, names, skipped, None
195              sheet = workbook[selected]
196              sheet.reset_dimensions()
197              rows = list(sheet.iter_rows(min_row=1, max_row=65, min_col=1,
198                                          max_col=22, values_only=True))
199          finally:
200              workbook.close()
201  
202      def cell(row, column, data=rows):
203          """Return a 1-based cell value from the selected worksheet only."""
204          if 1 <= row <= len(data) and 1 <= column <= len(data[row - 1]):
205              return data[row - 1][column - 1]
206          return None
207  
208      return selected, cell, names, skipped, selection_warning
```


### 4.8 `read_annex_cells(path, sheet_name)` — lignes 212–236

**Mission :** Lire intégralement les annexes historiques sans supposer de structure de menu.

- **`path`** désigne le classeur ; **`sheet_name`** désigne l’onglet annexe précis à lire.
- Contrairement à `read_selected_sheet()`, **aucune limite fixe 65 × 22 n’est appliquée** : le parcours couvre toutes les lignes et colonnes exposées par le lecteur Excel.
- Avec **`xlrd`**, `r` parcourt les indices de lignes de **0 à `sheet.nrows - 1`**, et `c` les colonnes de **0 à `sheet.ncols - 1`**. Le résultat enregistre `(r + 1, c + 1, texte)` pour retrouver les coordonnées Excel réelles.
- Avec **`openpyxl`**, `enumerate(..., 1)` commence directement à **1** pour les numéros de lignes et de colonnes.
- Le filtre `if cell_text(...)` ne conserve que les cellules dont le texte nettoyé n’est pas vide.
- **Résultat :** une liste de triplets `(ligne_excel, colonne_excel, valeur)`. Le contenu peut être une date, un commentaire, une donnée de livraison ou un plat : la fonction ne le classe pas sémantiquement.


**Code source correspondant :**

```python
212  def read_annex_cells(path, sheet_name):
213      """Read populated cells of a 2022-2023 annex, without assuming a menu layout.
214  
215      The returned positions are the real 1-based Excel coordinates. Formula cells
216      are read with their cached values, as in the original extraction.
217      """
218      if path.suffix.lower() == '.xls':
219          import xlrd
220          workbook = xlrd.open_workbook(str(path), on_demand=True)
221          try:
222              sheet = workbook.sheet_by_name(sheet_name)
223              return [(r + 1, c + 1, cell_text(sheet.cell_value(r, c)))
224                      for r in range(sheet.nrows) for c in range(sheet.ncols)
225                      if cell_text(sheet.cell_value(r, c))]
226          finally:
227              workbook.release_resources()
228      workbook = load_workbook(path, read_only=True, data_only=True)
229      try:
230          sheet = workbook[sheet_name]
231          sheet.reset_dimensions()
232          return [(r, c, cell_text(value))
233                  for r, values in enumerate(sheet.iter_rows(values_only=True), 1)
234                  for c, value in enumerate(values, 1) if cell_text(value)]
235      finally:
236          workbook.close()
```


### 4.9 `annex_kind(sheet_name)` — lignes 239–246

**Mission :** Reconnaître uniquement les deux catégories d’annexes demandées.

- **`sheet_name`** est un nom d’onglet non sélectionné comme feuille principale.
- **`normalized`** est ce nom sans distinction de casse ni d’accents.
- Si le nom contient `livraison`, la fonction renvoie `'livraison'`. Sinon, s’il contient `median` (ce qui couvre `Médian`), elle renvoie `'median'`.
- Si aucun des deux mots n’apparaît, elle renvoie `None` : cet onglet n’est pas extrait comme annexe et peut faire l’objet d’une alerte de feuille ignorée.
- **Périmètre :** `main()` n’appelle cette classification pour extraction que sur les classeurs **2022 et 2023**.


**Code source correspondant :**

```python
239  def annex_kind(sheet_name):
240      """Only the two explicitly requested historical annex types qualify."""
241      normalized = match_text(sheet_name)
242      if 'livraison' in normalized:
243          return 'livraison'
244      if 'median' in normalized:
245          return 'median'
246      return None
```


### 4.10 `extract_annex(path, root, sheet_name, kind)` — lignes 249–255

**Mission :** Transformer les cellules d’une annexe en enregistrements traçables.

- **`path`**, **`root`** et **`sheet_name`** identifient la provenance ; **`kind`** vaut `'livraison'` ou `'median'`.
- `read_annex_cells(path, sheet_name)` fournit des triplets **`row, col, value`** pour **chaque cellule non vide**.
- Pour chaque triplet, `source_info(...)` renseigne le fichier, son chemin relatif, l’onglet et la ligne. Le dictionnaire ajoute **`colonne_excel = col`**, **`categorie = kind`**, **`sous_categorie = 'cellule_brute'`**, **`proposition = value`**, **`type_feuille = kind`**.
- **Exemple :** si l’onglet `Livraison` contient `Camion 2` en B8, l’enregistrement garde la ligne **8**, la colonne **2**, le texte **`Camion 2`** et la catégorie **`livraison`**. Il n’est pas transformé en repas.
- **Résultat :** une liste de dictionnaires, ensuite ajoutés à `annexes` par `main()`.


**Code source correspondant :**

```python
249  def extract_annex(path, root, sheet_name, kind):
250      """Store raw nonempty cells; do not mislabel delivery metadata as dishes."""
251      return [{**source_info(path, root, sheet_name, row),
252               'colonne_excel': col, 'categorie': kind,
253               'sous_categorie': 'cellule_brute', 'proposition': value,
254               'type_feuille': kind}
255              for row, col, value in read_annex_cells(path, sheet_name)]
```


### 4.11 `find_header(cell)` — lignes 258–288

**Mission :** Identifier automatiquement les colonnes alimentaires dans les formats récents.

- **`cell`** est la fonction d’accès à la feuille principale, créée par `read_selected_sheet()`.
- **`best_mapping`**, **`best_row`**, **`best_score`** mémorisent la meilleure ligne d’en-têtes trouvée jusqu’ici ; le score démarre à **0**.
- La boucle **`for row in range(3, 7)`** teste exactement les lignes Excel **3, 4, 5 et 6**. Pour chacune, **`for column in range(1, 18)`** inspecte les colonnes **1 à 17**, donc **A à Q**.
- **`heading`** reçoit le texte de la cellule d’en-tête normalisé par `match_text()` ; **`found`** associe un nom de champ (`entree`, `proteine`, etc.) à son numéro de colonne Excel.
- Les tests acceptent notamment `proteine` **ou** `plat principal` pour le champ `proteine`. Pour les desserts, un en-tête contenant `dessert` et `resto`/`restaurant` devient `dessert_resto`, sinon `dessert_base`.
- `found.setdefault(champ, column)` garde **la première colonne** reconnue pour chaque champ sur cette ligne ; un second en-tête semblable ne remplace pas le premier.
- **`score`** compte combien des **quatre** champs de `EXPECTED_HEADERS` figurent dans `found` ; il ne compte pas les desserts. La meilleure ligne est celle avec le score le plus élevé ; à égalité, la première gagnante reste sélectionnée.
- Si **`best_score < 3`**, la feuille n’est pas considérée comme un tableau récent valide : résultat `(None, None)`, puis alerte `ENTETES_INVALIDES` dans `main()`.
- Sinon, le résultat est **`(best_mapping, best_row)`**. Exemple : `({'entree': 3, 'proteine': 5, 'feculent': 7, 'legumes': 9}, 4)` signifie que les en-têtes sont sur la ligne **4**, avec les données attendues en **C, E, G, I**.


**Code source correspondant :**

```python
258  def find_header(cell):
259      """Find food columns in recent worksheets using their header labels.
260  
261      Args:
262          cell (callable): 1-based Excel cell accessor.
263      Returns:
264          tuple[dict | None, int | None]: Column mapping and header row.
265      """
266      best_mapping, best_row, best_score = None, None, 0
267      for row in range(3, 7):
268          found = {}
269          for column in range(1, 18):
270              heading = match_text(cell(row, column))
271              if 'entree' in heading:
272                  found.setdefault('entree', column)
273              elif 'proteine' in heading or 'plat principal' in heading:
274                  found.setdefault('proteine', column)
275              elif 'feculent' in heading:
276                  found.setdefault('feculent', column)
277              elif 'legume' in heading:
278                  found.setdefault('legumes', column)
279              elif 'dessert' in heading and ('resto' in heading or 'restaurant' in heading):
280                  found.setdefault('dessert_resto', column)
281              elif 'dessert' in heading:
282                  found.setdefault('dessert_base', column)
283          score = sum(field in found for field in EXPECTED_HEADERS)
284          if score > best_score:
285              best_mapping, best_row, best_score = found, row, score
286      if best_score < 3:
287          return None, None
288      return best_mapping, best_row
```


### 4.12 `source_info(path, root, sheet, row)` — lignes 291–304

**Mission :** Associer chaque résultat à sa position d’origine dans les Excel.

- **`path`** est le chemin complet ou relatif du classeur ; **`root`** est le dossier d’entrée ; **`sheet`** le nom d’onglet ; **`row`** le numéro de ligne Excel.
- **`fichier_source`** prend seulement `path.name` (exemple : `Menu sem 15.xlsx`).
- **`chemin_source`** prend `path.relative_to(root).as_posix()` (exemple : `2023/Menu sem 15.xlsx`) pour distinguer deux fichiers de même nom dans des années différentes.
- **`onglet`** et **`ligne_excel`** permettent de retrouver la donnée dans le classeur.
- **Résultat :** un dictionnaire commun réutilisé pour les repas, propositions historiques et cellules d’annexe.


**Code source correspondant :**

```python
291  def source_info(path, root, sheet, row):
292      """Build provenance fields for an extracted cell or meal.
293  
294      Args:
295          path (Path): Workbook path.
296          root (Path): Input directory.
297          sheet (str): Worksheet name.
298          row (int): Excel row number.
299      Returns:
300          dict: Source file, relative path, sheet and row.
301      """
302      return dict(fichier_source=path.name,
303                  chemin_source=path.relative_to(root).as_posix(),
304                  onglet=sheet, ligne_excel=row)
```


### 4.13 `extract_meals(path, root, sheet, cell, mapping, header_row, recap)` — lignes 307–338

**Mission :** Construire une ligne de repas à partir des colonnes reconnues dans un tableau récent.

- **`path`, `root`, `sheet`** identifient la provenance ; **`cell`** lit une cellule ; **`mapping`** indique la colonne de chaque champ ; **`header_row`** est la ligne d’en-têtes ; **`recap`** indique si l’onglet sélectionné est nommé `Recap`.
- **`meals = []`** accumule les repas valides de cet onglet.
- La boucle **`range(max(5, header_row + 1), 19)`** parcourt les lignes Excel depuis **la plus grande valeur entre 5 et la ligne suivant les en-têtes**, jusqu’à **18 incluse**. Exemple : si `header_row = 4`, elle examine **5–18** (**14 lignes**) ; si `header_row = 6`, elle examine **7–18** (**12 lignes**).
- **`shifted`** vaut `True` uniquement si l’onglet n’est **pas** `Recap`, si le champ `entree` existe dans `mapping`, **et** si la cellule à sa colonne contient une étiquette de jour. Cette situation signale un décalage d’une colonne pour les valeurs de repas sur cette ligne.
- **`offset = 1 if shifted else 0`** ajoute donc **une colonne** à la lecture lorsque `shifted` est vrai. Exemple : `mapping['entree'] = 3` et `offset = 1` → l’entrée est lue en **colonne 4 (D)** au lieu de **3 (C)**.
- La boucle **`for field in MENU_FIELDS`** passe exactement sur les **six champs** définis en début de script. **`column = mapping.get(field)`** récupère leur colonne ; si le champ n’a pas été détecté, **`value = ''`**.
- **`values[field]`** reçoit le texte nettoyé de la cellule, sauf si ce texte est une étiquette de jour : il est alors remplacé par `''`. **`values['enfant'] = ''`** initialise systématiquement ce champ à vide : aucune extraction spécifique enfant n’est effectuée ici.
- Si **aucun des six champs alimentaires** n’est renseigné (`not any(...)`), la ligne est ignorée. Sinon, le programme ajoute à `meals` un dictionnaire contenant la provenance, **`type_menu = 'principal'`** et les valeurs des champs.
- **Résultat :** une liste de repas. Chaque repas correspond à **une ligne Excel**, pas à chaque cellule séparément.


**Code source correspondant :**

```python
307  def extract_meals(path, root, sheet, cell, mapping, header_row, recap=False):
308      """Extract meal rows while preserving the food-cell text.
309  
310      Args:
311          path, root (Path): Workbook and input directory.
312          sheet (str): Worksheet name.
313          cell (callable): 1-based cell accessor.
314          mapping (dict): Food field to Excel column mapping.
315          header_row (int): Header row number.
316          recap (bool): Whether this is a RECAP worksheet.
317      Returns:
318          list[dict]: Extracted meal records with source metadata.
319      """
320      meals = []
321      for row in range(max(5, header_row + 1), 19):
322          # In some MENUS sheets, the day label occupies the entree column.
323          shifted = (not recap and 'entree' in mapping
324                     and is_day_label(cell(row, mapping['entree'])))
325          offset = 1 if shifted else 0
326          values = {}
327          for field in MENU_FIELDS:
328              column = mapping.get(field)
329              value = cell_text(cell(row, column + offset)) if column is not None else ''
330              # A weekday label is a structural marker, never a dish.
331              values[field] = '' if is_day_label(value) else value
332          # Preserve the existing output schema; child meals are not extracted.
333          values['enfant'] = ''
334          if not any(values[field] for field in MENU_FIELDS):
335              continue
336          meals.append({**source_info(path, root, sheet, row),
337                        'type_menu': 'principal', **values})
338      return meals
```


### 4.14 `extract_grid(path, root, sheet, cell)` — lignes 341–360

**Mission :** Extraire les propositions individuelles des grilles historiques 2022–2024.

- **`path`, `root`, `sheet`** identifient la source ; **`cell`** lit les cellules par coordonnées Excel.
- **`proposals = []`** accumule les propositions de la feuille.
- La boucle externe **`for column in GRID_COLUMNS`** visite successivement les **7 colonnes D, F, H, J, L, N et P** (numéros 4, 6, 8, 10, 12, 14, 16).
- Pour **chaque** colonne, la boucle interne **`for row, (category, subcategory) in GRID_ROWS.items()`** parcourt les **18 lignes configurées** : **7–18**, puis **21–26**. Les lignes 19 et 20 ne sont pas parcourues.
- Le programme examine donc **7 × 18 = 126 positions candidates** par feuille principale, avant filtrage.
- **`row`** est le numéro Excel ; **`category`** vaut par exemple `entree`, `plat`, `poisson`, `dessert` ; **`subcategory`** vaut par exemple `ENTREE 1`, `PLAT 2`, `POISSON`. Ces deux informations viennent directement de `GRID_ROWS`.
- **`value = cell_text(cell(row, column))`** lit la position courante. Si la cellule est vide **ou** contient uniquement un jour, elle est ignorée.
- Sinon, `proposals.append(...)` ajoute une ligne avec le fichier, l’onglet, **la ligne et la colonne Excel**, la catégorie, la sous-catégorie et le texte de la proposition.
- **Exemple :** si D7 contient `Carottes râpées`, le résultat porte `ligne_excel=7`, `colonne_excel=4`, `categorie='entree'`, `sous_categorie='ENTREE 1'`, `proposition='Carottes râpées'`.
- **Résultat :** une liste de propositions individuelles, **pas** de repas complets.


**Code source correspondant :**

```python
341  def extract_grid(path, root, sheet, cell):
342      """Extract individual proposals from historical fixed-layout worksheets.
343  
344      Args:
345          path, root (Path): Workbook and input directory.
346          sheet (str): Worksheet name.
347          cell (callable): 1-based cell accessor.
348      Returns:
349          list[dict]: Proposal records, retaining category and Excel position.
350      """
351      proposals = []
352      for column in GRID_COLUMNS:
353          for row, (category, subcategory) in GRID_ROWS.items():
354              value = cell_text(cell(row, column))
355              if not value or is_day_label(value):
356                  continue
357              proposals.append({**source_info(path, root, sheet, row),
358                                'colonne_excel': column, 'categorie': category,
359                                'sous_categorie': subcategory, 'proposition': value})
360      return proposals
```


### 4.15 `write_csv(path, columns, records)` — lignes 364–377

**Mission :** Écrire un ensemble de dictionnaires dans un fichier CSV.

- **`path`** est le chemin du CSV à créer ; **`columns`** définit l’ordre exact des colonnes ; **`records`** est la liste des dictionnaires à écrire.
- `path.open('w', newline='', encoding='utf-8-sig')` ouvre le fichier en écriture : **`'w'` remplace son contenu s’il existe**. La protection contre un écrasement involontaire est donc effectuée **avant** cet appel, dans `main()`.
- **`utf-8-sig`** ajoute une signature UTF-8 qui facilite la reconnaissance des accents par certains tableurs, notamment Excel.
- `csv.DictWriter(..., fieldnames=columns, delimiter=';')` associe les clés des dictionnaires aux colonnes et utilise le **point-virgule** comme séparateur.
- `writeheader()` écrit les noms de colonnes ; `writerows(records)` écrit ensuite toutes les données. La fonction ne renvoie pas de résultat utile.


**Code source correspondant :**

```python
364  def write_csv(path, columns, records):
365      """Write records as a semicolon-delimited UTF-8 CSV file.
366  
367      Args:
368          path (Path): Destination file; overwritten if it exists.
369          columns (sequence[str]): Ordered CSV column names.
370          records (list[dict]): Data rows.
371      Returns:
372          None.
373      """
374      with path.open('w', newline='', encoding='utf-8-sig') as output:
375          writer = csv.DictWriter(output, fieldnames=columns, delimiter=';')
376          writer.writeheader()
377          writer.writerows(records)
```


### 4.16 `main()` — lignes 380–519

**Mission :** Piloter toute l’extraction, traiter les anomalies et construire les deux CSV.

**Étape A — paramètres et protection des sorties.** `parser` définit `--input` (par défaut `data/raw`), `--output` (par défaut `data/processed`) et `--ecraser` (faux si l’option est absente). **`args`** contient les valeurs réellement reçues. Si `args.input` n’est pas un dossier, le programme s’arrête. **`outputs`** liste les deux noms de CSV ; **`existing`** retient ceux qui existent déjà. Si au moins un existe et que `args.ecraser` est faux, le programme s’arrête **avant toute écriture**.

**Étape B — découverte des fichiers.** **`files`** est la liste triée de tous les fichiers trouvés par `args.input.rglob('*')` ; seuls les fichiers d’extension `.xls` ou `.xlsx` sont retenus, à l’exclusion des fichiers temporaires dont le nom commence par `~$`. La boucle principale `for path in files` s’exécute **une fois par fichier trouvé** : le nombre d’itérations est `len(files)`.

**Étape C — accumulateurs.** **`meals`** contient les repas récents, **`proposals`** les propositions des grilles historiques, **`annexes`** les cellules d’annexes, **`alerts`** les diagnostics et **`counts`** les compteurs (`MENUS`, `RECAP`, `GRILLE`, `annex_sheets`, `unknown`, `errors`, selon les événements rencontrés). La fonction interne **`alert(relative, level, code, message, sheet='', total_sheets=None)`** ajoute un dictionnaire à `alerts` : elle **n’écrit pas immédiatement** le CSV.

**Étape D — traitement d’un fichier.** **`relative`** est son chemin relatif au dossier d’entrée. **`fmt, week`** proviennent de `expected_sheet()` ; **`name, cell, names, skipped, selection_warning`** de `read_selected_sheet()`. Si `name is None`, `counts['unknown']` augmente, une alerte **`FEUILLE_INTROUVABLE`** est ajoutée, et le programme passe **au fichier suivant**. Un `selection_warning` devient une alerte **`NOM_FEUILLE_DIFFERENT`**.

**Étape E — onglets supplémentaires.** **`year`** est relu depuis les dossiers du chemin. Pour **2022 et 2023 seulement**, la boucle `for extra_name in skipped` examine **chaque onglet non principal** : `annex_kind()` identifie `livraison` ou `median`. Les annexes reconnues sont lues avec `extract_annex()`, puis ajoutées à **`annexes`** via `extend()`. **`remaining`** conserve les onglets non reconnus ou dont la lecture a échoué. Une annexe vide produit **`ANNEXE_VIDE`**, une lecture ratée **`LECTURE_ANNEXE`**. Pour **2024–2026**, tous les onglets secondaires restent dans `remaining`. Si `remaining` n’est pas vide, une alerte **`FEUILLES_IGNOREES`** est ajoutée.

**Étape F — extraction principale.** Si **`fmt != 'GRILLE'`**, `find_header(cell)` produit **`mapping`** et **`header`** ; si moins de trois en-têtes obligatoires sont trouvés, une alerte **`ENTETES_INVALIDES`** est ajoutée et le programme passe au fichier suivant. Si `fmt == 'GRILLE'`, **`extracted`** provient de `extract_grid()` et est ajouté à **`proposals`** ; sinon, il provient de `extract_meals()` et est ajouté à **`meals`**. L’argument `recap` vaut vrai **uniquement lorsque le nom réel de l’onglet choisi est `Recap`**, pas simplement parce que l’année vaut 2025. `counts[fmt]` augmente après l’extraction. Si `extracted` est vide, le programme ajoute **`AUCUNE_DONNEE`**.

**Étape G — erreurs non prévues.** Le `except Exception as exc` autour du traitement du fichier augmente **`counts['errors']`** et ajoute une alerte **`EXCEPTION`**, sans arrêter le parcours des autres fichiers.

**Étape H — fusion et dédoublonnage.** **`columns`** construit l’en-tête commun du CSV final. **`records`** regroupe les trois listes en ajoutant un **`type_enregistrement`** : `'repas'`, `'plat_historique'` ou `'cellule_annexe'`. **`unique_records`** est un dictionnaire dont la clé est le quintuplet **`(chemin_source, onglet, ligne_excel, colonne_excel, type_enregistrement)`**. `setdefault(key, record)` conserve la première occurrence de chaque clé. **Le texte du plat n’entre pas dans la clé** : deux plats identiques situés à deux coordonnées différentes restent tous les deux présents.

**Étape I — écriture et bilan.** `args.output.mkdir(...)` crée le dossier cible si nécessaire. Le programme écrit **`historique_menus.csv`** avec `records`, puis **`alert.csv`** avec `alerts`. **`successful`** additionne les fichiers principaux traités dans les trois formats ; **`error_count`** et **`warning_count`** comptent les lignes d’alerte par niveau. Le `print(...)` final affiche le nombre de fichiers trouvés, traités, non traités, les erreurs, avertissements, repas, propositions et cellules annexes.


**Code source — lignes 388–420:**

```python
388      parser = argparse.ArgumentParser(description=__doc__)
389      parser.add_argument('--input', type=Path, default=Path('data/raw'))
390      parser.add_argument('--output', type=Path, default=Path('data/processed'))
391      parser.add_argument('--ecraser', action='store_true',
392                          help='Allow overwriting previously generated extraction CSV files')
393      args = parser.parse_args()
394      if not args.input.is_dir():
395          parser.error(f'Input directory not found: {args.input}')
396  
397      # Refuse overwriting either generated CSV unless explicitly requested.
398      outputs = ('historique_menus.csv', 'alert.csv')
399      existing = [name for name in outputs if (args.output / name).exists()]
400      if existing and not args.ecraser:
401          parser.error('Output already exists: ' + ', '.join(existing)
402                       + '. Choose another --output or use --ecraser.')
403  
404      files = sorted(path for path in args.input.rglob('*')
405                     if path.is_file() and path.suffix.lower() in ('.xls', '.xlsx')
406                     and not path.name.startswith('~$'))
407      meals, proposals, annexes, alerts, counts = [], [], [], [], Counter()
408  
409      def alert(relative, level, code, message, sheet='', total_sheets=None):
410          """Collect a warning or error without printing it to the terminal."""
411          alerts.append({
412              'fichier_source': relative,
413              'niveau': level,
414              'code': code,
415              'onglet': sheet,
416              'nombre_feuilles': '' if total_sheets is None else total_sheets,
417              'message': message,
418          })
419  
420      for path in files:
```


**Code source — lignes 420–490:**

```python
420      for path in files:
421          relative = path.relative_to(args.input).as_posix()
422          try:
423              fmt, week = expected_sheet(path, args.input)
424              name, cell, names, skipped, selection_warning = read_selected_sheet(path, fmt, week)
425              if name is None:
426                  counts['unknown'] += 1
427                  expected = ('Menus' if fmt == 'MENUS' else 'Recap/Menus' if fmt == 'RECAP'
428                              else str(week))
429                  alert(relative, 'ERREUR', 'FEUILLE_INTROUVABLE',
430                        f"Feuille attendue : {expected!r}. Feuilles disponibles : {names!r}",
431                        total_sheets=len(names))
432                  continue
433              if selection_warning:
434                  alert(relative, 'AVERTISSEMENT', 'NOM_FEUILLE_DIFFERENT',
435                        selection_warning, sheet=name, total_sheets=len(names))
436              # 2022-2023: read delivery and Median tabs as additional raw cells.
437              # 2024-2026: preserve the original one-sheet-only behavior.
438              year = int(next(part for part in path.relative_to(args.input).parts[:-1]
439                              if re.fullmatch(r"20\d{2}", part)))
440              remaining = []
441              if year in (2022, 2023):
442                  for extra_name in skipped:
443                      kind = annex_kind(extra_name)
444                      if kind is None:
445                          remaining.append(extra_name)
446                          continue
447                      try:
448                          extracted_extra = extract_annex(path, args.input, extra_name, kind)
449                          annexes.extend(extracted_extra)
450                          counts['annex_sheets'] += 1
451                          if not extracted_extra:
452                              alert(relative, 'AVERTISSEMENT', 'ANNEXE_VIDE',
453                                    f"Aucune cellule non vide dans {extra_name!r}",
454                                    sheet=extra_name, total_sheets=len(names))
455                      except Exception as exc:
456                          remaining.append(extra_name)
457                          alert(relative, 'ERREUR', 'LECTURE_ANNEXE',
458                                f"{type(exc).__name__}: {exc}",
459                                sheet=extra_name, total_sheets=len(names))
460              else:
461                  remaining = skipped
462              if remaining:
463                  alert(relative, 'AVERTISSEMENT', 'FEUILLES_IGNOREES',
464                        f"1 feuille principale traitee ({name!r}); "
465                        f"{len(remaining)} ignoree(s) : {remaining!r}",
466                        sheet=name, total_sheets=len(names))
467              mapping, header = (find_header(cell) if fmt != 'GRILLE' else (None, None))
468              if fmt != 'GRILLE' and mapping is None:
469                  counts['unknown'] += 1
470                  alert(relative, 'ERREUR', 'ENTETES_INVALIDES',
471                        f"En-tetes non reconnus dans {name!r}",
472                        sheet=name, total_sheets=len(names))
473                  continue
474              if fmt == 'GRILLE':
475                  extracted = extract_grid(path, args.input, name, cell)
476                  proposals.extend(extracted)
477              else:
478                  extracted = extract_meals(path, args.input, name, cell, mapping,
479                                            header, recap=(match_text(name).strip() == 'recap'))
480                  meals.extend(extracted)
481              counts[fmt] += 1
482              if not extracted:
483                  alert(relative, 'AVERTISSEMENT', 'AUCUNE_DONNEE',
484                        f"Aucune ligne extraite (format {fmt})",
485                        sheet=name, total_sheets=len(names))
486          except Exception as exc:
487              counts['errors'] += 1
488              alert(relative, 'ERREUR', 'EXCEPTION',
489                    f'{type(exc).__name__}: {exc}')
490  
```


**Code source — lignes 491–523:**

```python
491      columns = ('type_enregistrement', *MENU_COLUMNS,
492                 'colonne_excel', 'categorie', 'sous_categorie', 'proposition', 'type_feuille')
493      records = ([{'type_enregistrement': 'repas', **row} for row in meals]
494                 + [{'type_enregistrement': 'plat_historique', 'type_feuille': 'principale', **row}
495                    for row in proposals]
496                 + [{'type_enregistrement': 'cellule_annexe', **row} for row in annexes])
497      # Coordinates identify a unique source cell, even if its text occurs elsewhere.
498      unique_records = {}
499      for record in records:
500          key = (record.get('chemin_source'), record.get('onglet'),
501                 record.get('ligne_excel'), record.get('colonne_excel'),
502                 record.get('type_enregistrement'))
503          unique_records.setdefault(key, record)
504      records = list(unique_records.values())
505      args.output.mkdir(parents=True, exist_ok=True)
506      write_csv(args.output / 'historique_menus.csv', columns, records)
507      write_csv(args.output / 'alert.csv',
508                ('fichier_source', 'niveau', 'code', 'onglet', 'nombre_feuilles', 'message'),
509                alerts)
510  
511      successful = counts['MENUS'] + counts['RECAP'] + counts['GRILLE']
512      error_count = sum(item['niveau'] == 'ERREUR' for item in alerts)
513      warning_count = sum(item['niveau'] == 'AVERTISSEMENT' for item in alerts)
514      print(f"Extraction terminee | fichiers : {len(files)} | traites : {successful} "
515            f"| non traites : {len(files) - successful} | erreurs : {error_count} "
516            f"| avertissements : {warning_count} | repas : {len(meals)} "
517            f"| propositions : {len(proposals)} | feuilles annexes : {counts['annex_sheets']} "
518            f"| cellules annexes : {len(annexes)} | sorties : "
519            f"{args.output / 'historique_menus.csv'}, {args.output / 'alert.csv'}")
520  
521  
522  if __name__ == '__main__':
523      main()
```


## 5. Suivi concret d’une donnée : trois exemples

### Exemple 1 — un fichier de grille 2023

Pour `data/raw/2023/Menu sem 15.xlsx`, `expected_sheet()` renvoie `('GRILLE', 15)`. `choose_sheet()` cherche une feuille correspondant à la semaine 15. `extract_grid()` lit les **126 positions** définies par `GRID_COLUMNS × GRID_ROWS`, conserve les cellules alimentaires non vides et crée un enregistrement par position. Si D7 contient « Carottes râpées », la proposition garde **D7**, la catégorie `entree` et la sous-catégorie `ENTREE 1`. Elle est fusionnée dans `records` avec le type `plat_historique`.

### Exemple 2 — un fichier `Menus` de 2026

`expected_sheet()` renvoie `('MENUS', None)`. La feuille `Menus` est choisie. `find_header()` cherche les intitulés sur les lignes **3–6** et colonnes **A–Q**. Si les en-têtes sont détectés ligne 4, `extract_meals()` examine les lignes **5–18** ; pour chacune, il lit les six champs selon `mapping`, avec un éventuel décalage de colonne si une étiquette de jour occupe la colonne d’entrée. Une ligne sans aucun champ alimentaire est ignorée ; les autres deviennent des enregistrements `repas`.

### Exemple 3 — un onglet `Médian` de 2022

Si un classeur 2022 contient, en plus de sa feuille principale, un onglet `Médian`, `annex_kind()` le classe `median`. `read_annex_cells()` visite **toutes les cellules non vides** de cet onglet. `extract_annex()` conserve leurs coordonnées et leur texte. Ces cellules sont exportées avec `type_enregistrement = 'cellule_annexe'` et `sous_categorie = 'cellule_brute'` : aucune interprétation automatique en entrée/plat/dessert n’est faite.

## 6. Ce que le programme **ne fait pas**

- Il ne corrige pas les intitulés culinaires, hormis le nettoyage des espaces.
- Il ne reconstruit pas des repas complets à partir des grilles historiques : il enregistre des propositions individuelles.
- Il ne classe pas sémantiquement les cellules des annexes `livraison` et `Médian`.
- Il ne parcourt pas au-delà des **65 premières lignes et 22 premières colonnes** pour les feuilles principales, ni au-delà de la **ligne 18** pour les repas récents.
- Il ne supprime pas les doublons de texte situés à des positions distinctes.
- Il ne modifie jamais les fichiers Excel originaux.
- Il n'écrit pas les CSV existants sans l'option `--ecraser`.

## 7. Commandes d’exécution

```bash
python -m pip install openpyxl xlrd
python extract_avec_alertes_multi_2022_2023.py --input data/raw --output data/test_extraction
python extract_avec_alertes_multi_2022_2023.py --input data/raw --output data/processed --ecraser
```

**Conseil de lecture :** commence par `main()` pour suivre le parcours complet, puis consulte les fonctions appelées au moment où tu veux comprendre précisément une décision ou une cellule lue.