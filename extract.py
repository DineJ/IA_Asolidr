#!/usr/bin/env python3
"""Extraire les menus Excel de 2022 à 2026 sans corriger les noms des plats.

Utilisation : python extract.py --input data/raw --output data/processed
Dépendances : openpyxl (.xlsx) et xlrd (.xls).

Sorties : historique_menus_YYYY.csv (un fichier par année), alert.csv
(et annexes.csv si demandé). Les fichiers existants ne sont écrasés
qu’avec --ecraser ; les classeurs Excel originaux restent intacts.

Par défaut, les annexes Livraison/Médian ne sont pas extraites.
Utiliser --inclure-annexes pour exporter leurs cellules brutes dans annexes.csv (2022-2023).
"""

# Options de la commande : dossiers, annexes facultatives et autorisation d'ecrasement.
import argparse
# Ecriture des CSV avec separateur point-virgule et encodage compatible Excel.
import csv
# Reconnaissance des annees, semaines, noms d'onglets et jours.
import re
# Comparaisons sans accents, sans modifier les noms des plats exportes.
import unicodedata
# Compteurs pour le bilan final de l'extraction.
from collections import Counter
# Gestion des chemins et parcours recursif des fichiers.
from pathlib import Path

# Lecture des fichiers .xlsx ; les .xls utilisent xlrd dans les fonctions.
from openpyxl import load_workbook

# Champs alimentaires des menus 2025-2026 : une ligne CSV par repas.
MENU_FIELDS = ('entree', 'proteine', 'feculent', 'legumes', 'dessert_base', 'dessert_resto')
# Deux formats de sortie : grilles historiques et repas récents.
# Les coordonnées Excel sont conservées pour faciliter les vérifications.
GRID_OUTPUT_COLUMNS = ('chemin_source', 'ligne_excel', 'colonne_excel',
                       'categorie', 'sous_categorie', 'proposition')
MEAL_OUTPUT_COLUMNS = ('chemin_source', 'ligne_excel', *MENU_FIELDS)
ANNEX_OUTPUT_COLUMNS = ('chemin_source', 'onglet', 'ligne_excel',
                        'colonne_excel', 'categorie', 'sous_categorie', 'proposition')

# GRILLE (2022-2024) : la catégorie est déduite de la LIGNE Excel.
# Important : les lignes 19 et 20 sont bien des laitages ; ne pas les omettre.
# Ces coordonnées sont propres au modèle historique, pas aux menus récents.
GRID_ROWS = {
    7: ('entree', 'ENTREE 1'), 8: ('entree', 'ENTREE 2'),
    9: ('entree', 'ENTREE 3'), 10: ('entree', 'ENTREE 4'),
    11: ('plat', 'PLAT 1'), 12: ('plat', 'PLAT 2'),
    13: ('plat', 'ASE SOIR'), 14: ('poisson', 'POISSON'),
    15: ('legume', 'LEGUME 1'), 16: ('legume', 'LEGUME 2'),
    17: ('feculent', 'FECULENTS'), 18: ('laitage', 'LAITAGE 1'),
    19: ('laitage', 'LAITAGE 2'), 20: ('laitage', 'LAITAGE 3'),
    21: ('fromage', 'FROMAGE 1'), 22: ('fromage', 'FROMAGE 2'),
    23: ('dessert', 'DESSERT 1'), 24: ('fruit', 'FRUIT 1'),
    25: ('dessert', 'DESSERT 2'), 26: ('fruit', 'FRUIT 2'),
}
# Sept colonnes de propositions de la grille : D, F, H, J, L, N et P.
GRID_COLUMNS = (4, 6, 8, 10, 12, 14, 16)
# Trois de ces quatre en-tetes suffisent a reconnaitre un menu recent.
EXPECTED_HEADERS = ('entree', 'proteine', 'feculent', 'legumes')
# Etiquettes de jour seules, avec ou sans MIDI/SOIR : ne sont pas des aliments.
DAY_LABEL = re.compile(r'(?:LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)'r'(?:\s+(?:MIDI|SOIR))?', re.IGNORECASE)


# Nettoyage des espaces uniquement : aucune correction orthographique ou culinaire.
def cell_text(value):
    """Return a readable cell string without changing its culinary meaning.

    Args:
        value: Raw Excel cell value, possibly None, a number or a string.
    Returns:
        str: Original text with whitespace standardized, or '' for empty cells.
    """
    if value is None:
        return ''
    # Boolean cells are metadata rather than food descriptions.
    if isinstance(value, bool):
        return ''
    return ' '.join(str(value).replace('\u00a0', ' ').split())


# Cette normalisation sert a reperer la structure, jamais a modifier les plats.
def match_text(value):
    """Normalize text only for detecting workbook structure, not CSV values.

    Args:
        value: Cell value or worksheet name.
    Returns:
        str: Accent-insensitive, case-folded text.
    """
    normalized = unicodedata.normalize('NFKD', cell_text(value).casefold())
    return ''.join(char for char in normalized if not unicodedata.combining(char))


# Un jour seul (ex. mercredi soir) doit etre ignore, sans decaler les colonnes.
def is_day_label(value):
    """Detect a weekday or weekday/service label in a structural cell.

    Args:
        value: Cell content.
    Returns:
        bool: Whether the entire value is a recognized day label.
    """
    return bool(DAY_LABEL.fullmatch(match_text(value).strip()))


# Le dossier annee choisit le format ; la semaine sert aux onglets 2022-2024.
def expected_sheet(path, root):
    """Return the exact expected sheet name type and week for a workbook.

    Args:
        path (Path): Source Excel file.
        root (Path): Root directory containing year folders.
    Returns:
        tuple[str, int | None]: Expected format and week number.
    Raises:
        ValueError: If the year or historical week cannot be identified.
    """
    relative_parts = path.relative_to(root).parts
    years = [int(part) for part in relative_parts[:-1]
             if re.fullmatch(r"20\d{2}", part)]
    if not years:
        raise ValueError("Year folder missing (expected data/raw/YYYY/)")
    year = years[-1]
    if year == 2026:
        return 'MENUS', None
    if year == 2025:
        return 'RECAP', None
    if 2022 <= year <= 2024:
        match = re.search(r"(?:sem(?:aine)?[\s_.-]*)(\d{1,2})(?!\d)",
                          match_text(path.stem))
        if not match:
            raise ValueError("Week number not found in file name")
        return 'GRILLE', int(match.group(1))
    raise ValueError(f"Unsupported year: {year}")


# Check whether a worksheet name matches the expected format.
def sheet_matches(name, fmt, week):
    """Check whether a sheet is suitable for the selected year's layout."""
    normalized = match_text(name).strip()
    if fmt == 'MENUS':
        return normalized == 'menus'
    if fmt == 'RECAP':
        return normalized in ('recap', 'menus')
    # Accept week labels with an optional 'midi' suffix.
    return bool(re.fullmatch(
        r"(?:sem(?:aine)?[\s_.-]*)?0*" + str(week) + r"(?:[\s_.-]*midi)?",
        normalized,
    ))


# Une seule feuille principale par classeur ; priorite aux onglets attendus.
def choose_sheet(names, fmt, week):
    """Choose exactly one sheet; never guess among ambiguous multi-sheet grids.

    Returns:
        tuple[str | None, str | None]: Selected sheet and optional warning.
    """
    if fmt == 'RECAP':
        # Older 2025 workbooks use RECAP; later ones use MENUS.
        for preferred in ('recap', 'menus'):
            for name in names:
                if match_text(name).strip() == preferred:
                    return name, None
        return None, None
    if fmt == 'MENUS':
        return next((name for name in names if sheet_matches(name, fmt, week)), None), None

    matches = [name for name in names if sheet_matches(name, fmt, week)]
    if matches:
        # Prefer the recap grid over other sheets of the same week.
        matches.sort(key=lambda name: ('midi' not in match_text(name), names.index(name)))
        return matches[0], None
    midi = [name for name in names if re.search(r'(?<![a-z])midi(?![a-z])', match_text(name))]
    if len(midi) == 1:
        return midi[0], (f"sheet name {midi[0]!r} does not match week {week}; "
                         "using the only midi sheet")
    if len(names) == 1:
        # Many historical workbooks were copied without renaming their only tab.
        return names[0], (f"sheet name {names[0]!r} does not match week {week}; "
                          "using the only sheet")
    return None, None


# Charger seulement A1:V65 pour la feuille principale, puis fermer le classeur.
def read_selected_sheet(path, fmt, week):
    """Read only the expected worksheet and report the other worksheet names.

    Args:
        path (Path): Excel workbook.
        fmt (str): MENUS, RECAP or GRILLE.
        week (int | None): Week for historical grids.
    Returns:
        tuple[str | None, callable | None, list[str], list[str], str | None]:
            Selected sheet, accessor, all names, skipped names and selection warning.
    """
    if path.suffix.lower() == '.xls':
        import xlrd
        workbook = xlrd.open_workbook(str(path), on_demand=True)
        try:
            names = workbook.sheet_names()
            selected, selection_warning = choose_sheet(names, fmt, week)
            skipped = [name for name in names if name != selected]
            if selected is None:
                return None, None, names, skipped, None
            sheet = workbook.sheet_by_name(selected)
            # Cache only the required region before closing the workbook.
            rows = [[sheet.cell_value(r, c) if r < sheet.nrows and c < sheet.ncols
                     else None for c in range(22)] for r in range(65)]
        finally:
            workbook.release_resources()
    else:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            names = workbook.sheetnames
            selected, selection_warning = choose_sheet(names, fmt, week)
            skipped = [name for name in names if name != selected]
            if selected is None:
                return None, None, names, skipped, None
            sheet = workbook[selected]
            sheet.reset_dimensions()
            rows = list(sheet.iter_rows(min_row=1, max_row=65, min_col=1,
                                        max_col=22, values_only=True))
        finally:
            workbook.close()

    def cell(row, column, data=rows):
        """Return a 1-based cell value from the selected worksheet only."""
        if 1 <= row <= len(data) and 1 <= column <= len(data[row - 1]):
            return data[row - 1][column - 1]
        return None

    return selected, cell, names, skipped, selection_warning



# Annexes facultatives : lire toutes les cellules, meme hors de A1:V65.
def read_annex_cells(path, sheet_name):
    """Read populated cells of a 2022-2023 annex, without assuming a menu layout.

    The returned positions are the real 1-based Excel coordinates. Formula cells
    are read with their cached values, as in the original extraction.
    """
    if path.suffix.lower() == '.xls':
        import xlrd
        workbook = xlrd.open_workbook(str(path), on_demand=True)
        try:
            sheet = workbook.sheet_by_name(sheet_name)
            return [(r + 1, c + 1, cell_text(sheet.cell_value(r, c)))
                    for r in range(sheet.nrows) for c in range(sheet.ncols)
                    if cell_text(sheet.cell_value(r, c))]
        finally:
            workbook.release_resources()
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[sheet_name]
        sheet.reset_dimensions()
        return [(r, c, cell_text(value))
                for r, values in enumerate(sheet.iter_rows(values_only=True), 1)
                for c, value in enumerate(values, 1) if cell_text(value)]
    finally:
        workbook.close()


# Seuls les onglets Livraison ou Median sont des annexes reconnues.
def annex_kind(sheet_name):
    """Only the two explicitly requested historical annex types qualify."""
    normalized = match_text(sheet_name)
    if 'livraison' in normalized:
        return 'livraison'
    if 'median' in normalized:
        return 'median'
    return None


# Les cellules annexes restent brutes et sont exportees a part, jamais en plats.
def extract_annex(path, root, sheet_name, kind):
    """Store raw nonempty cells; do not mislabel delivery metadata as dishes."""
    return [{**source_info(path, root, sheet_name, row),
             'colonne_excel': col, 'categorie': kind,
             'sous_categorie': 'cellule_brute', 'proposition': value,
             'type_feuille': kind}
            for row, col, value in read_annex_cells(path, sheet_name)]


# Chercher les en-tetes dans A3:Q6 et garder la meilleure ligne.
def find_header(cell):
    """Repérer les colonnes par leurs en-têtes (et non par position supposée).

    Args:
        cell (callable): 1-based Excel cell accessor.
    Returns:
        tuple[dict | None, int | None]: Column mapping and header row.
    """
    best_mapping, best_row, best_score = None, None, 0
    for row in range(3, 7):
        found = {}
        for column in range(1, 18):
            heading = match_text(cell(row, column))
            if 'entree' in heading:
                found.setdefault('entree', column)
            elif 'proteine' in heading or 'plat principal' in heading:
                found.setdefault('proteine', column)
            elif 'feculent' in heading:
                found.setdefault('feculent', column)
            elif 'legume' in heading:
                found.setdefault('legumes', column)
            elif 'dessert' in heading and ('resto' in heading or 'restaurant' in heading):
                found.setdefault('dessert_resto', column)
            elif 'dessert' in heading:
                found.setdefault('dessert_base', column)
        score = sum(field in found for field in EXPECTED_HEADERS)
        if score > best_score:
            best_mapping, best_row, best_score = found, row, score
    if best_score < 3:
        return None, None
    return best_mapping, best_row


# La provenance complete sert au dedoublonnage et au diagnostic interne.
def source_info(path, root, sheet, row):
    """Build provenance fields for an extracted cell or meal.

    Args:
        path (Path): Workbook path.
        root (Path): Input directory.
        sheet (str): Worksheet name.
        row (int): Excel row number.
    Returns:
        dict: Source file, relative path, sheet and row.
    """
    return dict(fichier_source=path.name,
                chemin_source=path.relative_to(root).as_posix(),
                onglet=sheet, ligne_excel=row)


# Menus recents : lire chaque champ dans la colonne de son en-tete.
def extract_meals(path, root, sheet, cell, mapping, header_row):
    """Extraire les repas récents sans décaler les colonnes selon le jour.

    Args:
        path, root (Path): Workbook and input directory.
        sheet (str): Worksheet name.
        cell (callable): 1-based cell accessor.
        mapping (dict): Food field to Excel column mapping.
        header_row (int): Header row number.
    Returns:
        list[dict]: Extracted meal records with source metadata.
    """
    meals = []
    for row in range(max(5, header_row + 1), 19):
        # CORRECTION IMPORTANTE : mapping reste FIXE sur toutes les lignes.
        # Un jour (ex. « mercredi soir ») peut figurer dans une cellule
        # d’entrée, mais cela ne déplace PAS les colonnes des aliments.
        # Un décalage +1 classait à tort « RIZ » comme protéine en 2026.
        values = {}
        for field in MENU_FIELDS:
            column = mapping.get(field)
            value = cell_text(cell(row, column)) if column is not None else ''
            # Une etiquette de jour devient une valeur vide ; aucun decalage +1.
            values[field] = '' if is_day_label(value) else value
        if not any(values[field] for field in MENU_FIELDS):
            continue
        meals.append({**source_info(path, root, sheet, row),
                      **values})
    return meals


# Grilles 2022-2024 : chaque cellule candidate est une proposition distincte.
def extract_grid(path, root, sheet, cell):
    """Extract individual proposals from historical fixed-layout worksheets.

    Args:
        path, root (Path): Workbook and input directory.
        sheet (str): Worksheet name.
        cell (callable): 1-based cell accessor.
    Returns:
        list[dict]: Proposal records, retaining category and Excel position.
    """
    proposals = []
    for column in GRID_COLUMNS:
        for row, (category, subcategory) in GRID_ROWS.items():
            # Row 13 can be named differently depending on the source workbook.
            if row == 13 and cell_text(cell(13, 3)):
                subcategory = cell_text(cell(13, 3))
            value = cell_text(cell(row, column))
            if not value or is_day_label(value):
                continue
            proposals.append({**source_info(path, root, sheet, row),
                              'colonne_excel': column, 'categorie': category,
                              'sous_categorie': subcategory, 'proposition': value})
    return proposals



# CSV separes par ; et encodes en UTF-8 avec BOM pour Excel.
def write_csv(path, columns, records):
    """Write records as a semicolon-delimited UTF-8 CSV file.

    Args:
        path (Path): Destination file; overwritten if it exists.
        columns (sequence[str]): Ordered CSV column names.
        records (list[dict]): Data rows.
    Returns:
        None.
    """
    with path.open('w', newline='', encoding='utf-8-sig') as output:
        writer = csv.DictWriter(output, fieldnames=columns, delimiter=';')
        writer.writeheader()
        writer.writerows(records)


# Pilotage : decouverte, extraction, alertes, dedoublonnage, exports annuels.
def main():
    """Extraire une feuille principale par classeur vers des CSV annuels.

    Args:
        None. CLI options: --input and --output.
    Returns:
        None.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, default=Path('data/raw'))
    parser.add_argument('--output', type=Path, default=Path('data/processed'))
    parser.add_argument('--inclure-annexes', action='store_true',
                        help='Inclure les cellules brutes des annexes livraison/Median')
    parser.add_argument('--ecraser', action='store_true',
                        help='Allow overwriting previously generated extraction CSV files')
    args = parser.parse_args()
    if not args.input.is_dir():
        parser.error(f'Input directory not found: {args.input}')

    # Protection : ne pas ecraser les historiques ou alertes sans --ecraser.
    outputs = ('alert.csv', 'annexes.csv')
    existing = [name for name in outputs if (args.output / name).exists()]
    existing += [path.name for path in args.output.glob('historique_menus_*.csv')] if args.output.exists() else []
    if existing and not args.ecraser:
        parser.error('Output already exists: ' + ', '.join(existing)
                     + '. Choose another --output or use --ecraser.')

    files = sorted(path for path in args.input.rglob('*')
                   if path.is_file() and path.suffix.lower() in ('.xls', '.xlsx')
                   and not path.name.startswith('~$'))
    meals, proposals, annexes, alerts, counts = [], [], [], [], Counter()

    # Alertes techniques : feuille absente, en-têtes inconnus, exception, etc.
    # Ce mécanisme ne détecte pas les fautes de frappe des plats.
    def alert(relative, level, code, message, sheet='', total_sheets=None):
        """Collect a warning or error without printing it to the terminal."""
        alerts.append({
            'fichier_source': relative,
            'niveau': level,
            'code': code,
            'onglet': sheet,
            'nombre_feuilles': '' if total_sheets is None else total_sheets,
            'message': message,
        })

    # Une erreur sur un classeur n'interrompt pas le traitement des suivants.
    for path in files:
        relative = path.relative_to(args.input).as_posix()
        try:
            fmt, week = expected_sheet(path, args.input)
            name, cell, names, skipped, selection_warning = read_selected_sheet(path, fmt, week)
            if name is None:
                counts['unknown'] += 1
                expected = ('Menus' if fmt == 'MENUS' else 'Recap/Menus' if fmt == 'RECAP'
                            else str(week))
                alert(relative, 'ERREUR', 'FEUILLE_INTROUVABLE',
                      f"Feuille attendue : {expected!r}. Feuilles disponibles : {names!r}",
                      total_sheets=len(names))
                continue
            if selection_warning:
                alert(relative, 'AVERTISSEMENT', 'NOM_FEUILLE_DIFFERENT',
                      selection_warning, sheet=name, total_sheets=len(names))
            # Annexes desactivees par defaut : elles contiennent parfois des
            # donnees logistiques, non alimentaires. Seules celles de 2022-2023
            # peuvent etre exportees via --inclure-annexes.
            year = int(next(part for part in path.relative_to(args.input).parts[:-1]
                            if re.fullmatch(r"20\d{2}", part)))
            remaining = []
            if year in (2022, 2023) and args.inclure_annexes:
                for extra_name in skipped:
                    kind = annex_kind(extra_name)
                    if kind is None:
                        remaining.append(extra_name)
                        continue
                    try:
                        extracted_extra = extract_annex(path, args.input, extra_name, kind)
                        annexes.extend(extracted_extra)
                        counts['annex_sheets'] += 1
                        if not extracted_extra:
                            alert(relative, 'AVERTISSEMENT', 'ANNEXE_VIDE',
                                  f"Aucune cellule non vide dans {extra_name!r}",
                                  sheet=extra_name, total_sheets=len(names))
                    except Exception as exc:
                        remaining.append(extra_name)
                        alert(relative, 'ERREUR', 'LECTURE_ANNEXE',
                              f"{type(exc).__name__}: {exc}",
                              sheet=extra_name, total_sheets=len(names))
            else:
                remaining = skipped
            # Les annexes reconnues ne sont pas signalees comme feuilles ignorees.
            remaining = [n for n in remaining if annex_kind(n) is None]
            if remaining:
                alert(relative, 'AVERTISSEMENT', 'FEUILLES_IGNOREES',
                      f"1 feuille principale traitee ({name!r}); "
                      f"{len(remaining)} ignoree(s) : {remaining!r}",
                      sheet=name, total_sheets=len(names))
            mapping, header = (find_header(cell) if fmt != 'GRILLE' else (None, None))
            if fmt != 'GRILLE' and mapping is None:
                counts['unknown'] += 1
                alert(relative, 'ERREUR', 'ENTETES_INVALIDES',
                      f"En-tetes non reconnus dans {name!r}",
                      sheet=name, total_sheets=len(names))
                continue
            if fmt == 'GRILLE':
                extracted = extract_grid(path, args.input, name, cell)
                proposals.extend(extracted)
            else:
                extracted = extract_meals(path, args.input, name, cell, mapping, header)
                meals.extend(extracted)
            counts[fmt] += 1
            if not extracted:
                alert(relative, 'AVERTISSEMENT', 'AUCUNE_DONNEE',
                      f"Aucune ligne extraite (format {fmt})",
                      sheet=name, total_sheets=len(names))
        except Exception as exc:
            counts['errors'] += 1
            alert(relative, 'ERREUR', 'EXCEPTION',
                  f'{type(exc).__name__}: {exc}')

    # Les champs de provenance internes restent disponibles jusqu'a la fin.
    # Chaque format possède ses propres colonnes : pas de colonnes vides
    # type_enregistrement, type_menu, enfant ou type_feuille dans les historiques.
    # On dédoublonne par emplacement Excel, jamais par intitulé de plat.
    def deduplicate(items, with_column):
        unique = {}
        for item in items:
            key = (item['chemin_source'], item['onglet'], item['ligne_excel'])
            if with_column:
                key += (item['colonne_excel'],)
            unique.setdefault(key, item)
        return list(unique.values())

    meals = deduplicate(meals, with_column=False)
    proposals = deduplicate(proposals, with_column=True)
    annexes = deduplicate(annexes, with_column=True)

    # L'annee provient du dossier source et non d'une date dans le menu.
    def source_year(item):
        years = [part for part in Path(item['chemin_source']).parts[:-1]
                 if re.fullmatch(r'20\d{2}', part)]
        return years[-1] if years else None

    by_year = {}
    for item in proposals:
        year = source_year(item)
        if year:
            by_year.setdefault(year, []).append(item)
        else:
            alert(item['chemin_source'], 'ERREUR', 'ANNEE_INTROUVABLE',
                  'Dossier annee introuvable')
    for item in meals:
        year = source_year(item)
        if year:
            by_year.setdefault(year, []).append(item)
        else:
            alert(item['chemin_source'], 'ERREUR', 'ANNEE_INTROUVABLE',
                  'Dossier annee introuvable')

    args.output.mkdir(parents=True, exist_ok=True)
    generated = []
    for year, year_records in sorted(by_year.items()):
        # 2022-2024 : une ligne = une proposition dans une cellule Excel.
        # 2025-2026 : une ligne = un repas avec six champs alimentaires.
        columns = GRID_OUTPUT_COLUMNS if int(year) <= 2024 else MEAL_OUTPUT_COLUMNS
        destination = args.output / f'historique_menus_{year}.csv'
        write_csv(destination, columns,
                  [{key: item.get(key, '') for key in columns}
                   for item in year_records])
        generated.append(destination.name)

    if args.inclure_annexes:
        write_csv(args.output / 'annexes.csv', ANNEX_OUTPUT_COLUMNS,
                  [{key: item.get(key, '') for key in ANNEX_OUTPUT_COLUMNS}
                   for item in annexes])
        generated.append('annexes.csv')

    # Un seul fichier d'alertes : filtrer niveau=ERREUR pour les erreurs.
    # Aucun erreurs_par_fichier.csv n'est genere.
    alert_columns = ('fichier_source', 'niveau', 'code', 'onglet',
                     'nombre_feuilles', 'message')
    write_csv(args.output / 'alert.csv', alert_columns, alerts)

    # Calculate final summary counts and print one terminal line.
    successful = counts['MENUS'] + counts['RECAP'] + counts['GRILLE']
    error_count = sum(item['niveau'] == 'ERREUR' for item in alerts)
    warning_count = sum(item['niveau'] == 'AVERTISSEMENT' for item in alerts)
    print(f"Extraction terminee | fichiers : {len(files)} | traites : {successful} "
          f"| non traites : {len(files) - successful} | erreurs : {error_count} "
          f"| avertissements : {warning_count} | repas : {len(meals)} "
          f"| propositions : {len(proposals)} | feuilles annexes : {counts['annex_sheets']} "
          f"| cellules annexes : {len(annexes)} | sorties : "
          f"{', '.join(generated)}, alert.csv")


# Evite de lancer une extraction si le module est simplement importe.
if __name__ == '__main__':
    main()