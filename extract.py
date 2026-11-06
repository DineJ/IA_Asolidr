#!/usr/bin/env python3
"""Extract Excel menu data without normalizing dish names.

Usage: python extract.py --input data/raw --output data/processed
Dependencies: openpyxl and xlrd (for legacy .xls files).

The two output CSV files are overwritten only with --ecraser. Original
Excel workbooks are never modified.
"""

# Parse command-line options such as --input, --output and --ecraser.
import argparse
# Read and write CSV files using the Python standard library.
import csv
# Use regular expressions to identify years, weeks and worksheet labels.
import re
# Normalize Unicode characters when comparing labels.
import unicodedata
# Count processed formats and annex sheets.
from collections import Counter
# Represent filesystem paths safely and portably.
from pathlib import Path

# Read modern .xlsx workbooks with openpyxl.
from openpyxl import load_workbook

# Define the food fields extracted from recent menus.
MENU_FIELDS = ('entree', 'proteine', 'feculent', 'legumes', 'dessert_base', 'dessert_resto')
# Define CSV provenance and food columns for meals.
MENU_COLUMNS = ('fichier_source', 'chemin_source', 'onglet', 'ligne_excel',
                'type_menu', *MENU_FIELDS, 'enfant')
# Define the source and classification columns for historical proposals.
PROPOSAL_COLUMNS = ('fichier_source', 'chemin_source', 'onglet', 'ligne_excel',
                    'colonne_excel', 'categorie', 'sous_categorie', 'proposition', 'type_feuille')

# The historical layout uses fixed rows for categories and fixed columns for days.
# Map Excel row numbers to historical food categories.
GRID_ROWS = {
    7: ('entree', 'ENTREE 1'), 8: ('entree', 'ENTREE 2'),
    9: ('entree', 'ENTREE 3'), 10: ('entree', 'ENTREE 4'),
    11: ('plat', 'PLAT 1'), 12: ('plat', 'PLAT 2'),
    13: ('plat', 'ASE SOIR'), 14: ('poisson', 'POISSON'),
    15: ('legume', 'LEGUME 1'), 16: ('legume', 'LEGUME 2'),
    17: ('feculent', 'FECULENTS'), 18: ('laitage', 'LAITAGES'),
    21: ('fromage', 'FROMAGE 1'), 22: ('fromage', 'FROMAGE 2'),
    23: ('dessert', 'DESSERT 1'), 24: ('fruit', 'FRUIT 1'),
    25: ('dessert', 'DESSERT 2'), 26: ('fruit', 'FRUIT 2'),
}
# List the Excel columns used by the historical weekly grid.
GRID_COLUMNS = (4, 6, 8, 10, 12, 14, 16)
# Define the mandatory header names used to recognize a recent menu.
EXPECTED_HEADERS = ('entree', 'proteine', 'feculent', 'legumes')
# Compile a reusable regular expression for day labels.
DAY_LABEL = re.compile(r'(?:LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)'r'(?:\s+(?:MIDI|SOIR))?', re.IGNORECASE)


# Convert an Excel value to readable text, leaving dish names unstandardized.
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


# Normalize text only for structural comparisons, not exported dish content.
def match_text(value):
    """Normalize text only for detecting workbook structure, not CSV values.

    Args:
        value: Cell value or worksheet name.
    Returns:
        str: Accent-insensitive, case-folded text.
    """
    normalized = unicodedata.normalize('NFKD', cell_text(value).casefold())
    return ''.join(char for char in normalized if not unicodedata.combining(char))


# Identify a weekday label so it is not exported as food.
def is_day_label(value):
    """Detect a weekday or weekday/service label in a structural cell.

    Args:
        value: Cell content.
    Returns:
        bool: Whether the entire value is a recognized day label.
    """
    return bool(DAY_LABEL.fullmatch(match_text(value).strip()))


# Determine the workbook layout and week from its directory and filename.
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


# Choose one main worksheet; prefer matching weeks and midi tabs.
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


# Read a limited rectangle from the chosen main worksheet.
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



# Read every nonempty cell from a delivery or Median annex.
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


# Classify eligible historical annexes by worksheet name.
def annex_kind(sheet_name):
    """Only the two explicitly requested historical annex types qualify."""
    normalized = match_text(sheet_name)
    if 'livraison' in normalized:
        return 'livraison'
    if 'median' in normalized:
        return 'median'
    return None


# Convert annex cells into raw records with precise Excel provenance.
def extract_annex(path, root, sheet_name, kind):
    """Store raw nonempty cells; do not mislabel delivery metadata as dishes."""
    return [{**source_info(path, root, sheet_name, row),
             'colonne_excel': col, 'categorie': kind,
             'sous_categorie': 'cellule_brute', 'proposition': value,
             'type_feuille': kind}
            for row, col, value in read_annex_cells(path, sheet_name)]


# Detect menu column headers from rows 3 to 6.
def find_header(cell):
    """Find food columns in recent worksheets using their header labels.

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


# Attach workbook, relative path, worksheet and row provenance.
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


# Extract complete meal rows from recent layouts.
def extract_meals(path, root, sheet, cell, mapping, header_row, recap=False):
    """Extract meal rows while preserving the food-cell text.

    Args:
        path, root (Path): Workbook and input directory.
        sheet (str): Worksheet name.
        cell (callable): 1-based cell accessor.
        mapping (dict): Food field to Excel column mapping.
        header_row (int): Header row number.
        recap (bool): Whether this is a RECAP worksheet.
    Returns:
        list[dict]: Extracted meal records with source metadata.
    """
    meals = []
    for row in range(max(5, header_row + 1), 19):
        # In some MENUS sheets, the day label occupies the entree column.
        shifted = (not recap and 'entree' in mapping
                   and is_day_label(cell(row, mapping['entree'])))
        offset = 1 if shifted else 0
        values = {}
        for field in MENU_FIELDS:
            column = mapping.get(field)
            value = cell_text(cell(row, column + offset)) if column is not None else ''
            # A weekday label is a structural marker, never a dish.
            values[field] = '' if is_day_label(value) else value
        # Preserve the existing output schema; child meals are not extracted.
        values['enfant'] = ''
        if not any(values[field] for field in MENU_FIELDS):
            continue
        meals.append({**source_info(path, root, sheet, row),
                      'type_menu': 'principal', **values})
    return meals


# Extract individual food proposals from fixed historical grid positions.
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
            value = cell_text(cell(row, column))
            if not value or is_day_label(value):
                continue
            proposals.append({**source_info(path, root, sheet, row),
                              'colonne_excel': column, 'categorie': category,
                              'sous_categorie': subcategory, 'proposition': value})
    return proposals



# Write a UTF-8 BOM semicolon-delimited CSV file.
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


# Orchestrate input discovery, extraction, alerts, deduplication and output.
def main():
    """Extract one designated sheet per workbook to a unified CSV.

    Args:
        None. CLI options: --input and --output.
    Returns:
        None.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, default=Path('data/raw'))
    parser.add_argument('--output', type=Path, default=Path('data/processed'))
    parser.add_argument('--ecraser', action='store_true',
                        help='Allow overwriting previously generated extraction CSV files')
    args = parser.parse_args()
    if not args.input.is_dir():
        parser.error(f'Input directory not found: {args.input}')

    # Refuse overwriting either generated CSV unless explicitly requested.
    outputs = ('historique_menus.csv', 'alert.csv')
    existing = [name for name in outputs if (args.output / name).exists()]
    if existing and not args.ecraser:
        parser.error('Output already exists: ' + ', '.join(existing)
                     + '. Choose another --output or use --ecraser.')

    files = sorted(path for path in args.input.rglob('*')
                   if path.is_file() and path.suffix.lower() in ('.xls', '.xlsx')
                   and not path.name.startswith('~$'))
    meals, proposals, annexes, alerts, counts = [], [], [], [], Counter()

    # Collect alerts in memory without per-file terminal output.
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

    # Process every discovered Excel workbook.
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
            # Extract annex sheets only for 2022 and 2023.
            # 2022-2023: read delivery and Median tabs as additional raw cells.
            # 2024-2026: preserve the original one-sheet-only behavior.
            year = int(next(part for part in path.relative_to(args.input).parts[:-1]
                            if re.fullmatch(r"20\d{2}", part)))
            remaining = []
            if year in (2022, 2023):
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
                extracted = extract_meals(path, args.input, name, cell, mapping,
                                          header, recap=(match_text(name).strip() == 'recap'))
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

    # Unify meal, historical proposal and annex records into one schema.
    columns = ('type_enregistrement', *MENU_COLUMNS,
               'colonne_excel', 'categorie', 'sous_categorie', 'proposition', 'type_feuille')
    records = ([{'type_enregistrement': 'repas', **row} for row in meals]
               + [{'type_enregistrement': 'plat_historique', 'type_feuille': 'principale', **row}
                  for row in proposals]
               + [{'type_enregistrement': 'cellule_annexe', **row} for row in annexes])
    # Remove duplicates by source coordinates, not by dish text.
    # Coordinates identify a unique source cell, even if its text occurs elsewhere.
    unique_records = {}
    for record in records:
        key = (record.get('chemin_source'), record.get('onglet'),
               record.get('ligne_excel'), record.get('colonne_excel'),
               record.get('type_enregistrement'))
        unique_records.setdefault(key, record)
    records = list(unique_records.values())
    args.output.mkdir(parents=True, exist_ok=True)
    write_csv(args.output / 'historique_menus.csv', columns, records)
    write_csv(args.output / 'alert.csv',
              ('fichier_source', 'niveau', 'code', 'onglet', 'nombre_feuilles', 'message'),
              alerts)

    # Calculate final summary counts and print one terminal line.
    successful = counts['MENUS'] + counts['RECAP'] + counts['GRILLE']
    error_count = sum(item['niveau'] == 'ERREUR' for item in alerts)
    warning_count = sum(item['niveau'] == 'AVERTISSEMENT' for item in alerts)
    print(f"Extraction terminee | fichiers : {len(files)} | traites : {successful} "
          f"| non traites : {len(files) - successful} | erreurs : {error_count} "
          f"| avertissements : {warning_count} | repas : {len(meals)} "
          f"| propositions : {len(proposals)} | feuilles annexes : {counts['annex_sheets']} "
          f"| cellules annexes : {len(annexes)} | sorties : "
          f"{args.output / 'historique_menus.csv'}, {args.output / 'alert.csv'}")


# Run main only when this file is executed directly.
if __name__ == '__main__':
    main()