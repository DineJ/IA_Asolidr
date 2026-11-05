#!/usr/bin/env python3
"""Extract Excel menu data without normalizing dish names.

Usage: python extract_corrige_v2.py --input data/raw --output data/processed
Dependencies: openpyxl and xlrd (for legacy .xls files).

The two output CSV files are overwritten when the program runs. Original
Excel workbooks are never modified.
"""

import argparse
import csv
import re
import unicodedata
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

MENU_FIELDS = ('entree', 'proteine', 'feculent', 'legumes', 'dessert_base', 'dessert_resto')
MENU_COLUMNS = ('fichier_source', 'chemin_source', 'onglet', 'ligne_excel',
                'type_menu', *MENU_FIELDS, 'enfant')
PROPOSAL_COLUMNS = ('fichier_source', 'chemin_source', 'onglet', 'ligne_excel',
                    'colonne_excel', 'categorie', 'sous_categorie', 'proposition')

# The historical layout uses fixed rows for categories and fixed columns for days.
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
GRID_COLUMNS = (4, 6, 8, 10, 12, 14, 16)
EXPECTED_HEADERS = ('entree', 'proteine', 'feculent', 'legumes')
DAY_LABEL = re.compile(
    r'(?:LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)'
    r'(?:\s+(?:MIDI|SOIR))?', re.IGNORECASE
)


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


def match_text(value):
    """Normalize text only for detecting workbook structure, not CSV values.

    Args:
        value: Cell value or worksheet name.
    Returns:
        str: Accent-insensitive, case-folded text.
    """
    normalized = unicodedata.normalize('NFKD', cell_text(value).casefold())
    return ''.join(char for char in normalized if not unicodedata.combining(char))


def is_day_label(value):
    """Detect a weekday or weekday/service label in a structural cell.

    Args:
        value: Cell content.
    Returns:
        bool: Whether the entire value is a recognized day label.
    """
    return bool(DAY_LABEL.fullmatch(match_text(value).strip()))


def read_sheets(path):
    """Yield sheet names and 1-based cell accessors from an Excel workbook.

    Args:
        path (Path): .xls or .xlsx workbook path.
    Yields:
        tuple[str, callable]: Sheet name and cell(row, column) accessor.
    """
    if path.suffix.lower() == '.xls':
        import xlrd
        workbook = xlrd.open_workbook(str(path))
        for sheet in workbook.sheets():
            def cell(row, column, data=sheet):
                """Return a legacy Excel cell value, or None outside the sheet."""
                if 1 <= row <= data.nrows and 1 <= column <= data.ncols:
                    return data.cell_value(row - 1, column - 1)
                return None
            yield sheet.name, cell
    else:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            for sheet in workbook:
                sheet.reset_dimensions()
                # Cache the area used by both known menu formats.
                rows = list(sheet.iter_rows(min_row=1, max_row=65, min_col=1,
                                            max_col=22, values_only=True))
                def cell(row, column, data=rows):
                    """Return a modern Excel cell value, or None outside the cache."""
                    if 1 <= row <= len(data) and 1 <= column <= len(data[row - 1]):
                        return data[row - 1][column - 1]
                    return None
                yield sheet.title, cell
        finally:
            workbook.close()


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


def detect_format(sheets):
    """Select a MENUS, RECAP or historical grid sheet.

    Args:
        sheets (list[tuple]): Sheet names and cell accessors.
    Returns:
        tuple | None: Format, name, accessor, mapping and header row.
    """
    for target in ('menus', 'recap'):
        for name, cell in sheets:
            if match_text(name).strip() == target:
                mapping, header = find_header(cell)
                if mapping:
                    return target.upper(), name, cell, mapping, header
    for name, cell in sheets:
        if 'entree' in match_text(cell(7, 2)) and 'plat' in match_text(cell(11, 2)):
            return 'GRILLE', name, cell, None, None
    return None


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


def main():
    """Extract supported workbooks and write the two CSV output files.

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

    files = sorted(path for path in args.input.rglob('*')
                   if path.is_file() and path.suffix.lower() in ('.xls', '.xlsx')
                   and not path.name.startswith('~$'))
    meals, proposals, counts = [], [], Counter()
    for path in files:
        relative = path.relative_to(args.input).as_posix()
        try:
            selection = detect_format(list(read_sheets(path)))
            if selection is None:
                counts['unknown'] += 1
                print(f'WARNING unknown format: {relative}')
                continue
            fmt, name, cell, mapping, header = selection
            if fmt == 'GRILLE':
                extracted = extract_grid(path, args.input, name, cell)
                proposals.extend(extracted)
            else:
                extracted = extract_meals(path, args.input, name, cell, mapping,
                                          header, recap=(fmt == 'RECAP'))
                meals.extend(extracted)
            counts[fmt] += 1
            if not extracted:
                print(f'WARNING no data: {relative} ({fmt})')
            print(f'{relative}: {len(extracted)} rows ({fmt})')
        except Exception as exc:
            counts['errors'] += 1
            print(f'ERROR {relative}: {type(exc).__name__}: {exc}')

    args.output.mkdir(parents=True, exist_ok=True)
    existing = [name for name in ('menus.csv', 'propositions_anciennes.csv')
                if (args.output / name).exists()]
    if existing and not args.ecraser:
        parser.error('Output already exists: ' + ', '.join(existing)
                     + '. Choose another --output or use --ecraser.')
    write_csv(args.output / 'menus.csv', MENU_COLUMNS, meals)
    write_csv(args.output / 'propositions_anciennes.csv', PROPOSAL_COLUMNS, proposals)
    print(f'Files: {len(files)} | MENUS: {counts["MENUS"]} | '
          f'RECAP: {counts["RECAP"]} | GRILLE: {counts["GRILLE"]} | '
          f'unknown: {counts["unknown"]} | errors: {counts["errors"]}')
    print(f'Meals: {len(meals)} | Proposals: {len(proposals)}')
    print(f'Output: {args.output.resolve()}')


if __name__ == '__main__':
    main()