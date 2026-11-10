"""Unit tests for dish normalization and year-separated counting."""
import csv
from pathlib import Path

from occurrences_menus import compter, nettoyer_pour_compter


def test_normalize_without_splitting_dish():
    assert nettoyer_pour_compter('  salade de chou, vinaigrette  ') == 'SALADE DE CHOU, VINAIGRETTE'
    assert nettoyer_pour_compter('NaN') == ''


def test_count_dishes_by_year(tmp_path):
    for year, dishes in [(2025, ['POULET RÔTI', 'PURÉE']),
                         (2026, ['POULET RÔTI', 'POULET RÔTI'])]:
        path = tmp_path / f'historique_menus_{year}_nettoye.csv'
        with path.open('w', encoding='utf-8-sig', newline='') as handle:
            writer = csv.DictWriter(handle, fieldnames=['entree', 'proteine'], delimiter=';')
            writer.writeheader()
            for dish in dishes:
                writer.writerow({'entree': dish, 'proteine': 'NaN'})
    counters = compter(tmp_path)
    assert counters[2025]['POULET RÔTI'] == 1
    assert counters[2026]['POULET RÔTI'] == 2
    assert counters[2025]['PURÉE'] == 1
    assert 'NAN' not in counters[2025]
