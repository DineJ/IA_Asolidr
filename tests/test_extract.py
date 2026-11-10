"""Unit and integration tests for extract.py using data/raw/YYYY/."""
import csv
import subprocess
import sys
from pathlib import Path

import pytest

from extract import expected_sheet, choose_sheet


@pytest.mark.parametrize(("year", "filename", "expected"), [
    (2022, "menu_sem13_2022.xlsx", ("GRILLE", 13)),
    (2023, "menu_semaine_08_2023.xls", ("GRILLE", 8)),
    (2024, "menu_sem-25_2024.xlsx", ("GRILLE", 25)),
    (2025, "menu_2025.xlsx", ("RECAP", None)),
    (2026, "menu_2026.xlsx", ("MENUS", None)),
])
def test_detect_year_layout(tmp_path, year, filename, expected):
    """Layout detection depends on year folders, not real Excel content."""
    path = tmp_path / str(year) / filename
    assert expected_sheet(path, tmp_path) == expected


def test_reject_missing_year(tmp_path):
    with pytest.raises(ValueError, match="Year folder"):
        expected_sheet(tmp_path / "menu.xlsx", tmp_path)


def test_reject_missing_week(tmp_path):
    with pytest.raises(ValueError, match="Week number"):
        expected_sheet(tmp_path / "2022" / "menu.xlsx", tmp_path)


def test_select_recap_sheet():
    assert choose_sheet(["Autre", "RECAP", "MENUS"], "RECAP", None)[0] == "RECAP"


def test_select_2026_menus_sheet():
    assert choose_sheet(["Autre", "MENUS"], "MENUS", None)[0] == "MENUS"


def test_extract_from_real_raw_folder(raw_dir, project_root, tmp_path):
    """Integration test: read real Excel files but write only to a temp folder."""
    if not raw_dir.is_dir():
        pytest.skip(f"Missing input directory: {raw_dir}")

    excel_files = [p for p in raw_dir.rglob("*")
                   if p.is_file() and p.suffix.lower() in (".xls", ".xlsx")
                   and not p.name.startswith("~$")]
    if not excel_files:
        pytest.skip(f"No Excel files in {raw_dir}")

    output = tmp_path / "processed"
    result = subprocess.run(
        [sys.executable, str(project_root / "extract.py"),
         "--input", str(raw_dir), "--output", str(output)],
        cwd=project_root, capture_output=True, text=True,
    )
    assert result.returncode == 0, result.stdout + "\n" + result.stderr
    assert (output / "alert.csv").exists()

    generated = sorted(output.glob("historique_menus_*.csv"))
    assert generated, "No yearly CSV generated; inspect alert.csv"
    for csv_path in generated:
        with csv_path.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.DictReader(handle, delimiter=";"))
        assert rows, f"Empty CSV: {csv_path.name}"
        year = int(csv_path.stem.rsplit("_", 1)[1])
        if year <= 2024:
            assert {"categorie", "proposition", "colonne_excel"} <= rows[0].keys()
        else:
            assert {"entree", "proteine", "feculent", "legumes"} <= rows[0].keys()
