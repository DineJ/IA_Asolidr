#!/usr/bin/env python3
"""Run the IA_Asolidr extraction, cleaning and occurrence-counting pipeline."""

import argparse
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent


def run_step(script, arguments):
    """Execute one existing project script and stop immediately on failure."""
    command = [sys.executable, str(PROJECT_ROOT / script), *map(str, arguments)]
    print(f'\n>>> {script}', flush=True)
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, default=Path('data/raw'), help='Excel source directory')
    parser.add_argument('--output', type=Path, default=Path('data/processed'), help='Extraction output directory')
    parser.add_argument('--corrections', type=Path, default=PROJECT_ROOT / 'corrections.csv')
    parser.add_argument('--seuil', type=int, default=88, help='RapidFuzz suggestion threshold')
    parser.add_argument('--ecraser', action='store_true', help='Overwrite previously generated CSVs')
    parser.add_argument('--inclure-annexes', action='store_true', help='Extract optional annex data')
    args = parser.parse_args()

    # Resolve paths relative to the repository, not the caller's working directory.
    def project_path(path):
        return path if path.is_absolute() else PROJECT_ROOT / path

    raw = project_path(args.input)
    output = project_path(args.output)
    cleaned = output / 'nettoyage'
    corrections = project_path(args.corrections)
    if not raw.is_dir():
        parser.error(f'Source directory not found: {raw}')
    if not corrections.is_file():
        parser.error(f'Correction rules not found: {corrections}')
    if not 0 <= args.seuil <= 100:
        parser.error('--seuil must be between 0 and 100')
    overwrite = ['--ecraser'] if args.ecraser else []

    try:
        run_step('extract.py', ['--input', raw, '--output', output,
                                *(['--inclure-annexes'] if args.inclure_annexes else []), *overwrite])
        run_step('clean_menu.py', ['--entree', output, '--sortie', cleaned,
                                   '--corrections', corrections, '--seuil', args.seuil, *overwrite])
        run_step('occurrences_menus.py', ['--entree', cleaned,
                                          '--sortie', cleaned / 'occurrences_plats.csv', *overwrite])
    except subprocess.CalledProcessError as exc:
        print(f'Pipeline stopped: {exc.cmd[1]} returned exit code {exc.returncode}', file=sys.stderr)
        return exc.returncode or 1
    print(f'\nPipeline completed successfully. Results: {cleaned}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
