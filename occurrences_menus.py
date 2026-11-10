#!/usr/bin/env python3
"""Count cleaned menu-item occurrences for 2025 and 2026.

Example:
    python occurrences_menus.py --entree data/processed/nettoyage \
        --sortie data/processed/nettoyage/occurrences_plats.csv

Read the yearly cleaned CSV files produced by clean_menu.py.
The source files are never modified.
"""

import argparse
import csv
from collections import Counter
from pathlib import Path

ANNEES = (2025, 2026)
CHAMPS = ('entree', 'proteine', 'feculent', 'legumes', 'dessert_base', 'dessert_resto')
VALEURS_VIDES = {'', 'NAN', '(VIDE)', 'NONE', 'NULL', 'N/A'}
COLONNES_SORTIE = ('plat', 'occurrences_2025', 'occurrences_2026')


def nettoyer_pour_compter(valeur):
    """Normalize whitespace and case without splitting dish names."""
    plat = ' '.join((valeur or '').replace('\u00a0', ' ').split()).upper()
    return '' if plat in VALEURS_VIDES else plat


def compter(dossier):
    """Count every non-empty menu field separately for each available year."""
    compteurs = {annee: Counter() for annee in ANNEES}
    fichiers_trouves = []
    for annee in ANNEES:
        chemin = dossier / f'historique_menus_{annee}_nettoye.csv'
        if not chemin.is_file():
            print(f'Information : fichier absent, annee ignoree : {chemin}')
            continue
        fichiers_trouves.append(chemin)
        with chemin.open('r', encoding='utf-8-sig', newline='') as fichier:
            lecteur = csv.DictReader(fichier, delimiter=';')
            colonnes = lecteur.fieldnames or []
            champs_presents = [champ for champ in CHAMPS if champ in colonnes]
            if not champs_presents:
                raise ValueError(f'{chemin} : aucune colonne de plat reconnue.')
            for ligne in lecteur:
                for categorie in champs_presents:
                    plat = nettoyer_pour_compter(ligne.get(categorie))
                    if plat:
                        compteurs[annee][plat] += 1
    if not fichiers_trouves:
        raise FileNotFoundError('Aucun historique nettoye 2025 ou 2026 trouve dans ' + str(dossier))
    return compteurs


def main():
    """Parse CLI options, aggregate counts and export the result CSV."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--entree', type=Path, default=Path('data/processed/nettoyage'),
                        help='Dossier des historiques nettoyes 2025 et 2026')
    parser.add_argument('--sortie', type=Path,
                        default=Path('data/processed/nettoyage/occurrences_plats.csv'),
                        help='Chemin du CSV des occurrences')
    parser.add_argument('--ecraser', action='store_true', help='Autoriser le remplacement du CSV de sortie')
    args = parser.parse_args()
    if args.sortie.exists() and not args.ecraser:
        parser.error('Le fichier de sortie existe deja : utiliser --ecraser.')
    try:
        compteurs = compter(args.entree)
    except (OSError, ValueError) as erreur:
        parser.error(str(erreur))

    cles = set(compteurs[2025]) | set(compteurs[2026])
    lignes = []
    for plat in cles:
        n2025 = compteurs[2025][plat]
        n2026 = compteurs[2026][plat]
        lignes.append({'plat': plat, 'occurrences_2025': n2025,
                       'occurrences_2026': n2026})
    lignes.sort(key=lambda ligne: (-(ligne['occurrences_2025'] + ligne['occurrences_2026']), ligne['plat']))
    args.sortie.parent.mkdir(parents=True, exist_ok=True)
    with args.sortie.open('w', encoding='utf-8-sig', newline='') as fichier:
        writer = csv.DictWriter(fichier, fieldnames=COLONNES_SORTIE, delimiter=';')
        writer.writeheader()
        writer.writerows(lignes)
    print(f'{len(lignes)} plats distincts exportes dans {args.sortie}')


if __name__ == '__main__':
    main()