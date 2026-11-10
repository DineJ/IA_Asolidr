#!/usr/bin/env python3
"""Clean yearly menu histories without changing the original source files.

Example: python clean_menu.py --entree data/processed --sortie data/processed/nettoyage
Dependencies: pip install pandas rapidfuzz

Only 2025 and 2026 are processed. Outputs include one cleaned CSV per year,
alert.csv and modifications.csv. MENU, weekday names, MIDI and SOIR are kept.
Validated corrections are logged; RapidFuzz only suggests possible typos.
Empty values are exported as the literal NaN marker.
"""

import argparse
import re
import unicodedata
from pathlib import Path
import pandas as pd

CHAMPS_REPAS = ('entree', 'proteine', 'feculent', 'legumes', 'dessert_base', 'dessert_resto')
VIDES = {'', '(vide)', '(vides)', 'nan', 'none', 'null', 'n/a', 'na', '-', '--', '---', '/'}
JOUR = re.compile(r'(?:LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)(?:\s+(?:MIDI|SOIR))?', re.I)
QUANTITE_SEULE = re.compile(r'\d+(?:[.,]\d+)?\s*(?:KG|KGS|G|GRAMMES?|PARTS?|PORTIONS?)?', re.I)
QUANTITE_DANS = re.compile(r'\b\d+(?:[.,]\d+)?\s*(?:KG|KGS|G|GRAMMES?|PARTS?|PORTIONS?)\b', re.I)
CONSIGNE = re.compile(r'\b(?:PREVOIR|PRÉVOIR|ATTENTION|DLC|À DÉCOUPER|A DECOUPER|PAS DE REPAS)\b', re.I)
# Validated replacements live in corrections.csv.
# RapidFuzz suggests possible typos but never changes menu text automatically.
import csv
from rapidfuzz import fuzz, process


def charger_corrections(chemin):
    """Load validated word and phrase corrections from a semicolon-separated CSV."""
    if not chemin.is_file():
        raise FileNotFoundError(f'Fichier de corrections introuvable : {chemin}')
    regles = {'mot': {}, 'expression': {}}
    with chemin.open(encoding='utf-8-sig', newline='') as fichier:
        lecteur = csv.DictReader(fichier, delimiter=';')
        if not {'type', 'incorrect', 'correct'} <= set(lecteur.fieldnames or []):
            raise ValueError('corrections.csv doit contenir type;incorrect;correct')
        for numero, ligne in enumerate(lecteur, 2):
            type_regle = (ligne['type'] or '').strip().lower()
            incorrect = (ligne['incorrect'] or '').strip().upper()
            correct = (ligne['correct'] or '').strip().upper()
            if type_regle not in regles or not incorrect or not correct:
                raise ValueError(f'Règle invalide à la ligne {numero} de {chemin}')
            if incorrect in regles[type_regle] and regles[type_regle][incorrect] != correct:
                raise ValueError(f'Corrections contradictoires pour {incorrect} (ligne {numero})')
            regles[type_regle][incorrect] = correct
    return regles


def suggestions_rapidfuzz(texte, vocabulaire, seuil=88):
    """Suggest similar words from the validated vocabulary without replacing them."""
    # Ignore short words to reduce false positives.
    resultat = []
    vus = set()
    for mot in re.findall(r"[A-ZÀ-ÖØ-ÝŒ]+(?:['’-][A-ZÀ-ÖØ-ÝŒ]+)*", texte):
        if len(mot) < 5 or mot in vocabulaire or mot in vus:
            continue
        vus.add(mot)
        proposition = process.extractOne(mot, vocabulaire, scorer=fuzz.ratio,
                                         score_cutoff=seuil)
        if proposition and proposition[0] != mot:
            resultat.append((mot, proposition[0], round(proposition[1], 1)))
    return resultat


# Keep service labels because they may be meaningful source information.
SERVICE = re.compile(r'\bMENU\s+(?:LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)\s+(?:MIDI|SOIR)\b', re.I)


def nettoyer_avec_regles(valeur, regles):
    """Clean one value and return each applied rule with its before/after values.

    Case conversion and whitespace normalization are intentionally not logged.
    """
    if pd.isna(valeur):
        return pd.NA, []

    # Normalize case and whitespace without adding change-log entries.
    texte = re.sub(r'\s+', ' ', str(valeur).replace('\u00a0', ' ')).strip().upper()
    etapes = []

    def enregistrer(code, explication, apres):
        nonlocal texte
        if apres != texte:
            etapes.append((code, explication, texte, apres))
            texte = apres

    # Convert recognized empty-value markers to missing values.
    if texte.casefold() in VIDES:
        if texte:
            etapes.append(('REMPLACE_PAR_NAN',
                           f'Valeur considérée comme vide ({texte}) : remplacement par NaN.',
                           texte, 'NaN'))
        return pd.NA, etapes

    # Remove the (HALAL) label and normalize the remaining whitespace.
    apres, nombre = re.subn(r'\(\s*HALAL\s*\)', '', texte)
    if nombre:
        apres = re.sub(r'\s+', ' ', apres).strip()
        enregistrer('MENTION_HALAL_SUPPRIMEE',
                    f'Mention (HALAL) supprimée ({nombre} occurrence(s)).', apres)

    # Normalize apostrophes and compatible Unicode characters; log changes.
    apres = unicodedata.normalize('NFKC', texte).replace('’', "'").replace('`', "'")
    enregistrer('APOSTROPHE_NORMALISEE',
                "Harmonisation des caractères typographiques et apostrophes.", apres)

    # Apply longer validated phrases first and log each replacement.
    for erreur, correction in sorted(regles['expression'].items(),
                                     key=lambda item: -len(item[0])):
        apres, nombre = re.subn(r'(?<!\w)' + re.escape(erreur) + r'(?!\w)',
                                correction, texte)
        if nombre:
            enregistrer('EXPRESSION_CORRIGEE',
                        f'Expression corrigée : {erreur} → {correction} ({nombre} occurrence(s)).',
                        apres)

    # Apply validated whole-word spelling and accent corrections.
    for erreur, correction in regles['mot'].items():
        if erreur == correction:
            continue
        apres, nombre = re.subn(r'(?<!\w)' + re.escape(erreur) + r'(?!\w)',
                                correction, texte)
        if nombre:
            enregistrer('ORTHOGRAPHE_CORRIGEE',
                        f'Orthographe/accents : {erreur} → {correction} ({nombre} occurrence(s)).',
                        apres)

    # Treat values emptied by cleaning as missing.
    if not texte:
        etapes.append(('REMPLACE_PAR_NAN',
                       'Aucun texte ne reste après nettoyage : remplacement par NaN.',
                       '', 'NaN'))
        return pd.NA, etapes
    return texte, etapes


def normaliser(valeur, regles):
    """Return only the cleaned value, without the change log."""
    return nettoyer_avec_regles(valeur, regles)[0]


def motifs_suspects(texte):
    """Flag suspicious menu values without changing their content."""
    if pd.isna(texte):
        return []
    motifs = []
    if JOUR.fullmatch(texte):
        motifs.append('JOUR_SERVICE')
    if QUANTITE_SEULE.fullmatch(texte):
        motifs.append('QUANTITE_OU_NOMBRE_SEUL')
    elif QUANTITE_DANS.search(texte):
        motifs.append('QUANTITE_DANS_INTITULE')
    if CONSIGNE.search(texte):
        motifs.append('CONSIGNE')
    if SERVICE.search(texte):
        motifs.append('MENTION_SERVICE')
    if len(texte) > 200:
        motifs.append('INTITULE_TRES_LONG')
    return motifs


def traiter_fichier(chemin, annee, regles, vocabulaire, seuil):
    """Clean one yearly file and collect alerts and detailed modifications."""
    df = pd.read_csv(chemin, sep=';', encoding='utf-8-sig', dtype=str, keep_default_na=False)
    if annee not in (2025, 2026):
        raise ValueError(f'Annee non prise en charge : {annee}')
    champs = CHAMPS_REPAS
    requis = {'chemin_source', 'ligne_excel', *champs}
    absents = requis.difference(df.columns)
    if absents:
        raise ValueError(f'{chemin.name} : colonnes manquantes : {", ".join(sorted(absents))}')

    alertes = []
    modifications = []
    # Preserve provenance columns and clean only the menu-item fields.
    for champ in champs:
        valeurs_brutes = df[champ].copy()
        valeurs_nettoyees = []
        for index, brut in enumerate(valeurs_brutes):
            propre, etapes = nettoyer_avec_regles(brut, regles)
            valeurs_nettoyees.append(propre)
            source = df.iloc[index]

            # Record one log entry per applied rule, in execution order.
            for numero, (code, explication, avant, apres) in enumerate(etapes, 1):
                modifications.append({
                    'annee': annee,
                    'chemin_source': source['chemin_source'],
                    'ligne_excel': source['ligne_excel'],
                    'colonne_excel': source.get('colonne_excel', ''),
                    'champ': champ,
                    'numero_etape': numero,
                    'regle_appliquee': code,
                    'explication': explication,
                    'valeur_avant_regle': avant,
                    'valeur_apres_regle': apres,
                    'valeur_originale': brut,
                    'valeur_nettoyee': 'NaN' if pd.isna(propre) else propre,
                })

            # Send RapidFuzz suggestions to alert.csv for manual review only.
            if not pd.isna(propre):
                for mot, suggestion, score in suggestions_rapidfuzz(
                        propre, vocabulaire, seuil):
                    alertes.append({
                        'annee': annee,
                        'chemin_source': source['chemin_source'],
                        'ligne_excel': source['ligne_excel'],
                        'colonne_ou_champ': champ,
                        'valeur_originale': brut,
                        'valeur_nettoyee': propre,
                        'motif': f'SUGGESTION_ORTHOGRAPHE : {mot} → {suggestion} (score {score})',
                    })

            for motif in motifs_suspects(propre):
                alertes.append({
                    'annee': annee,
                    'chemin_source': source['chemin_source'],
                    'ligne_excel': source['ligne_excel'],
                    'colonne_ou_champ': champ,
                    'valeur_originale': brut,
                    'valeur_nettoyee': propre,
                    'motif': motif,
                })
        df[champ] = valeurs_nettoyees

    # Flag rows with no menu items in any category.
    sans_plat = df[list(champs)].isna().all(axis=1)
    for _, ligne in df.loc[sans_plat].iterrows():
        alertes.append({
            'annee': annee,
            'chemin_source': ligne['chemin_source'],
            'ligne_excel': ligne['ligne_excel'],
            'colonne_ou_champ': 'repas',
            'valeur_originale': '',
            'valeur_nettoyee': pd.NA,
            'motif': 'AUCUN_PLAT',
        })
    return df, alertes, modifications


def main():
    """Validate CLI arguments and generate the cleaned files and audit logs."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--entree', type=Path, default=Path('data/processed'),
                        help='Dossier des historique_menus_YYYY.csv')
    parser.add_argument('--sortie', type=Path, default=Path('data/processed/nettoyage'))
    parser.add_argument('--ecraser', action='store_true', help='Autoriser le remplacement des sorties')
    parser.add_argument('--corrections', type=Path, default=Path(__file__).with_name('corrections.csv'),
                        help='CSV des corrections validées')
    parser.add_argument('--seuil', type=int, default=88,
                        help='Score minimal RapidFuzz des suggestions (0 à 100, défaut 88)')
    args = parser.parse_args()

    if not 0 <= args.seuil <= 100:
        parser.error('--seuil doit être compris entre 0 et 100')
    try:
        regles = charger_corrections(args.corrections)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    # Build the fuzzy-matching vocabulary from validated spellings only.
    vocabulaire = sorted(set(regles['mot'].values()))
    if not args.entree.is_dir():
        parser.error(f'Dossier introuvable : {args.entree}')
    if args.entree.resolve() == args.sortie.resolve():
        parser.error('Le dossier de sortie doit etre different du dossier source.')

    # Ignore years 2022-2024 and unrelated files.
    sources = [args.entree / f'historique_menus_{annee}.csv' for annee in (2025, 2026)
               if (args.entree / f'historique_menus_{annee}.csv').is_file()]
    if not sources:
        parser.error('Aucun historique_menus_2025.csv ou historique_menus_2026.csv trouve.')

    destinations = [args.sortie / f'{p.stem}_nettoye.csv' for p in sources]
    destinations.extend([args.sortie / 'alert.csv', args.sortie / 'modifications.csv'])
    deja = [p.name for p in destinations if p.exists()]
    if deja and not args.ecraser:
        parser.error(f'Sorties deja presentes : {", ".join(deja)}. Utiliser --ecraser.')

    # Process all input files before writing any output.
    resultats, alertes, modifications = [], [], []
    for source in sources:
        annee = int(source.stem.rsplit('_', 1)[-1])
        propre, problemes, changements = traiter_fichier(source, annee, regles, vocabulaire, args.seuil)
        resultats.append((args.sortie / f'{source.stem}_nettoye.csv', propre))
        alertes.extend(problemes)
        modifications.extend(changements)

    args.sortie.mkdir(parents=True, exist_ok=True)
    for destination, donnees in resultats:
        donnees.to_csv(destination, sep=';', index=False, encoding='utf-8-sig', na_rep='NaN')
        print(f'{destination.name} : {len(donnees)} lignes')

    colonnes = ('annee', 'chemin_source', 'ligne_excel', 'colonne_ou_champ',
                'valeur_originale', 'valeur_nettoyee', 'motif')
    pd.DataFrame(alertes, columns=colonnes).to_csv(
        args.sortie / 'alert.csv', sep=';', index=False, encoding='utf-8-sig', na_rep='NaN')
    colonnes_modifications = ('annee', 'chemin_source', 'ligne_excel',
                              'colonne_excel', 'champ', 'numero_etape',
                              'regle_appliquee', 'explication',
                              'valeur_avant_regle', 'valeur_apres_regle',
                              'valeur_originale', 'valeur_nettoyee')
    pd.DataFrame(modifications, columns=colonnes_modifications).to_csv(
        args.sortie / 'modifications.csv', sep=';', index=False,
        encoding='utf-8-sig', na_rep='NaN')
    print(f'alert.csv : {len(alertes)} alertes')
    print(f'modifications.csv : {len(modifications)} regles appliquees '
          '(hors majuscules et espaces)')
    print('Fichiers sources inchanges.')


if __name__ == '__main__':
    main()