
#!/usr/bin/env python3
# Clean extracted Excel data while retaining source values and review flags.
# Dependances : pip install pandas
from pathlib import Path
import argparse
import re
import unicodedata
import pandas as pd

CHAMPS_MENU = ['entree', 'proteine', 'feculent', 'legumes', 'dessert_base', 'dessert_resto', 'enfant']
VIDES = {'', '-', '--', '---', '/', 'nan', 'none', 'null', 'n/a', '0'}

# Uniquement les quantites en fin de texte. Ne pas toucher aux noms de recettes
# tels que "4 epices", "7 legumes" ou "2 chocolats".
QUANTITE_FIN = re.compile(
    r'(?i)(?:\s*[,;:/-]?\s*'
    r'(?:\(\s*\d+(?:[.,]\d+)?\s*'
    r'(?:kg|kgs|g|grammes?|parts?|portions?)\s*\)'
    r'|\d+(?:[.,]\d+)?\s*'
    r'(?:kg|kgs|g|grammes?|parts?|portions?))'
    r'\s*)+$'
)
QUANTITE_DANS = re.compile(
    r'(?i)\b\d+(?:[.,]\d+)?\s*(?:kg|kgs|g|grammes?|parts?|portions?)\b'
)
NOMBRE_SEUL = re.compile(r'[+-]?\d+(?:[.,]\d+)?')
QUANTITE_SEULE = re.compile(
    r'(?i)\d+(?:[.,]\d+)?\s*(?:kg|kgs|g|grammes?|parts?|portions?)'
)
JOUR_SERVICE = re.compile(
    r'(?i)(?:lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)'
    r'(?:\s+(?:midi|soir))?'
)

# A service reference embedded in a dish name is not a standalone dish label.
SERVICE_IN_TITLE = re.compile(
    r'(?i)\bmenu\s+(?:lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)'
    r'\s+(?:midi|soir)\b'
)

CONSIGNE = re.compile(
    r'(?i)\b(?:prevoir|prévoir|attention|dlc|à découper|a decouper|'
    r'pas de repas|cuisine\s*\+\s*menage|cuisine\s*\+\s* ménage)\b'
)


def normaliser_espaces(valeur):
    if pd.isna(valeur):
        return ''
    s = str(valeur).replace('\u00a0', ' ').replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r'\s+', ' ', s).strip(' \t;|')
    if s.casefold() in VIDES or not re.search(r'\w', s, re.UNICODE):
        return ''
    return s


def nettoyer(valeur):
    original = normaliser_espaces(valeur)
    if not original:
        return '', ''

    # Flag structural or quantity-only values without losing their source text.
    if NOMBRE_SEUL.fullmatch(original) or QUANTITE_SEULE.fullmatch(original):
        return original, 'quantite_seule_a_verifier'
    if JOUR_SERVICE.fullmatch(original):
        return original, 'jour_service_a_verifier'

    nettoye = QUANTITE_FIN.sub('', original).strip(' ,;:/-')

    if not nettoye:
        return original, 'quantite_seule_a_verifier'

    remarques = []

    if nettoye != original:
        remarques.append('quantite_finale_retirée')

    if QUANTITE_DANS.search(nettoye):
        remarques.append('quantite_interne_a_verifier')

    if CONSIGNE.search(nettoye):
        remarques.append('consigne_a_verifier')

    if SERVICE_IN_TITLE.search(nettoye):
        remarques.append('mention_service_a_verifier')

    return nettoye, '|'.join(remarques)


def cle_texte(texte):
    # Regrouper les variantes de casse et d'espaces sans corriger
    # automatiquement les fautes ni fusionner des recettes differentes.
    texte = unicodedata.normalize('NFKC', texte).casefold()
    return re.sub(r'\s+', ' ', texte).strip()


def lire_csv(chemin, requis):
    if not chemin.exists():
        raise FileNotFoundError(f'Fichier introuvable : {chemin}')

    df = pd.read_csv(
        chemin,
        sep=';',
        encoding='utf-8-sig',
        dtype=str,
        keep_default_na=False
    )

    manquantes = set(requis) - set(df.columns)

    if manquantes:
        raise ValueError(
            f'{chemin.name} : colonnes absentes : {sorted(manquantes)}'
        )

    return df


def traiter_menus(df):
    df = df.copy()

    df.insert(
        0,
        'id_repas',
        [f'M{i:05d}' for i in range(1, len(df) + 1)]
    )

    alertes = []

    for col in CHAMPS_MENU:
        if col not in df.columns:
            df[col] = ''

        originaux = df[col].map(normaliser_espaces)
        nettoyes = originaux.map(nettoyer)

        df[col + '_original'] = originaux
        df[col] = nettoyes.map(lambda x: x[0])
        df[col + '_alerte'] = nettoyes.map(lambda x: x[1])

        for idx, ((_, remarque), original) in enumerate(
            zip(nettoyes, originaux)
        ):
            if remarque:
                alertes.append({
                    'origine': 'menus',
                    'identifiant': df.iloc[idx]['id_repas'],
                    'fichier_source': df.iloc[idx].get('fichier_source', ''),
                    'champ': col,
                    'valeur_originale': original,
                    'motif': remarque
                })

    df['nb_elements_renseignes'] = df[CHAMPS_MENU].ne('').sum(axis=1)

    df['statut_structure'] = df['nb_elements_renseignes'].map(
        lambda n: (
            'vide_a_verifier' if n == 0
            else 'incomplet_a_verifier' if n == 1
            else 'conserve'
        )
    )

    for _, ligne in df[df['nb_elements_renseignes'] <= 1].iterrows():
        alertes.append({
            'origine': 'menus',
            'identifiant': ligne['id_repas'],
            'fichier_source': ligne.get('fichier_source', ''),
            'champ': 'repas',
            'valeur_originale': '',
            'motif': ligne['statut_structure']
        })

    return df, alertes


def traiter_propositions(df):
    df = df.copy()

    df.insert(
        0,
        'id_proposition',
        [f'P{i:05d}' for i in range(1, len(df) + 1)]
    )

    df['proposition_originale'] = df['proposition'].map(normaliser_espaces)

    nettoyes = df['proposition_originale'].map(nettoyer)

    df['proposition'] = nettoyes.map(lambda x: x[0])
    df['cle_intitule'] = df['proposition'].map(cle_texte)
    df['proposition_alerte'] = nettoyes.map(lambda x: x[1])

    alertes = []

    for _, ligne in df.iterrows():
        _, remarque = nettoyer(ligne['proposition_originale'])

        if not ligne['proposition']:
            remarque = (remarque + '|valeur_vide').strip('|')

        if remarque:
            alertes.append({
                'origine': 'propositions_anciennes',
                'identifiant': ligne['id_proposition'],
                'fichier_source': ligne.get('fichier_source', ''),
                'champ': 'proposition',
                'valeur_originale': ligne['proposition_originale'],
                'motif': remarque
            })

    return df, alertes


def catalogue(menus, propositions):
    morceaux = []

    correspondance = {
        'entree': 'entree',
        'proteine': 'proteine',
        'feculent': 'feculent',
        'legumes': 'legume',
        'dessert_base': 'dessert',
        'dessert_resto': 'dessert',
        'enfant': 'enfant'
    }

    for col, categorie in correspondance.items():
        part = menus.loc[
            menus[col].ne('')
            & ~menus[col + '_alerte'].str.contains(
                'quantite_seule_a_verifier|jour_service_a_verifier|'
                'mention_service_a_verifier|consigne_a_verifier', regex=True
            ),
            ['id_repas', 'fichier_source', col]
        ].copy()

        part.columns = [
            'identifiant_source',
            'fichier_source',
            'intitule'
        ]

        part['categorie'] = categorie
        part['origine'] = 'menus'

        morceaux.append(part)

    part = propositions.loc[
        propositions['proposition'].ne('')
        & ~propositions['proposition_alerte'].str.contains(
            'quantite_seule_a_verifier|jour_service_a_verifier|'
            'mention_service_a_verifier|consigne_a_verifier', regex=True
        ),
        ['id_proposition', 'fichier_source', 'proposition', 'categorie']
    ].copy()

    part.columns = [
        'identifiant_source',
        'fichier_source',
        'intitule',
        'categorie'
    ]

    part['origine'] = 'propositions_anciennes'
    morceaux.append(part)

    toutes = pd.concat(morceaux, ignore_index=True)
    toutes['cle_intitule'] = toutes['intitule'].map(cle_texte)

    # Garder les categories distinctes :
    # plat != proteine, fruit != dessert.
    cat = (
        toutes.groupby(['categorie', 'cle_intitule'], as_index=False)
        .agg(
            intitule=('intitule', 'first'),
            occurrences=('intitule', 'size'),
            fichiers_distincts=('fichier_source', 'nunique')
        )
    )

    return cat[
        [
            'categorie',
            'intitule',
            'occurrences',
            'fichiers_distincts',
            'cle_intitule'
        ]
    ].sort_values(
        ['categorie', 'occurrences', 'intitule'],
        ascending=[True, False, True]
    )


def main():
    parser = argparse.ArgumentParser(
        description='Nettoyage des menus sans dates ni services'
    )

    parser.add_argument(
        '--entree',
        type=Path,
        default=Path('data/processed'),
        help='Dossier contenant menus.csv et propositions_anciennes.csv'
    )

    parser.add_argument(
        '--sortie',
        type=Path,
        default=Path('data/processed/nettoyage')
    )

    parser.add_argument('--ecraser', action='store_true',
                        help='Allow overwriting previously generated cleaned CSV files')
    args = parser.parse_args()

    if args.entree.resolve() == args.sortie.resolve():
        parser.error(
            'Le dossier de sortie doit etre different du dossier des sources.'
        )

    menus = lire_csv(
        args.entree / 'menus.csv',
        ['fichier_source', 'entree', 'proteine', 'feculent', 'legumes']
    )

    propositions = lire_csv(
        args.entree / 'propositions_anciennes.csv',
        ['fichier_source', 'categorie', 'proposition']
    )

    menus_propres, alertes_m = traiter_menus(menus)
    propositions_propres, alertes_p = traiter_propositions(propositions)

    cat = catalogue(menus_propres, propositions_propres)

    args.sortie.mkdir(parents=True, exist_ok=True)
    existing = [name for name in ('menus_nettoyes.csv', 'propositions_nettoyees.csv',
                'catalogue_plats.csv', 'a_verifier.csv') if (args.sortie / name).exists()]
    if existing and not args.ecraser:
        parser.error('Sortie deja presente : ' + ', '.join(existing)
                     + '. Utiliser un autre --sortie ou --ecraser.')

    sorties = {
        'menus_nettoyes.csv': menus_propres,
        'propositions_nettoyees.csv': propositions_propres,
        'catalogue_plats.csv': cat,
        'a_verifier.csv': pd.DataFrame(
            alertes_m + alertes_p,
            columns=[
                'origine',
                'identifiant',
                'fichier_source',
                'champ',
                'valeur_originale',
                'motif'
            ]
        )
    }

    for nom, table in sorties.items():
        table.to_csv(
            args.sortie / nom,
            sep=';',
            index=False,
            encoding='utf-8-sig'
        )
        print(f'{nom} : {len(table)} lignes')

    print(
        '\nRepas incomplets (0 ou 1 element) :',
        (menus_propres.nb_elements_renseignes <= 1).sum()
    )

    print('Aucun fichier source n a ete modifie.')
    print('Les lignes sources sont conservees; les intitules suspects sont exclus du catalogue uniquement.')


if __name__ == '__main__':
    main()