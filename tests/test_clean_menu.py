"""Unit tests for validated corrections, missing values and RapidFuzz suggestions."""
import csv

import pandas as pd
import pytest

from clean_menu import charger_corrections, nettoyer_avec_regles, suggestions_rapidfuzz


def test_word_and_phrase_corrections():
    rules = {'mot': {'PUREE': 'PURÉE'}, 'expression': {'TARTE POMMES': 'TARTE AUX POMMES'}}
    text, steps = nettoyer_avec_regles('  puree et tarte pommes  ', rules)
    assert text == 'PURÉE ET TARTE AUX POMMES'
    assert {item[0] for item in steps} == {'ORTHOGRAPHE_CORRIGEE', 'EXPRESSION_CORRIGEE'}


def test_halal_removed_and_logged():
    text, steps = nettoyer_avec_regles('Poulet ( HALAL ) rôti', {'mot': {}, 'expression': {}})
    assert text == 'POULET RÔTI'
    assert any(item[0] == 'MENTION_HALAL_SUPPRIMEE' for item in steps)


def test_empty_value_becomes_missing():
    text, _ = nettoyer_avec_regles('(vide)', {'mot': {}, 'expression': {}})
    assert pd.isna(text)


def test_service_words_remain():
    text, _ = nettoyer_avec_regles('Menu lundi soir', {'mot': {}, 'expression': {}})
    assert text == 'MENU LUNDI SOIR'


def test_rapidfuzz_only_suggests():
    suggestions = suggestions_rapidfuzz('ESCALLOPE', ['ESCALOPE'], seuil=80)
    assert suggestions and suggestions[0][0] == 'ESCALLOPE'
    assert suggestions[0][1] == 'ESCALOPE'


def test_conflicting_corrections_raise_error(tmp_path):
    rules = tmp_path / 'corrections.csv'
    rules.write_text('type;incorrect;correct\nmot;TEST;A\nmot;TEST;B\n', encoding='utf-8')
    with pytest.raises(ValueError, match='contradictoires'):
        charger_corrections(rules)


def test_valid_correction_csv(tmp_path):
    rules = tmp_path / 'corrections.csv'
    rules.write_text('type;incorrect;correct\nmot;PUREE;PURÉE\n', encoding='utf-8')
    assert charger_corrections(rules)['mot']['PUREE'] == 'PURÉE'
