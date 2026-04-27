#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Testes para o módulo src/doc_sequence.py."""

import pytest

from src.doc_sequence import (
    delete_serie,
    get_next_number,
    init_doc_sequences_table,
    list_series,
    peek_next_number,
    reset_serie,
    upsert_serie,
)


@pytest.fixture(autouse=True)
def _clean_db(tmp_path, monkeypatch):
    """Usa uma BD temporária isolada para cada teste."""
    import src.database as db_mod

    db_file = str(tmp_path / 'test.db')
    monkeypatch.setattr(db_mod, '_get_db_path', lambda: db_file)
    from src.database import init_db
    init_db()
    yield


class TestPeekNextNumber:
    def test_returns_0001_for_unknown_serie(self):
        result = peek_next_number('FT')
        assert result.endswith('/0001')
        assert result.startswith('FT ')

    def test_increments_after_upsert(self):
        upsert_serie('FT', ultimo_numero=5, ano=2026)
        assert peek_next_number('FT') == 'FT 2026/0006'

    def test_does_not_change_counter(self):
        upsert_serie('FT', ultimo_numero=3, ano=2026)
        peek_next_number('FT')
        peek_next_number('FT')
        assert peek_next_number('FT') == 'FT 2026/0004'


class TestGetNextNumber:
    def test_first_number_is_0001(self):
        upsert_serie('FR', ano=2026)
        assert get_next_number('FR') == 'FR 2026/0001'

    def test_increments_sequentially(self):
        upsert_serie('FT', ano=2026)
        nums = [get_next_number('FT') for _ in range(3)]
        assert nums == ['FT 2026/0001', 'FT 2026/0002', 'FT 2026/0003']

    def test_creates_serie_if_absent(self):
        result = get_next_number('REC')
        assert 'REC' in result
        assert '/0001' in result

    def test_anno_reset_on_new_year(self, monkeypatch):
        import src.doc_sequence as ds
        from datetime import datetime

        # Simular ano 2025
        class _FakeDate2025:
            @staticmethod
            def now():
                return datetime(2025, 12, 31)

        monkeypatch.setattr(ds, 'datetime', _FakeDate2025)
        upsert_serie('FT', ano=2025, ultimo_numero=10)
        get_next_number('FT')  # 2025/0011

        # Avançar para 2026
        class _FakeDate2026:
            @staticmethod
            def now():
                return datetime(2026, 1, 1)

        monkeypatch.setattr(ds, 'datetime', _FakeDate2026)
        result = get_next_number('FT')
        assert result == 'FT 2026/0001'

    def test_no_reset_when_disabled(self, monkeypatch):
        import src.doc_sequence as ds
        from datetime import datetime

        class _FakeDate2025:
            @staticmethod
            def now():
                return datetime(2025, 6, 1)

        monkeypatch.setattr(ds, 'datetime', _FakeDate2025)
        upsert_serie('ND', ano=2025, ultimo_numero=50, reset_anual=False)
        get_next_number('ND')  # 2025/0051

        class _FakeDate2026:
            @staticmethod
            def now():
                return datetime(2026, 1, 1)

        monkeypatch.setattr(ds, 'datetime', _FakeDate2026)
        result = get_next_number('ND')
        assert result == 'ND 2025/0052'


class TestUpsertSerie:
    def test_create_new(self):
        upsert_serie('NC', ultimo_numero=0, ano=2026)
        series = {s['serie']: s for s in list_series()}
        assert 'NC' in series
        assert series['NC']['ultimo_numero'] == 0

    def test_update_existing(self):
        upsert_serie('FT', ultimo_numero=10, ano=2026)
        upsert_serie('FT', ultimo_numero=99, ano=2026)
        series = {s['serie']: s for s in list_series()}
        assert series['FT']['ultimo_numero'] == 99


class TestListSeries:
    def test_empty(self):
        assert list_series() == []

    def test_multiple_series_sorted(self):
        upsert_serie('FR', ano=2026)
        upsert_serie('FT', ano=2026)
        upsert_serie('REC', ano=2026)
        codes = [s['serie'] for s in list_series()]
        assert codes == sorted(codes)

    def test_includes_proximo(self):
        upsert_serie('FT', ultimo_numero=0, ano=2026)
        series = list_series()
        assert series[0]['proximo'] == 'FT 2026/0001'


class TestResetSerie:
    def test_resets_counter_to_zero(self):
        upsert_serie('FT', ultimo_numero=50, ano=2026)
        reset_serie('FT', ano=2026)
        assert peek_next_number('FT') == 'FT 2026/0001'

    def test_noop_on_unknown_serie(self):
        reset_serie('XX')  # deve não lançar excepção


class TestDeleteSerie:
    def test_removes_serie(self):
        upsert_serie('FT', ano=2026)
        delete_serie('FT')
        codes = [s['serie'] for s in list_series()]
        assert 'FT' not in codes

    def test_noop_on_unknown_serie(self):
        delete_serie('NAOEXISTE')  # deve não lançar excepção
