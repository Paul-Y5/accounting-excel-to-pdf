"""
Testes unitários para get_history_filtered em database.py.
"""

import pytest
from src.database import add_history_entry, get_history_filtered


def _populate(isolated_db):
    """Insere dados de teste no DB isolado."""
    add_history_entry('jan_relatorio.xlsx', '/out/jan.pdf', 'aggregate', 5, True)
    add_history_entry('jan_errado.xlsx', '', 'aggregate', 0, False, 'Ficheiro inválido')
    add_history_entry('fev_relatorio.xlsx', '/out/fev.pdf', 'individual', 3, True)


class TestHistoryFiltered:
    """Testes para get_history_filtered."""

    def test_no_filters_returns_all(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(limit=100)
        assert len(entries) == 3

    def test_filter_success_only(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(success_only=True)
        assert all(e['success'] for e in entries)
        assert len(entries) == 2

    def test_filter_errors_only(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(success_only=False)
        assert all(not e['success'] for e in entries)
        assert len(entries) == 1

    def test_search_by_filename_partial(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(search_term='jan')
        assert len(entries) == 2
        for e in entries:
            assert 'jan' in e['source_file']

    def test_search_no_match_returns_empty(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(search_term='nao_existe_xyz')
        assert entries == []

    def test_limit_respected(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(limit=1)
        assert len(entries) == 1

    def test_combine_search_and_success(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(search_term='jan', success_only=True)
        assert len(entries) == 1
        assert entries[0]['source_file'] == 'jan_relatorio.xlsx'

    def test_combine_search_and_error(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(search_term='jan', success_only=False)
        assert len(entries) == 1
        assert entries[0]['source_file'] == 'jan_errado.xlsx'

    def test_returns_most_recent_first(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(limit=100)
        assert 'fev' in entries[0]['source_file']

    def test_result_has_required_keys(self, isolated_db):
        _populate(isolated_db)
        entries = get_history_filtered(limit=1)
        required = {'timestamp', 'source_file', 'output_path', 'mode', 'clients_count', 'success'}
        assert required.issubset(entries[0].keys())
