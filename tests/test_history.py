"""
Testes unitários para o módulo de histórico.
"""

import pytest

from src import database as db
from src import history


@pytest.fixture(autouse=True)
def isolated_db(tmp_path, monkeypatch):
    """Redireciona a base de dados para um ficheiro temporário por teste."""
    db_path = str(tmp_path / 'test.db')
    monkeypatch.setattr('src.database._get_db_path', lambda: db_path)
    db.init_db()


class TestHistory:
    """Testes para o módulo de histórico (wrapper sobre database)."""

    def test_add_entry_and_get(self):
        """Verifica que uma entrada adicionada é recuperada."""
        history.add_entry('/path/file.xlsx', '/out/file.pdf', 'aggregate', 3, True)
        entries = history.get_history()
        assert len(entries) == 1
        assert entries[0]['source_file'] == 'file.xlsx'
        assert entries[0]['mode'] == 'aggregate'

    def test_add_entry_with_error(self):
        """Verifica que uma entrada com erro é registada."""
        history.add_entry('/path/file.xlsx', '', 'individual', 0, False, 'Ficheiro corrompido')
        entries = history.get_history()
        assert entries[0]['success'] is False
        assert entries[0]['error'] == 'Ficheiro corrompido'

    def test_get_history_with_limit(self):
        """Verifica que o limite funciona."""
        for i in range(10):
            history.add_entry(f'/path/f{i}.xlsx', f'/out/{i}.pdf', 'aggregate', 1, True)
        entries = history.get_history(limit=5)
        assert len(entries) == 5

    def test_clear_history(self):
        """Verifica que clear_history apaga tudo."""
        history.add_entry('/path/file.xlsx', '/out/file.pdf', 'aggregate', 1, True)
        history.clear_history()
        entries = history.get_history()
        assert len(entries) == 0
