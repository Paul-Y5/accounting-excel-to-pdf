"""
Testes unitários para o módulo de base de dados SQLite.
"""

import json
import os
import pytest

from src.config import DEFAULT_CONFIG
from src import database as db


@pytest.fixture(autouse=True)
def isolated_db(tmp_path, monkeypatch):
    """Redireciona a base de dados para um ficheiro temporário por teste."""
    db_path = str(tmp_path / 'test.db')
    monkeypatch.setattr('src.database._get_db_path', lambda: db_path)
    db.init_db()
    return db_path


# ============================================
# INIT
# ============================================

class TestInitDb:
    """Testes para a inicialização da base de dados."""

    def test_init_creates_tables(self, isolated_db):
        """Verifica que init_db cria as tabelas necessárias."""
        conn = db._get_connection()
        try:
            cursor = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
            )
            tables = {row['name'] for row in cursor.fetchall()}
            assert 'history' in tables
            assert 'profiles' in tables
            assert 'client_cache' in tables
        finally:
            conn.close()

    def test_init_is_idempotent(self, isolated_db):
        """Verifica que init_db pode ser chamado múltiplas vezes."""
        db.init_db()
        db.init_db()
        # Não deve lançar exceção


# ============================================
# HISTÓRICO
# ============================================

class TestHistory:
    """Testes para funções de histórico."""

    def test_add_and_get_history(self):
        """Verifica que uma entrada adicionada aparece no histórico."""
        db.add_history_entry('/path/to/file.xlsx', '/output/file.pdf',
                             'aggregate', 5, True)
        history = db.get_history()
        assert len(history) == 1
        assert history[0]['source_file'] == 'file.xlsx'
        assert history[0]['source_path'] == '/path/to/file.xlsx'
        assert history[0]['output_path'] == '/output/file.pdf'
        assert history[0]['mode'] == 'aggregate'
        assert history[0]['clients_count'] == 5
        assert history[0]['success'] is True
        assert history[0]['error'] == ''

    def test_add_failed_entry(self):
        """Verifica que uma entrada falhada é registada corretamente."""
        db.add_history_entry('/path/file.xlsx', '', 'individual', 0, False, 'Erro de leitura')
        history = db.get_history()
        assert len(history) == 1
        assert history[0]['success'] is False
        assert history[0]['error'] == 'Erro de leitura'

    def test_history_order_most_recent_first(self):
        """Verifica que o histórico é ordenado por mais recente primeiro."""
        db.add_history_entry('a.xlsx', '/out/a.pdf', 'aggregate', 1, True)
        db.add_history_entry('b.xlsx', '/out/b.pdf', 'individual', 2, True)
        history = db.get_history()
        assert history[0]['source_file'] == 'b.xlsx'
        assert history[1]['source_file'] == 'a.xlsx'

    def test_history_limit(self):
        """Verifica que o limite de entradas funciona."""
        for i in range(5):
            db.add_history_entry(f'file{i}.xlsx', f'/out/{i}.pdf', 'aggregate', 1, True)
        history = db.get_history(limit=3)
        assert len(history) == 3

    def test_clear_history(self):
        """Verifica que clear_history apaga tudo."""
        db.add_history_entry('file.xlsx', '/out/file.pdf', 'aggregate', 1, True)
        db.clear_history()
        history = db.get_history()
        assert len(history) == 0

    def test_history_auto_prune(self):
        """Verifica que o histórico é limitado a 500 entradas."""
        for i in range(505):
            db.add_history_entry(f'file{i}.xlsx', f'/out/{i}.pdf', 'aggregate', 1, True)
        history = db.get_history(limit=1000)
        assert len(history) <= 500


# ============================================
# PERFIS
# ============================================

class TestProfiles:
    """Testes para funções de perfis."""

    def test_list_profiles_empty(self):
        """Verifica que lista vazia é retornada quando não há perfis."""
        assert db.list_profiles_db() == []

    def test_save_and_list_profile(self):
        """Verifica que um perfil guardado aparece na lista."""
        config = {'pdf': {'page_size': 'A4'}}
        assert db.save_profile_db('teste', config) is True
        profiles = db.list_profiles_db()
        assert 'teste' in profiles

    def test_save_and_load_profile(self):
        """Verifica que um perfil guardado pode ser carregado."""
        config = {'pdf': {'page_size': 'Letter'}}
        db.save_profile_db('meu_perfil', config)
        loaded = db.load_profile_db('meu_perfil')
        assert loaded is not None
        assert loaded['pdf']['page_size'] == 'Letter'

    def test_load_profile_merges_with_defaults(self):
        """Verifica que carregar perfil faz merge com DEFAULT_CONFIG."""
        config = {'pdf': {'page_size': 'Letter'}}
        db.save_profile_db('parcial', config)
        loaded = db.load_profile_db('parcial')
        # Deve ter as secções do DEFAULT_CONFIG
        assert 'header' in loaded
        assert 'colors' in loaded
        assert 'banking' in loaded

    def test_load_nonexistent_profile(self):
        """Verifica que carregar perfil inexistente retorna None."""
        assert db.load_profile_db('nao_existe') is None

    def test_update_existing_profile(self):
        """Verifica que guardar com mesmo nome atualiza o perfil."""
        db.save_profile_db('perfil', {'pdf': {'page_size': 'A4'}})
        db.save_profile_db('perfil', {'pdf': {'page_size': 'Letter'}})
        profiles = db.list_profiles_db()
        assert profiles.count('perfil') == 1
        loaded = db.load_profile_db('perfil')
        assert loaded['pdf']['page_size'] == 'Letter'

    def test_delete_profile(self):
        """Verifica que um perfil pode ser apagado."""
        db.save_profile_db('para_apagar', {'pdf': {}})
        result = db.delete_profile_db('para_apagar')
        assert result is True
        assert 'para_apagar' not in db.list_profiles_db()

    def test_delete_nonexistent_profile(self):
        """Verifica que apagar perfil inexistente retorna False."""
        result = db.delete_profile_db('fantasma')
        assert result is False

    def test_profiles_sorted_by_name(self):
        """Verifica que os perfis são listados por ordem alfabética."""
        db.save_profile_db('zebra', {'pdf': {}})
        db.save_profile_db('alfa', {'pdf': {}})
        db.save_profile_db('meio', {'pdf': {}})
        profiles = db.list_profiles_db()
        assert profiles == sorted(profiles)


# ============================================
# CACHE DE CLIENTES
# ============================================

class TestClientCache:
    """Testes para funções de cache de clientes."""

    def test_update_and_get_clients(self):
        """Verifica que clientes são adicionados e recuperados."""
        clients = [
            {'name': 'Empresa A', 'sigla': 'EA', 'nif': '123456789'},
            {'name': 'Empresa B', 'sigla': 'EB', 'nif': ''},
        ]
        db.update_client_cache('dados.xlsx', clients)
        cached = db.get_cached_clients('dados.xlsx')
        assert len(cached) == 2
        names = {c['name'] for c in cached}
        assert 'Empresa A' in names
        assert 'Empresa B' in names

    def test_update_increments_count(self):
        """Verifica que atualizar o mesmo cliente incrementa o contador."""
        clients = [{'name': 'Empresa A', 'sigla': 'EA', 'nif': '123456789'}]
        db.update_client_cache('dados.xlsx', clients)
        db.update_client_cache('dados.xlsx', clients)
        cached = db.get_cached_clients('dados.xlsx')
        assert cached[0]['conversion_count'] == 2

    def test_empty_name_is_skipped(self):
        """Verifica que clientes sem nome são ignorados."""
        clients = [{'name': '', 'sigla': 'X', 'nif': ''}]
        db.update_client_cache('dados.xlsx', clients)
        cached = db.get_cached_clients('dados.xlsx')
        assert len(cached) == 0

    def test_get_all_clients(self):
        """Verifica que sem filtro retorna todos os clientes."""
        db.update_client_cache('a.xlsx', [{'name': 'A'}])
        db.update_client_cache('b.xlsx', [{'name': 'B'}])
        cached = db.get_cached_clients()
        assert len(cached) == 2

    def test_get_clients_filtered_by_source(self):
        """Verifica filtro por ficheiro de origem."""
        db.update_client_cache('a.xlsx', [{'name': 'A'}])
        db.update_client_cache('b.xlsx', [{'name': 'B'}])
        cached = db.get_cached_clients('a.xlsx')
        assert len(cached) == 1
        assert cached[0]['name'] == 'A'

    def test_clear_client_cache(self):
        """Verifica que a cache pode ser limpa."""
        db.update_client_cache('dados.xlsx', [{'name': 'Teste'}])
        db.clear_client_cache()
        cached = db.get_cached_clients()
        assert len(cached) == 0

    def test_nif_not_overwritten_with_empty(self):
        """Verifica que um NIF existente não é substituído por vazio."""
        db.update_client_cache('dados.xlsx', [{'name': 'A', 'nif': '123456789'}])
        db.update_client_cache('dados.xlsx', [{'name': 'A', 'nif': ''}])
        cached = db.get_cached_clients('dados.xlsx')
        assert cached[0]['nif'] == '123456789'

    def test_client_result_structure(self):
        """Verifica que o resultado tem as chaves esperadas."""
        db.update_client_cache('dados.xlsx', [{'name': 'Empresa'}])
        cached = db.get_cached_clients('dados.xlsx')
        entry = cached[0]
        expected_keys = {'source_file', 'name', 'sigla', 'nif', 'last_seen', 'conversion_count'}
        assert expected_keys == set(entry.keys())


# ============================================
# MIGRAÇÃO
# ============================================

class TestMigration:
    """Testes para a migração JSON → SQLite."""

    def test_migrate_history_from_json(self, tmp_path, monkeypatch):
        """Verifica que o histórico JSON é migrado para SQLite."""
        config_dir = str(tmp_path / 'config')
        os.makedirs(config_dir, exist_ok=True)
        monkeypatch.setattr('src.database.get_config_dir', lambda: config_dir)

        history_data = [
            {
                'timestamp': '2025-01-01T10:00:00',
                'source_file': 'test.xlsx',
                'source_path': '/path/test.xlsx',
                'output_path': '/out/test.pdf',
                'mode': 'aggregate',
                'clients_count': 3,
                'success': True,
                'error': '',
            }
        ]
        history_path = os.path.join(config_dir, 'history.json')
        with open(history_path, 'w', encoding='utf-8') as f:
            json.dump(history_data, f)

        db.migrate_from_json()

        # Verificar que os dados foram migrados
        history = db.get_history()
        assert len(history) == 1
        assert history[0]['source_file'] == 'test.xlsx'

        # Verificar que o ficheiro foi renomeado
        assert not os.path.exists(history_path)
        assert os.path.exists(history_path + '.bak')

    def test_migrate_profiles_from_json(self, tmp_path, monkeypatch):
        """Verifica que os perfis JSON são migrados para SQLite."""
        config_dir = str(tmp_path / 'config')
        profiles_dir = os.path.join(config_dir, 'profiles')
        os.makedirs(profiles_dir, exist_ok=True)
        monkeypatch.setattr('src.database.get_config_dir', lambda: config_dir)

        profile_data = {'pdf': {'page_size': 'Letter'}}
        with open(os.path.join(profiles_dir, 'meu_perfil.json'), 'w', encoding='utf-8') as f:
            json.dump(profile_data, f)

        db.migrate_from_json()

        profiles = db.list_profiles_db()
        assert 'meu_perfil' in profiles

        # Verificar que a pasta foi renomeada
        assert not os.path.isdir(profiles_dir)
        assert os.path.isdir(profiles_dir + '.bak')

    def test_migrate_no_files(self, tmp_path, monkeypatch):
        """Verifica que migração sem ficheiros não causa erro."""
        config_dir = str(tmp_path / 'config')
        os.makedirs(config_dir, exist_ok=True)
        monkeypatch.setattr('src.database.get_config_dir', lambda: config_dir)
        db.migrate_from_json()  # Não deve lançar exceção
