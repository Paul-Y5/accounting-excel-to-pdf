"""
Testes unitários para export_config e import_config.
"""

import copy
import json
import os
import pytest

from src.config import DEFAULT_CONFIG, export_config, import_config


class TestExportConfig:
    """Testes para export_config."""

    def test_creates_json_file(self, temp_dir):
        path = os.path.join(temp_dir, 'export.json')
        config = copy.deepcopy(DEFAULT_CONFIG)
        result = export_config(config, path)
        assert result is True
        assert os.path.exists(path)

    def test_exported_file_is_valid_json(self, temp_dir):
        path = os.path.join(temp_dir, 'export.json')
        config = copy.deepcopy(DEFAULT_CONFIG)
        export_config(config, path)
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        assert isinstance(data, dict)

    def test_exported_values_match_input(self, temp_dir):
        path = os.path.join(temp_dir, 'export.json')
        config = copy.deepcopy(DEFAULT_CONFIG)
        config['header']['company_name'] = 'Empresa Exportada'
        export_config(config, path)
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        assert data['header']['company_name'] == 'Empresa Exportada'

    def test_returns_false_on_invalid_path(self):
        result = export_config({}, '/caminho/inexistente/pasta/export.json')
        assert result is False


class TestImportConfig:
    """Testes para import_config."""

    def test_raises_on_missing_file(self, temp_dir):
        with pytest.raises(FileNotFoundError):
            import_config(os.path.join(temp_dir, 'nao_existe.json'))

    def test_raises_on_invalid_json(self, temp_dir):
        path = os.path.join(temp_dir, 'invalid.json')
        with open(path, 'w') as f:
            f.write('isto nao e json {{{')
        with pytest.raises(ValueError):
            import_config(path)

    def test_returns_dict(self, temp_dir):
        path = os.path.join(temp_dir, 'config.json')
        export_config(copy.deepcopy(DEFAULT_CONFIG), path)
        result = import_config(path)
        assert isinstance(result, dict)

    def test_merges_with_defaults(self, temp_dir):
        path = os.path.join(temp_dir, 'partial.json')
        partial = {'header': {'company_name': 'Importada'}}
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(partial, f)
        result = import_config(path)
        assert result['header']['company_name'] == 'Importada'
        assert 'pdf' in result
        assert result['pdf']['page_size'] == DEFAULT_CONFIG['pdf']['page_size']

    def test_unknown_keys_ignored(self, temp_dir):
        path = os.path.join(temp_dir, 'unknown.json')
        data = {'chave_desconhecida': {'algo': 1}}
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f)
        result = import_config(path)
        assert 'chave_desconhecida' not in result

    def test_roundtrip_preserves_values(self, temp_dir):
        path = os.path.join(temp_dir, 'roundtrip.json')
        config = copy.deepcopy(DEFAULT_CONFIG)
        config['header']['company_name'] = 'Roundtrip Lda'
        config['ui']['theme'] = 'dark'
        export_config(config, path)
        result = import_config(path)
        assert result['header']['company_name'] == 'Roundtrip Lda'
        assert result['ui']['theme'] == 'dark'
