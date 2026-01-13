"""
Testes unitários para o módulo de configuração.
"""

import copy
import json
import os
import pytest

from src.config import DEFAULT_CONFIG, load_config, save_config, get_config_path


class TestDefaultConfig:
    """Testes para a configuração padrão."""
    
    def test_default_config_has_required_sections(self):
        """Verifica que DEFAULT_CONFIG tem todas as secções necessárias."""
        required_sections = ['pdf', 'header', 'colors', 'table', 'footer', 'output', 'contabilidade', 'banking']
        for section in required_sections:
            assert section in DEFAULT_CONFIG, f"Secção '{section}' em falta no DEFAULT_CONFIG"
    
    def test_pdf_section_has_required_keys(self):
        """Verifica que a secção PDF tem as chaves necessárias."""
        required_keys = ['page_size', 'orientation', 'margin_top', 'margin_bottom', 'margin_left', 'margin_right']
        for key in required_keys:
            assert key in DEFAULT_CONFIG['pdf'], f"Chave '{key}' em falta em DEFAULT_CONFIG['pdf']"
    
    def test_banking_section_has_required_keys(self):
        """Verifica que a secção banking tem as chaves necessárias."""
        required_keys = ['show_banking', 'bank_name', 'iban']
        for key in required_keys:
            assert key in DEFAULT_CONFIG['banking'], f"Chave '{key}' em falta em DEFAULT_CONFIG['banking']"
    
    def test_default_banking_values(self):
        """Verifica os valores default dos dados bancários."""
        assert DEFAULT_CONFIG['banking']['bank_name'] == 'ABANCA'
        assert 'PT50' in DEFAULT_CONFIG['banking']['iban']


class TestLoadConfig:
    """Testes para a função load_config."""
    
    def test_load_config_returns_dict(self):
        """Verifica que load_config retorna um dicionário."""
        config = load_config()
        assert isinstance(config, dict)
    
    def test_load_config_returns_deepcopy(self, sample_config):
        """Verifica que load_config retorna uma cópia profunda."""
        config1 = load_config()
        config2 = load_config()
        
        # Modificar config1 não deve afetar config2
        config1['pdf']['page_size'] = 'Letter'
        assert config2['pdf']['page_size'] == 'A4'
    
    def test_load_config_has_all_sections(self):
        """Verifica que load_config retorna todas as secções."""
        config = load_config()
        for section in DEFAULT_CONFIG.keys():
            assert section in config


class TestSaveConfig:
    """Testes para a função save_config."""
    
    def test_save_config_creates_file(self, temp_dir, monkeypatch):
        """Verifica que save_config cria o ficheiro."""
        config_path = os.path.join(temp_dir, 'config.json')
        
        # Monkey-patch get_config_path para usar o caminho temporário
        monkeypatch.setattr('src.config.get_config_path', lambda: config_path)
        
        config = copy.deepcopy(DEFAULT_CONFIG)
        result = save_config(config)
        
        assert result == True
        assert os.path.exists(config_path)
    
    def test_save_and_load_roundtrip(self, temp_dir, monkeypatch):
        """Verifica que guardar e carregar preserva os dados."""
        config_path = os.path.join(temp_dir, 'config.json')
        monkeypatch.setattr('src.config.get_config_path', lambda: config_path)
        
        # Modificar config e guardar
        config = copy.deepcopy(DEFAULT_CONFIG)
        config['banking']['bank_name'] = 'CGD'
        config['banking']['iban'] = 'PT50 0035 0000 0000 0000 0000 0'
        
        save_config(config)
        loaded_config = load_config()
        
        assert loaded_config['banking']['bank_name'] == 'CGD'
        assert loaded_config['banking']['iban'] == 'PT50 0035 0000 0000 0000 0000 0'


class TestDeepCopyPrevention:
    """Testes para verificar que deepcopy previne mutação."""
    
    def test_modifying_returned_config_does_not_affect_default(self):
        """Verifica que modificar config retornado não afeta DEFAULT_CONFIG."""
        original_bank = DEFAULT_CONFIG['banking']['bank_name']
        
        config = load_config()
        config['banking']['bank_name'] = 'OUTRO_BANCO'
        
        # DEFAULT_CONFIG deve manter o valor original
        assert DEFAULT_CONFIG['banking']['bank_name'] == original_bank
