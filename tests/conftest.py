"""
Fixtures partilhadas para testes.
"""

import copy
import pytest
import tempfile
import os

from src.config import DEFAULT_CONFIG


@pytest.fixture
def sample_config():
    """Retorna uma cópia do DEFAULT_CONFIG para testes."""
    return copy.deepcopy(DEFAULT_CONFIG)


@pytest.fixture
def temp_dir():
    """Cria um diretório temporário para testes."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


@pytest.fixture
def temp_config_file(temp_dir):
    """Cria um ficheiro de configuração temporário."""
    config_path = os.path.join(temp_dir, 'config.json')
    yield config_path
    # Limpar após o teste
    if os.path.exists(config_path):
        os.remove(config_path)
