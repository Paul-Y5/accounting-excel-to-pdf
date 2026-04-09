"""
Testes para o módulo de geração de QR Code (feature/v3.3-remaining).

Valida que:
- get_qr_data devolve NIF ou IBAN conforme configuração
- build_qr_image cria ficheiro PNG temporário
- dados vazios lançam ValueError
- qrcode não instalado lança ImportError (mockado)
"""
import copy
import os
import pytest
from unittest.mock import patch, MagicMock

from src.config import DEFAULT_CONFIG
from src.qr_generator import get_qr_data, build_qr_image


@pytest.fixture
def config_nif():
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    cfg['header']['company_nif'] = 'PT 500 000 000'
    cfg['qrcode'] = {'enabled': True, 'content': 'nif', 'size_mm': 25}
    return cfg


@pytest.fixture
def config_iban():
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    cfg['banking']['accounts'] = [{'bank_name': 'TEST', 'iban': 'PT50 1234 5678 9012 3456 7890 1', 'default': True}]
    cfg['qrcode'] = {'enabled': True, 'content': 'iban', 'size_mm': 25}
    return cfg


class TestGetQrData:
    def test_nif_content(self, config_nif):
        result = get_qr_data(config_nif)
        assert result == 'PT500000000'

    def test_iban_content(self, config_iban):
        result = get_qr_data(config_iban)
        assert result == 'PT50123456789012345678901'

    def test_default_is_nif(self):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        cfg['header']['company_nif'] = 'PT123'
        result = get_qr_data(cfg)
        assert result == 'PT123'

    def test_empty_nif_returns_empty(self):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        cfg['header']['company_nif'] = ''
        result = get_qr_data(cfg)
        assert result == ''


class TestBuildQrImage:
    def test_creates_png_file(self):
        path = build_qr_image('PT500000000', size_mm=25)
        try:
            assert os.path.isfile(path)
            assert path.endswith('.png')
            assert os.path.getsize(path) > 0
        finally:
            os.remove(path)

    def test_empty_data_raises(self):
        with pytest.raises(ValueError):
            build_qr_image('')

    def test_none_data_raises(self):
        with pytest.raises(ValueError):
            build_qr_image(None)

    def test_different_sizes(self):
        for size in [10, 25, 50]:
            path = build_qr_image('TEST', size_mm=size)
            try:
                assert os.path.isfile(path)
            finally:
                os.remove(path)

    def test_import_error_when_missing(self):
        with patch.dict('sys.modules', {'qrcode': None}):
            with pytest.raises(ImportError):
                build_qr_image('TEST')
