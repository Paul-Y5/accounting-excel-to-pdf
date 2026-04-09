"""
Testes para o módulo de gestão de fontes (feature/v3.3-remaining).

Valida que:
- register_font devolve False para caminho inválido
- register_font regista fonte válida via mock
- load_fonts_from_config carrega fontes da configuração
- get_body_font e get_header_font devolvem valores corretos
"""
import copy
import pytest
from unittest.mock import patch, MagicMock

from src.config import DEFAULT_CONFIG
from src.font_manager import register_font, load_fonts_from_config, get_body_font, get_header_font


class TestRegisterFont:
    def test_invalid_path_returns_false(self):
        assert register_font('Test', '/caminho/inexistente/fonte.ttf') is False

    def test_empty_path_returns_false(self):
        assert register_font('Test', '') is False

    def test_none_path_returns_false(self):
        assert register_font('Test', None) is False

    @patch('src.font_manager.pdfmetrics')
    def test_valid_font_calls_register(self, mock_metrics, tmp_path):
        # Criar ficheiro .ttf fictício
        font_file = tmp_path / 'test.ttf'
        font_file.write_bytes(b'\x00' * 100)

        with patch('src.font_manager.TTFont') as mock_ttfont:
            result = register_font('TestFont', str(font_file))

        # Se TTFont lançar exceção (ficheiro não é TTF válido), retorna False
        # Mas o register_font tenta — aqui mockamos para sucesso
        mock_ttfont.assert_called_once()
        mock_metrics.registerFont.assert_called_once()
        assert result is True

    @patch('src.font_manager.pdfmetrics')
    def test_exception_returns_false(self, mock_metrics, tmp_path):
        font_file = tmp_path / 'bad.ttf'
        font_file.write_bytes(b'\x00')

        with patch('src.font_manager.TTFont', side_effect=Exception('bad font')):
            result = register_font('BadFont', str(font_file))

        assert result is False


class TestLoadFontsFromConfig:
    def test_empty_registered_returns_empty(self):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        result = load_fonts_from_config(cfg)
        assert result == []

    @patch('src.font_manager.register_font', return_value=True)
    def test_loads_registered_fonts(self, mock_reg):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        cfg['fonts']['registered'] = [
            {'name': 'FontA', 'path': '/a.ttf'},
            {'name': 'FontB', 'path': '/b.ttf'},
        ]
        result = load_fonts_from_config(cfg)
        assert result == ['FontA', 'FontB']
        assert mock_reg.call_count == 2

    @patch('src.font_manager.register_font', return_value=False)
    def test_failed_fonts_excluded(self, mock_reg):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        cfg['fonts']['registered'] = [{'name': 'Bad', 'path': '/x.ttf'}]
        result = load_fonts_from_config(cfg)
        assert result == []


class TestFontGetters:
    def test_body_font_default(self):
        assert get_body_font(DEFAULT_CONFIG) == 'Helvetica'

    def test_header_font_default(self):
        assert get_header_font(DEFAULT_CONFIG) == 'Helvetica-Bold'

    def test_body_font_custom(self):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        cfg['fonts']['body_font'] = 'MyFont'
        assert get_body_font(cfg) == 'MyFont'

    def test_header_font_custom(self):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        cfg['fonts']['header_font'] = 'MyBoldFont'
        assert get_header_font(cfg) == 'MyBoldFont'

    def test_no_fonts_section_uses_defaults(self):
        assert get_body_font({}) == 'Helvetica'
        assert get_header_font({}) == 'Helvetica-Bold'
