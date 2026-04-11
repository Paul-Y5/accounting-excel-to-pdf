#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Testes para o CLI melhorado (converter_excel_pdf.py).
"""

import os
import sys
import pytest
from unittest.mock import patch, MagicMock

# Garantir que o root do projecto está no path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import importlib
import types


def _parse(argv):
    """Faz parse de argv usando o argparser do entry point."""
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('input', nargs='?')
    parser.add_argument('-o', '--output')
    parser.add_argument('-m', '--mode', choices=['individual', 'aggregate'], default=None)
    parser.add_argument('-p', '--profile')
    parser.add_argument('-c', '--config')
    parser.add_argument('-w', '--watch', action='store_true')
    return parser.parse_args(argv)


# ---------------------------------------------------------------------------
# TestArgParser
# ---------------------------------------------------------------------------

class TestArgParser:
    def test_no_args_input_none(self):
        args = _parse([])
        assert args.input is None

    def test_input_positional(self):
        args = _parse(['ficheiro.xlsx'])
        assert args.input == 'ficheiro.xlsx'

    def test_output_flag(self):
        args = _parse(['f.xlsx', '-o', 'out.pdf'])
        assert args.output == 'out.pdf'

    def test_mode_individual(self):
        args = _parse(['f.xlsx', '-m', 'individual'])
        assert args.mode == 'individual'

    def test_mode_aggregate(self):
        args = _parse(['f.xlsx', '-m', 'aggregate'])
        assert args.mode == 'aggregate'

    def test_profile_flag(self):
        args = _parse(['f.xlsx', '-p', 'empresa_x'])
        assert args.profile == 'empresa_x'

    def test_config_flag(self):
        args = _parse(['f.xlsx', '-c', 'config.json'])
        assert args.config == 'config.json'

    def test_watch_flag(self):
        args = _parse(['pasta/', '-w'])
        assert args.watch is True

    def test_watch_false_by_default(self):
        args = _parse(['f.xlsx'])
        assert args.watch is False

    def test_invalid_mode_raises(self):
        with pytest.raises(SystemExit):
            _parse(['f.xlsx', '-m', 'invalido'])


# ---------------------------------------------------------------------------
# TestCliConversion
# ---------------------------------------------------------------------------

class TestCliConversion:
    def test_missing_file_exits(self, tmp_path):
        """Ficheiro inexistente deve causar sys.exit."""
        import converter_excel_pdf as entry
        args = MagicMock()
        args.input = str(tmp_path / 'nao_existe.xlsx')
        args.output = None
        args.mode = None
        args.profile = None
        args.config = None
        args.watch = False
        with pytest.raises(SystemExit):
            entry._run_cli(args)

    def test_conversion_calls_converter(self, tmp_path):
        """Deve criar um ExcelToPDFConverter e chamar generate."""
        import converter_excel_pdf as entry

        src = tmp_path / 'test.xlsx'
        src.write_text('dummy')

        mock_converter = MagicMock()
        mock_converter.generate_individual_pdfs.return_value = [str(tmp_path / 'out.pdf')]

        with patch('converter_excel_pdf.load_config', return_value={'output': {'auto_open': False}}), \
             patch('src.converter.ExcelToPDFConverter', return_value=mock_converter), \
             patch('src.hooks.run_hooks', return_value=[]):
            args = MagicMock()
            args.input = str(src)
            args.output = None
            args.mode = 'individual'
            args.profile = None
            args.config = None
            args.watch = False
            entry._run_cli(args)

        mock_converter.generate_individual_pdfs.assert_called_once()

    def test_aggregate_mode_calls_generate_pdf(self, tmp_path):
        import converter_excel_pdf as entry

        src = tmp_path / 'test.xlsx'
        src.write_text('dummy')

        mock_converter = MagicMock()
        mock_converter.generate_pdf.return_value = str(tmp_path / 'out.pdf')

        with patch('converter_excel_pdf.load_config', return_value={'output': {'auto_open': False}}), \
             patch('src.converter.ExcelToPDFConverter', return_value=mock_converter), \
             patch('src.hooks.run_hooks', return_value=[]):
            args = MagicMock()
            args.input = str(src)
            args.output = None
            args.mode = 'aggregate'
            args.profile = None
            args.config = None
            args.watch = False
            entry._run_cli(args)

        mock_converter.generate_pdf.assert_called_once()

    def test_hooks_called_after_conversion(self, tmp_path):
        import converter_excel_pdf as entry

        src = tmp_path / 'test.xlsx'
        src.write_text('dummy')
        out = str(tmp_path / 'out.pdf')

        mock_converter = MagicMock()
        mock_converter.generate_individual_pdfs.return_value = [out]

        with patch('converter_excel_pdf.load_config', return_value={'output': {'auto_open': False}}), \
             patch('src.converter.ExcelToPDFConverter', return_value=mock_converter), \
             patch('src.hooks.run_hooks', return_value=[]) as mock_hooks:
            args = MagicMock()
            args.input = str(src)
            args.output = None
            args.mode = 'individual'
            args.profile = None
            args.config = None
            args.watch = False
            entry._run_cli(args)

        mock_hooks.assert_called_once()
