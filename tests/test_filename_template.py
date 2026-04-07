"""
Testes unitários para o módulo de templates de nomes de ficheiro.
"""

import pytest
from src.filename_template import render_template, get_template_context


class TestRenderTemplate:
    """Testes para render_template."""

    def test_empty_template_returns_empty(self):
        assert render_template('', {}) == ''

    def test_single_token_substitution(self):
        result = render_template('{empresa}', {'empresa': 'Empresa Teste'})
        assert result == 'Empresa Teste'

    def test_multiple_tokens(self):
        result = render_template('{empresa}_{mes}', {'empresa': 'ABC', 'mes': 'Janeiro_2026'})
        assert result == 'ABC_Janeiro_2026'

    def test_all_tokens(self):
        context = {
            'empresa': 'ABC Lda',
            'mes': 'Jan2026',
            'nr': '0042',
            'data': '20260101',
            'sigla': 'CLI',
            'cliente': 'Cliente X',
        }
        result = render_template('{empresa}_{mes}_{nr}_{data}_{sigla}_{cliente}', context)
        assert result == 'ABC Lda_Jan2026_0042_20260101_CLI_Cliente X'

    def test_missing_token_remains_intact(self):
        result = render_template('{empresa}_{inexistente}', {'empresa': 'ABC'})
        assert result == 'ABC_{inexistente}'

    def test_invalid_filename_chars_removed(self):
        result = render_template('{empresa}', {'empresa': 'ABC/Lda:2026'})
        assert '/' not in result
        assert ':' not in result

    def test_multiple_underscores_collapsed(self):
        result = render_template('{empresa}__{mes}', {'empresa': 'ABC', 'mes': 'Jan'})
        assert '__' not in result

    def test_leading_trailing_underscores_stripped(self):
        result = render_template('_{empresa}_', {'empresa': 'ABC'})
        assert not result.startswith('_')
        assert not result.endswith('_')

    def test_no_tokens_returns_literal(self):
        result = render_template('relatorio_mensal', {})
        assert result == 'relatorio_mensal'

    def test_empty_context_missing_tokens_remain(self):
        result = render_template('{empresa}_{mes}', {})
        assert '{empresa}' in result
        assert '{mes}' in result


class TestGetTemplateContext:
    """Testes para get_template_context."""

    def test_returns_dict(self):
        context = get_template_context({}, {})
        assert isinstance(context, dict)

    def test_empresa_from_data(self):
        data = {'empresa': 'Empresa do Excel'}
        context = get_template_context(data, {})
        assert context['empresa'] == 'Empresa do Excel'

    def test_empresa_fallback_to_config(self):
        data = {'empresa': ''}
        config = {'header': {'company_name': 'Empresa Config'}}
        context = get_template_context(data, config)
        assert context['empresa'] == 'Empresa Config'

    def test_mes_from_data(self):
        data = {'mes_referencia': 'Janeiro 2026'}
        context = get_template_context(data, {})
        assert context['mes'] == 'Janeiro 2026'

    def test_sigla_from_first_item(self):
        data = {'itens': [{'SIGLA': 'CLI', 'Cliente': 'Cliente X'}]}
        context = get_template_context(data, {})
        assert context['sigla'] == 'CLI'

    def test_sigla_empty_when_no_items(self):
        data = {'itens': []}
        context = get_template_context(data, {})
        assert context['sigla'] == ''

    def test_data_token_is_8_digits(self):
        context = get_template_context({}, {})
        assert len(context['data']) == 8
        assert context['data'].isdigit()

    def test_nr_defaults_to_empty(self):
        context = get_template_context({}, {})
        assert context['nr'] == ''
