#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de templates para nomes de ficheiros gerados.

Tokens disponíveis:
    {empresa}  — nome da empresa (do cabeçalho ou do Excel)
    {mes}      — mês de referência do documento
    {nr}       — número do documento (se definido)
    {data}     — data atual no formato YYYYMMDD
    {sigla}    — sigla do cliente (apenas em modo individual)
    {cliente}  — nome do cliente (apenas em modo individual)
"""

import re
from datetime import datetime


class _SafeDict(dict):
    """Dicionário que devolve a chave entre chavetas quando não existe."""

    def __missing__(self, key):
        return '{' + key + '}'


def render_template(template: str, context: dict) -> str:
    """Substitui os tokens do template pelos valores do contexto.

    Tokens em falta ficam intactos (não levantam exceção).
    Caracteres inválidos para nomes de ficheiro são removidos do resultado.

    Args:
        template: String com tokens entre chavetas, ex: '{empresa}_{mes}'.
        context:  Dicionário com os valores para substituição.

    Returns:
        Nome de ficheiro resultante, sem extensão.
    """
    if not template:
        return ''

    result = template.format_map(_SafeDict(context))

    # Remover caracteres inválidos em nomes de ficheiro (Windows e Linux)
    result = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', result)
    # Colapsar espaços e underscores múltiplos
    result = re.sub(r'_{2,}', '_', result)
    result = result.strip('_').strip()

    return result


def get_template_context(data: dict, config: dict) -> dict:
    """Constrói o contexto de substituição a partir dos dados do Excel e config.

    Args:
        data:   Dicionário devolvido por ExcelToPDFConverter.read_excel_data().
        config: Configuração da aplicação.

    Returns:
        Dicionário com os tokens disponíveis para render_template.
    """
    empresa = (
        data.get('empresa', '')
        or config.get('header', {}).get('company_name', '')
    )
    mes = data.get('mes_referencia', '')
    cliente = data.get('cliente', '')

    # Sigla: primeiro item se existir
    sigla = ''
    itens = data.get('itens', [])
    if itens and isinstance(itens[0], dict):
        sigla = itens[0].get('SIGLA', itens[0].get('sigla', ''))

    return {
        'empresa': empresa,
        'mes': mes,
        'nr': '',
        'data': datetime.now().strftime('%Y%m%d'),
        'sigla': sigla,
        'cliente': cliente,
    }
