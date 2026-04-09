#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de gestão de fontes personalizadas (.ttf) para ReportLab.
Permite registar e utilizar fontes externas na geração de PDFs.
"""

import os

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def register_font(name: str, path: str) -> bool:
    """Regista uma fonte TrueType (.ttf) no ReportLab.

    Args:
        name: Nome lógico da fonte (ex: ``'MinhaFonte'``).
        path: Caminho absoluto para o ficheiro ``.ttf``.

    Returns:
        True se registou com sucesso, False caso contrário.
    """
    if not path or not os.path.isfile(path):
        return False

    try:
        pdfmetrics.registerFont(TTFont(name, path))
        return True
    except Exception:
        return False


def load_fonts_from_config(config: dict) -> list:
    """Carrega todas as fontes registadas na configuração.

    Args:
        config: Dicionário de configuração da aplicação.

    Returns:
        Lista de nomes de fontes registadas com sucesso.
    """
    fonts_cfg = config.get('fonts', {})
    registered = fonts_cfg.get('registered', [])
    loaded = []

    for entry in registered:
        name = entry.get('name', '')
        path = entry.get('path', '')
        if name and path and register_font(name, path):
            loaded.append(name)

    return loaded


def get_body_font(config: dict) -> str:
    """Devolve o nome da fonte de corpo configurada, com fallback para Helvetica."""
    return config.get('fonts', {}).get('body_font', 'Helvetica')


def get_header_font(config: dict) -> str:
    """Devolve o nome da fonte de cabeçalho configurada, com fallback para Helvetica-Bold."""
    return config.get('fonts', {}).get('header_font', 'Helvetica-Bold')
