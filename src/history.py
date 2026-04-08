#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de histórico de conversões.
Regista todas as conversões realizadas para rastreabilidade.
Utiliza base de dados SQLite para persistência.
"""

from src.database import (add_history_entry, get_history as _get_history,
                          get_history_filtered as _get_history_filtered,
                          clear_history as _clear_history,
                          export_history_csv as _export_csv,
                          export_history_excel as _export_excel)


def add_entry(source_file: str, output_path: str, mode: str,
              clients_count: int, success: bool, error_msg: str = ''):
    """Adiciona uma entrada ao histórico.

    Args:
        source_file: Caminho do ficheiro Excel de origem.
        output_path: Caminho do ficheiro/pasta de saída.
        mode: Modo de geração ('aggregate' ou 'individual').
        clients_count: Número de clientes/registos processados.
        success: Se a conversão foi bem sucedida.
        error_msg: Mensagem de erro (se aplicável).
    """
    add_history_entry(source_file, output_path, mode, clients_count, success, error_msg)


def get_history(limit: int = 50) -> list:
    """Retorna as últimas entradas do histórico.

    Args:
        limit: Número máximo de entradas a retornar.

    Returns:
        Lista de entradas do histórico (mais recentes primeiro).
    """
    return _get_history(limit)


def get_history_filtered(
    limit: int = 100,
    date_from: str = None,
    date_to: str = None,
    success_only: bool = None,
    search_term: str = None,
) -> list:
    """Retorna entradas do histórico com filtros opcionais.

    Args:
        limit:        Número máximo de resultados.
        date_from:    Data de início ISO (YYYY-MM-DD), inclusivo.
        date_to:      Data de fim ISO (YYYY-MM-DD), inclusivo.
        success_only: True = sucesso, False = erros, None = todos.
        search_term:  Pesquisa parcial no nome do ficheiro.

    Returns:
        Lista de entradas (mais recentes primeiro).
    """
    return _get_history_filtered(limit, date_from, date_to, success_only, search_term)


def clear_history():
    """Limpa o histórico."""
    _clear_history()


def export_to_csv(output_path: str, limit: int = None) -> str:
    """Exporta o histórico para CSV.

    Args:
        output_path: Caminho do ficheiro .csv a criar.
        limit: Número máximo de entradas (None = todas).

    Returns:
        Caminho do ficheiro criado.
    """
    return _export_csv(output_path, limit)


def export_to_excel(output_path: str, limit: int = None) -> str:
    """Exporta o histórico para Excel (.xlsx).

    Args:
        output_path: Caminho do ficheiro .xlsx a criar.
        limit: Número máximo de entradas (None = todas).

    Returns:
        Caminho do ficheiro criado.
    """
    return _export_excel(output_path, limit)
