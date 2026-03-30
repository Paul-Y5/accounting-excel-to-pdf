#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de histórico de conversões.
Regista todas as conversões realizadas para rastreabilidade.
Utiliza base de dados SQLite para persistência.
"""

from src.database import add_history_entry, get_history as _get_history, clear_history as _clear_history


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


def clear_history():
    """Limpa o histórico."""
    _clear_history()
