#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de histórico de conversões.
Regista todas as conversões realizadas para rastreabilidade.
"""

import json
import os
from datetime import datetime

from src.config import get_config_dir


def _get_history_path() -> str:
    """Retorna o caminho do ficheiro de histórico."""
    return os.path.join(get_config_dir(), 'history.json')


def _load_history() -> list:
    """Carrega o histórico existente."""
    path = _get_history_path()
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []
    return []


def _save_history(history: list):
    """Guarda o histórico."""
    path = _get_history_path()
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(history, f, indent=2, ensure_ascii=False)


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
    history = _load_history()

    entry = {
        'timestamp': datetime.now().isoformat(),
        'source_file': os.path.basename(source_file),
        'source_path': source_file,
        'output_path': output_path,
        'mode': mode,
        'clients_count': clients_count,
        'success': success,
        'error': error_msg,
    }

    history.append(entry)

    # Manter apenas as últimas 500 entradas
    if len(history) > 500:
        history = history[-500:]

    _save_history(history)


def get_history(limit: int = 50) -> list:
    """Retorna as últimas entradas do histórico.

    Args:
        limit: Número máximo de entradas a retornar.

    Returns:
        Lista de entradas do histórico (mais recentes primeiro).
    """
    history = _load_history()
    return list(reversed(history[-limit:]))


def clear_history():
    """Limpa o histórico."""
    _save_history([])
