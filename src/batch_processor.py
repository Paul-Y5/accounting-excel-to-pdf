#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de processamento em lote de ficheiros Excel.
Processa múltiplos ficheiros de uma pasta e gera PDFs para cada um.
"""

import os

from src.converter import ExcelToPDFConverter


def find_excel_files(folder_path: str) -> list:
    """Retorna lista de ficheiros Excel (.xlsx/.xls) numa pasta.

    Ignora ficheiros temporários do Excel (prefixo ~$).
    Retorna lista ordenada pelo nome do ficheiro.

    Args:
        folder_path: Caminho da pasta a pesquisar.

    Returns:
        Lista de caminhos absolutos dos ficheiros Excel.

    Raises:
        ValueError: Se a pasta não existir.
    """
    if not os.path.isdir(folder_path):
        raise ValueError(f"Pasta não encontrada: {folder_path}")

    files = []
    for name in sorted(os.listdir(folder_path)):
        if name.startswith('~$'):
            continue
        if name.lower().endswith(('.xlsx', '.xls')):
            files.append(os.path.join(folder_path, name))
    return files


def process_batch(folder_path: str, config: dict, mode: str = 'individual',
                  progress_callback=None) -> list:
    """Processa todos os ficheiros Excel de uma pasta.

    Args:
        folder_path: Pasta com os ficheiros Excel.
        config: Configurações da aplicação.
        mode: 'individual' (1 PDF por cliente) ou 'aggregate' (1 PDF por ficheiro).
        progress_callback: Função chamada a cada ficheiro com (current, total, filename).
                           current=0..total-1 antes do ficheiro, current=total depois do último.

    Returns:
        Lista de resultados, um por ficheiro:
        [{file, filename, success, output_path, clients_count, error}]
    """
    files = find_excel_files(folder_path)
    if not files:
        return []

    total = len(files)
    results = []

    for i, excel_path in enumerate(files):
        filename = os.path.basename(excel_path)

        if progress_callback:
            progress_callback(i, total, filename)

        try:
            converter = ExcelToPDFConverter(excel_path, None, config)
            data = converter.read_excel_data()
            clients_count = len(data.get('itens', []))

            if mode == 'individual':
                output_files = converter.generate_individual_pdfs()
                output_path = os.path.dirname(output_files[0]) if output_files else folder_path
            else:
                output_path = converter.generate_pdf()

            results.append({
                'file': excel_path,
                'filename': filename,
                'success': True,
                'output_path': output_path,
                'clients_count': clients_count,
                'error': '',
            })

        except Exception as e:
            results.append({
                'file': excel_path,
                'filename': filename,
                'success': False,
                'output_path': '',
                'clients_count': 0,
                'error': str(e),
            })

        if progress_callback:
            progress_callback(i + 1, total, filename)

    return results
