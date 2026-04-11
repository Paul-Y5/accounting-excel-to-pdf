#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de post-conversion hooks.
Executa comandos shell configurados pelo utilizador após cada conversão.
"""

import os
import subprocess
import sys


def run_hooks(config: dict, source_path: str, output_paths: list) -> list:
    """Executa os hooks configurados após uma conversão.

    Cada hook é um comando shell com suporte a variáveis de substituição:
    - {source}   — caminho do ficheiro Excel de origem
    - {output}   — caminho do primeiro PDF gerado
    - {outputs}  — todos os PDFs separados por vírgula
    - {folder}   — pasta onde os PDFs foram gerados

    Args:
        config: Configuração da aplicação.
        source_path: Caminho do ficheiro Excel de origem.
        output_paths: Lista de ficheiros PDF gerados.

    Returns:
        Lista de resultados [{hook, command, returncode, stdout, stderr, error}].
    """
    hooks = config.get('automation', {}).get('hooks', [])
    if not hooks:
        return []

    results = []
    first_output = output_paths[0] if output_paths else ''
    folder = os.path.dirname(first_output) if first_output else ''
    outputs_str = ','.join(output_paths)

    for hook in hooks:
        if not hook.get('enabled', True):
            continue
        cmd_template = hook.get('command', '').strip()
        if not cmd_template:
            continue

        cmd = cmd_template.replace('{source}', source_path)
        cmd = cmd.replace('{output}', first_output)
        cmd = cmd.replace('{outputs}', outputs_str)
        cmd = cmd.replace('{folder}', folder)

        result = {'hook': hook.get('name', ''), 'command': cmd,
                  'returncode': None, 'stdout': '', 'stderr': '', 'error': ''}
        try:
            proc = subprocess.run(
                cmd,
                shell=True,
                capture_output=True,
                text=True,
                timeout=hook.get('timeout', 30),
            )
            result['returncode'] = proc.returncode
            result['stdout'] = proc.stdout.strip()
            result['stderr'] = proc.stderr.strip()
        except subprocess.TimeoutExpired:
            result['error'] = 'Timeout expirado'
        except Exception as e:
            result['error'] = str(e)
        results.append(result)

    return results
