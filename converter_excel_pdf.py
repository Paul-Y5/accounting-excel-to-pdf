#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor Excel → PDF
Aplicação com Interface Gráfica Simples

Autor: Paulo Cunha with the help of AI
Versão: 3.4

Este ficheiro é o entry point da aplicação.
A lógica está modularizada em src/.
"""

import sys
import os
import argparse

# Adicionar src ao path para imports funcionarem
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.config import load_config, load_profile, import_config


def _open_file(path: str):
    """Abre um ficheiro com a aplicação padrão do sistema."""
    import subprocess
    if sys.platform == 'darwin':
        subprocess.run(['open', path])
    elif sys.platform == 'win32':
        os.startfile(path)
    else:
        subprocess.run(['xdg-open', path])


def _run_cli(args):
    """Executa em modo CLI com os argumentos fornecidos."""
    from src.converter import ExcelToPDFConverter
    from src.hooks import run_hooks

    # Carregar config base
    if args.config:
        try:
            config = import_config(args.config)
            print(f"Config importada de: {args.config}")
        except Exception as e:
            print(f"Erro ao importar config: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        config = load_config()

    # Sobrepor com perfil se especificado
    if args.profile:
        try:
            profile_config = load_profile(args.profile)
            config.update(profile_config)
            print(f"Perfil carregado: {args.profile}")
        except Exception as e:
            print(f"Aviso: não foi possível carregar o perfil '{args.profile}': {e}",
                  file=sys.stderr)

    # Modo watch folder
    if args.watch:
        _run_watch(args.input, config)
        return

    # Conversão normal
    excel_path = args.input
    if not os.path.exists(excel_path):
        print(f"Erro: Ficheiro não encontrado: {excel_path}", file=sys.stderr)
        sys.exit(1)

    print(f"A converter: {excel_path}")

    try:
        output_pdf = args.output if args.output else None
        converter = ExcelToPDFConverter(excel_path, output_pdf, config)

        mode = args.mode or 'individual'
        if mode == 'aggregate':
            output_path = converter.generate_pdf()
            outputs = [output_path]
            print(f"PDF gerado: {output_path}")
        else:
            outputs = converter.generate_individual_pdfs()
            print(f"{len(outputs)} PDF(s) gerados em: {os.path.dirname(outputs[0]) if outputs else '—'}")

        # Executar hooks
        hook_results = run_hooks(config, excel_path, outputs)
        for r in hook_results:
            status = 'OK' if r['returncode'] == 0 else f"ERRO (código {r['returncode']})"
            print(f"Hook '{r['hook']}': {status}")
            if r['error']:
                print(f"  {r['error']}", file=sys.stderr)

        # Abrir PDF se configurado
        if config['output'].get('auto_open', True) and outputs:
            _open_file(outputs[0])

    except Exception as e:
        print(f"Erro na conversão: {e}", file=sys.stderr)
        sys.exit(1)


def _run_watch(folder: str, config: dict):
    """Inicia monitorização de pasta no modo CLI (bloqueia até Ctrl+C)."""
    import signal
    from src.watch_folder import WatchFolder

    if not os.path.isdir(folder):
        print(f"Erro: Pasta não encontrada: {folder}", file=sys.stderr)
        sys.exit(1)

    def on_new(path):
        print(f"[watch] Novo ficheiro detectado: {os.path.basename(path)}")

    def on_converted(path, outputs):
        print(f"[watch] Convertido: {os.path.basename(path)} → {len(outputs)} PDF(s)")

    def on_error(path, msg):
        print(f"[watch] Erro em {os.path.basename(path)}: {msg}", file=sys.stderr)

    interval = config.get('automation', {}).get('watch_interval', 5)
    watcher = WatchFolder(folder, config, on_new_file=on_new,
                          on_converted=on_converted, on_error=on_error,
                          interval=interval)
    watcher.start()
    print(f"[watch] A monitorizar: {folder}  (Ctrl+C para parar)")

    def _stop(sig, frame):
        watcher.stop()
        print("\n[watch] Monitorização terminada.")
        sys.exit(0)

    signal.signal(signal.SIGINT, _stop)
    signal.signal(signal.SIGTERM, _stop)

    # Bloquear a thread principal
    import time
    while watcher.is_running:
        time.sleep(1)


def main():
    """Função principal."""
    parser = argparse.ArgumentParser(
        prog='conversor_excel_pdf',
        description='Conversor Excel → PDF',
    )
    parser.add_argument('input', nargs='?',
                        help='Ficheiro Excel (.xlsx) ou pasta (com --watch)')
    parser.add_argument('-o', '--output',
                        help='Caminho de saída do PDF (modo aggregate)')
    parser.add_argument('-m', '--mode', choices=['individual', 'aggregate'],
                        default=None,
                        help='Modo de geração: individual (default) ou aggregate')
    parser.add_argument('-p', '--profile',
                        help='Nome do perfil de configuração a usar')
    parser.add_argument('-c', '--config',
                        help='Caminho para ficheiro de configuração JSON')
    parser.add_argument('-w', '--watch', action='store_true',
                        help='Monitorizar pasta e converter novos ficheiros automaticamente')

    args = parser.parse_args()

    if args.input:
        _run_cli(args)
    else:
        # Modo GUI
        from src.gui.app import ConverterApp
        app = ConverterApp()
        app.run()


if __name__ == "__main__":
    main()
