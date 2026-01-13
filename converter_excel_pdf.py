#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor Excel → PDF
Aplicação com Interface Gráfica Simples

Autor: Paulo Cunha with the help of AI
Versão: 2.0

Este ficheiro é o entry point da aplicação.
A lógica está modularizada em src/.
"""

import sys
import os

# Adicionar src ao path para imports funcionarem
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui.app import ConverterApp
from src.converter import ExcelToPDFConverter
from src.config import load_config


def main():
    """Função principal."""
    if len(sys.argv) > 1:
        # Modo CLI - converter ficheiro diretamente
        excel_path = sys.argv[1]
        
        if not os.path.exists(excel_path):
            print(f"Erro: Ficheiro não encontrado: {excel_path}")
            sys.exit(1)
        
        print(f"A converter: {excel_path}")
        
        try:
            config = load_config()
            converter = ExcelToPDFConverter(excel_path, config=config)
            output_path = converter.generate_pdf()
            print(f"PDF gerado com sucesso: {output_path}")
            
            # Abrir PDF se configurado
            if config['output'].get('auto_open', True):
                import subprocess
                if sys.platform == 'darwin':
                    subprocess.run(['open', output_path])
                elif sys.platform == 'win32':
                    os.startfile(output_path)
                else:
                    subprocess.run(['xdg-open', output_path])
                    
        except Exception as e:
            print(f"Erro na conversão: {e}")
            sys.exit(1)
    else:
        # Modo GUI
        app = ConverterApp()
        app.run()


if __name__ == "__main__":
    main()