#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de configuração do Conversor Excel → PDF.
Gestão de configurações padrão e persistência.
"""

import os
import sys
import json
import copy


# ============================================
# CONFIGURAÇÕES PADRÃO
# ============================================
DEFAULT_CONFIG = {
    'pdf': {
        'page_size': 'A4',
        'orientation': 'portrait',
        'margin_top': 15,
        'margin_bottom': 15,
        'margin_left': 15,
        'margin_right': 15,
    },
    'header': {
        'show_header': True,
        'company_name': 'EMPRESA EXEMPLO, LDA',
        'company_address': 'Rua Exemplo, 123 - 4000-000 Porto',
        'company_phone': '+351 220 000 000',
        'company_email': 'geral@empresa.pt',
        'company_nif': 'PT 500 000 000',
        'logo_path': '',
    },
    'colors': {
        'header_bg': '#2d3748',
        'header_text': '#ffffff',
        'row_alt': '#f7fafc',
        'border': '#e2e8f0',
        'title': '#1a365d',
        'total_bg': '#edf2f7',
        'total_text': '#1a365d',
        'highlight_positive': '#48bb78',
        'highlight_negative': '#fc8181',
    },
    'table': {
        'font_size': 9,
        'header_font_size': 10,
        'row_padding': 6,
        'show_grid': True,
        'alternate_rows': True,
    },
    'footer': {
        'show_signatures': True,
        'show_date': True,
        'show_observations': True,
        'custom_footer': '',
    },
    'output': {
        'auto_open': True,
        'add_timestamp': False,
        'output_folder': '',
    },
    'contabilidade': {
        'enabled': True,
        'colunas': 'Nr., SIGLA, Cliente, CONTAB, Iva, Subtotal, Extras, Duodécimos, S.Social GER, S.Soc Emp, Ret. IRS, Ret. IRS EXT, SbTx/Fcomp, Outro, TOTAL',
        'destacar_total': True,
        'destacar_valores': True,
    },
    'security': {
        'pdf_password': '',
        'pdf_owner_password': '',
    },
    'watermark': {
        'enabled': False,
        'text': 'RASCUNHO',
        'opacity': 0.1,
    },
    'banking': {
        'show_banking': True,
        'title': 'Nossos Dados Bancários:',
        'accounts': [
            {
                'bank_name': 'ABANCA',
                'iban': 'PT50 0170 3782 0304 0053 5672 9',
                'default': True,
            }
        ],
    },
    'recent': {
        'last_excel_dir': '',
        'last_output_dir': '',
    }
}


def get_config_dir() -> str:
    """Retorna o diretório de configuração do utilizador.
    
    Windows: %APPDATA%/ConversorExcelPDF/
    Linux:   ~/.config/ConversorExcelPDF/
    macOS:   ~/Library/Application Support/ConversorExcelPDF/
    """
    app_name = "ConversorExcelPDF"
    
    if sys.platform == "win32":
        # Windows: %APPDATA%
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
        config_dir = os.path.join(base, app_name)
    elif sys.platform == "darwin":
        # macOS: ~/Library/Application Support/
        config_dir = os.path.join(os.path.expanduser("~"), "Library", "Application Support", app_name)
    else:
        # Linux/Unix: ~/.config/
        xdg_config = os.environ.get("XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config"))
        config_dir = os.path.join(xdg_config, app_name)
    
    # Criar diretório se não existir
    os.makedirs(config_dir, exist_ok=True)
    return config_dir


def get_config_path() -> str:
    """Retorna o caminho do ficheiro de configuração."""
    return os.path.join(get_config_dir(), 'config.json')


def load_config() -> dict:
    """Carrega configurações do ficheiro JSON."""
    config_path = get_config_path()
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                saved_config = json.load(f)

                # Merge com defaults para garantir que todas as chaves existem
                config = copy.deepcopy(DEFAULT_CONFIG)
                for section, values in saved_config.items():
                    if section in config:
                        if isinstance(values, dict):
                            config[section].update(values)
                        else:
                            config[section] = values

                # Migrar banking antigo (bank_name/iban flat) para accounts list
                banking = config.get('banking', {})
                if 'accounts' not in banking and 'bank_name' in banking:
                    config['banking'] = {
                        'show_banking': banking.get('show_banking', True),
                        'title': banking.get('title', 'Nossos Dados Bancários:'),
                        'accounts': [
                            {
                                'bank_name': banking.get('bank_name', ''),
                                'iban': banking.get('iban', ''),
                                'default': True,
                            }
                        ],
                    }

                return config
        except Exception:
            pass
    return copy.deepcopy(DEFAULT_CONFIG)


def get_profiles_dir() -> str:
    """Retorna o diretório de perfis de configuração."""
    profiles_dir = os.path.join(get_config_dir(), 'profiles')
    os.makedirs(profiles_dir, exist_ok=True)
    return profiles_dir


def list_profiles() -> list:
    """Lista os nomes dos perfis guardados."""
    profiles_dir = get_profiles_dir()
    profiles = []
    for f in sorted(os.listdir(profiles_dir)):
        if f.endswith('.json'):
            profiles.append(f[:-5])  # Remove .json
    return profiles


def save_profile(name: str, config: dict) -> bool:
    """Guarda uma configuração como perfil.

    Args:
        name: Nome do perfil.
        config: Configuração a guardar.

    Returns:
        True se guardou com sucesso.
    """
    path = os.path.join(get_profiles_dir(), f'{name}.json')
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def load_profile(name: str) -> dict:
    """Carrega um perfil de configuração.

    Args:
        name: Nome do perfil.

    Returns:
        Configuração do perfil ou None se não existir.
    """
    path = os.path.join(get_profiles_dir(), f'{name}.json')
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            saved = json.load(f)
        # Merge com defaults
        config = copy.deepcopy(DEFAULT_CONFIG)
        for section, values in saved.items():
            if section in config:
                if isinstance(values, dict):
                    config[section].update(values)
                else:
                    config[section] = values
        return config
    except Exception:
        return None


def delete_profile(name: str) -> bool:
    """Apaga um perfil de configuração.

    Args:
        name: Nome do perfil.

    Returns:
        True se apagou com sucesso.
    """
    path = os.path.join(get_profiles_dir(), f'{name}.json')
    try:
        os.remove(path)
        return True
    except Exception:
        return False


def save_config(config: dict) -> bool:
    """Guarda configurações no ficheiro JSON.
    
    Returns:
        bool: True se guardou com sucesso, False caso contrário.
    """
    config_path = get_config_path()
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Erro ao guardar configurações: {e}")
        return False
