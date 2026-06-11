#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de interface gráfica do Conversor Excel → PDF.

A classe ConverterApp compõe os mixins de src/gui/tabs/ — cada tab vive
no seu próprio módulo; aqui fica apenas a orquestração (janela, notebook,
atalhos, tema e leitura/escrita da configuração global da UI).
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox

from src.config import load_config, save_config
from src.database import init_db, migrate_from_json
from src.gui.tabs import (
    DashboardTabMixin,
    ConvertTabMixin,
    ProfilesTabMixin,
    BatchTabMixin,
    HistoryTabMixin,
    SettingsTabMixin,
    ContabilidadeTabMixin,
    BankingTabMixin,
    DocSequenceTabMixin,
    QrcodeFontsTabMixin,
    AutomationTabMixin,
)


class ConverterApp(
    DashboardTabMixin,
    ConvertTabMixin,
    ProfilesTabMixin,
    BatchTabMixin,
    HistoryTabMixin,
    SettingsTabMixin,
    ContabilidadeTabMixin,
    BankingTabMixin,
    DocSequenceTabMixin,
    QrcodeFontsTabMixin,
    AutomationTabMixin,
):
    """Aplicação principal com interface gráfica para conversão de Excel para PDF."""

    # Constantes de UI
    _PAD_OUTER = 12          # padding exterior das tabs
    _PAD_SECTION = (0, 8)    # espaço vertical entre secções
    _PAD_INNER = 10          # padding interior dos LabelFrames
    _FONT_FAMILY = 'Helvetica'
    _FONT_SIZE = 10
    _FONT_HEADER = 14

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor Excel")
        self.root.geometry("780x640")
        self.root.minsize(700, 560)
        self.root.resizable(True, True)

        # Tema moderno (Sun Valley) — aplicado antes de criar widgets
        self._sv_ttk_available = False
        try:
            import sv_ttk
            self._sv_ttk_available = True
        except ImportError:
            pass

        # Inicializar base de dados SQLite
        init_db()
        migrate_from_json()

        # Carregar configurações
        self.config = load_config()

        # Aplicar tema guardado
        if self._sv_ttk_available:
            import sv_ttk
            sv_ttk.set_theme(self.config.get('ui', {}).get('theme', 'light'))

        # Últimos PDFs gerados (para envio por email)
        self._last_generated_files = []

        # Variáveis
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self._setup_ui()
        self._load_config_to_ui()
        self._setup_keyboard_shortcuts()
        self._setup_drag_drop()

    def _setup_keyboard_shortcuts(self):
        """Configura atalhos de teclado globais."""
        self.root.bind('<Control-o>', lambda e: self._browse_excel())
        self.root.bind('<Control-g>', lambda e: self._generate())
        self.root.bind('<Control-s>', lambda e: self._save_config())
        self.root.bind('<Control-e>', lambda e: self._export_excel())
        self.root.bind('<Control-p>', lambda e: self._preview_excel())

    def _setup_drag_drop(self):
        """Configura drag & drop de ficheiros Excel."""
        try:
            self.root.tk.call('package', 'require', 'tkdnd')
            self._has_tkdnd = True
        except tk.TclError:
            self._has_tkdnd = False

        if self._has_tkdnd:
            self.root.tk.call('tkdnd::drop_target', 'register', str(self.root), ('DND_Files',))
            self.root.tk.call('bind', str(self.root), '<<Drop:DND_Files>>', self.root.register(self._on_drop))
        else:
            # Fallback: aceitar ficheiros via evento de ficheiro (funcional em todos os OS)
            pass

    def _on_drop(self, event_data):
        """Processa ficheiro largado via drag & drop."""
        # tkdnd pode envolver o path em {} se tiver espaços
        path = event_data.strip().strip('{}')
        if path.lower().endswith(('.xlsx', '.xls', '.xlsm')):
            self.excel_path.set(path)
            self.config.setdefault('recent', {})['last_excel_dir'] = os.path.dirname(path)
            save_config(self.config)
            self.status_var.set(f"Ficheiro carregado: {os.path.basename(path)}")
        else:
            messagebox.showwarning("Aviso", "Apenas ficheiros Excel (.xlsx, .xls, .xlsm) são suportados.")
        return event_data

    def _setup_ui(self):
        """Configura a interface."""
        # Fonte global
        default_font = (self._FONT_FAMILY, self._FONT_SIZE)
        self.root.option_add('*Font', default_font)

        # Estilos ttk
        style = ttk.Style()
        style.configure('TButton', padding=(12, 6))
        style.configure('TLabel', padding=2)
        style.configure('Header.TLabel', font=(self._FONT_FAMILY, self._FONT_HEADER, 'bold'))
        style.configure('Status.TLabel', font=(self._FONT_FAMILY, 9))

        # Estilo de destaque para o botão principal
        style.configure('Accent.TButton', padding=(16, 8))
        style.map('Accent.TButton',
                  background=[('active', '#005a9e'), ('!active', '#0078D4')],
                  foreground=[('active', 'white'), ('!active', 'white')])

        # Barra inferior (tema) — criada antes do notebook para ficar na base
        self._setup_bottom_bar()

        # Notebook (tabs) — 6 tabs principais
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=self._PAD_OUTER,
                           pady=(self._PAD_OUTER, 0))

        # Tab 0: Dashboard
        self.tab_dashboard = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_dashboard, text='Dashboard')
        self._setup_dashboard_tab()

        # Tab 1: Conversão
        self.tab_convert = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_convert, text='Converter')
        self._setup_convert_tab()

        # Tab 2: Perfis
        self.tab_profiles = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_profiles, text='Perfis')
        self._setup_profiles_tab()

        # Tab 3: Multificheiros
        self.tab_batch = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_batch, text='Multificheiros')
        self._setup_batch_tab()

        # Tab 4: Histórico
        self.tab_history = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_history, text='Histórico')
        self._setup_history_tab()

        # Tab 5: Definições (no final)
        self.tab_settings = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_settings, text='Definições')
        self._setup_settings_tab()

    def _setup_bottom_bar(self):
        """Barra inferior com controlos globais da interface."""
        bar = ttk.Frame(self.root, padding=(self._PAD_OUTER, 4))
        bar.pack(side='bottom', fill='x')

        ttk.Separator(self.root, orient='horizontal').pack(side='bottom', fill='x')

        current_theme = self.config.get('ui', {}).get('theme', 'light')
        self._theme_btn_text = tk.StringVar(
            value='Tema: Escuro' if current_theme == 'light' else 'Tema: Claro'
        )
        ttk.Button(bar, textvariable=self._theme_btn_text,
                   command=self._toggle_theme).pack(side='right')

    def _toggle_theme(self):
        """Alterna entre tema claro e escuro."""
        if not self._sv_ttk_available:
            messagebox.showinfo("Tema", "O pacote sv-ttk não está instalado.")
            return

        import sv_ttk
        current = self.config.get('ui', {}).get('theme', 'light')
        new_theme = 'dark' if current == 'light' else 'light'

        sv_ttk.set_theme(new_theme)
        self.config.setdefault('ui', {})['theme'] = new_theme
        self._theme_btn_text.set('Tema: Escuro' if new_theme == 'light' else 'Tema: Claro')
        save_config(self.config)

    def _browse_excel(self):
        """Seleciona ficheiro Excel, lembrando a última pasta usada."""
        from tkinter import filedialog
        initial_dir = self.config.get('recent', {}).get('last_excel_dir', '')
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser('~')

        path = filedialog.askopenfilename(
            title="Selecionar ficheiro Excel",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.excel_path.set(path)
            # Guardar última pasta
            self.config.setdefault('recent', {})['last_excel_dir'] = os.path.dirname(path)
            save_config(self.config)

    def _browse_output(self):
        """Seleciona ficheiro de saída, lembrando a última pasta."""
        from tkinter import filedialog
        initial_dir = self.config.get('recent', {}).get('last_output_dir', '')
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser('~')

        path = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            initialdir=initial_dir,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.output_path.set(path)
            self.config.setdefault('recent', {})['last_output_dir'] = os.path.dirname(path)
            save_config(self.config)

    def _load_config_to_ui(self):
        """Carrega configurações para a UI."""
        # Já feito nos setup_*_tab através dos valores padrão
        pass

    def _get_config_from_ui(self) -> dict:
        """Obtém configurações da UI."""
        # Obter texto das colunas de contabilidade
        contab_colunas = self.contab_colunas_text.get('1.0', tk.END).strip() if hasattr(self, 'contab_colunas_text') else ''
        # Obter larguras de colunas configuradas (ignorar zeros)
        contab_col_widths = {}
        if hasattr(self, 'contab_col_widths_vars'):
            for col, var in self.contab_col_widths_vars.items():
                try:
                    v = float(var.get())
                    if v > 0:
                        contab_col_widths[col] = v
                except (ValueError, TypeError):
                    pass

        return {
            'pdf': {
                'page_size': self.page_size_var.get(),
                'orientation': self.orientation_var.get(),
                'margin_top': self.margin_top_var.get(),
                'margin_bottom': self.margin_bottom_var.get(),
                'margin_left': self.margin_left_var.get(),
                'margin_right': self.margin_right_var.get(),
                'show_iva_summary': self.show_iva_summary_var.get()
                    if hasattr(self, 'show_iva_summary_var') else True,
            },
            'header': {
                'show_header': self.show_header_var.get(),
                'company_name': self.company_name_var.get(),
                'company_address': self.company_address_var.get(),
                'company_phone': self.company_phone_var.get(),
                'company_email': self.company_email_var.get(),
                'company_website': self.company_website_var.get(),
                'company_nif': self.company_nif_var.get(),
                'logo_path': self.logo_path_var.get(),
            },
            'colors': {key: var.get() for key, var in self.color_vars.items() if not key.endswith('_btn')},
            'table': {
                'font_size': self.font_size_var.get(),
                'header_font_size': self.header_font_size_var.get(),
                'row_padding': self.row_padding_var.get(),
                'show_grid': self.show_grid_var.get(),
                'alternate_rows': self.alternate_rows_var.get(),
            },
            'footer': {
                'show_signatures': self.show_signatures_var.get(),
                'show_date': self.show_date_var.get(),
                'show_observations': self.show_observations_var.get(),
                'custom_footer': self.custom_footer_var.get(),
            },
            'output': {
                'auto_open': self.auto_open_var.get(),
                'add_timestamp': self.add_timestamp_var.get(),
                'output_folder': '',
                'filename_template': self.filename_template_var.get()
                    if hasattr(self, 'filename_template_var') else '',
            },
            'contabilidade': {
                'enabled': True,
                'colunas': contab_colunas,
                'destacar_total': self.contab_destacar_total_var.get() if hasattr(self, 'contab_destacar_total_var') else True,
                'destacar_valores': self.contab_destacar_valores_var.get() if hasattr(self, 'contab_destacar_valores_var') else True,
                'col_widths': contab_col_widths,
            },
            'security': {
                'pdf_password': self.pdf_password_var.get() if hasattr(self, 'pdf_password_var') else '',
                'pdf_owner_password': '',
            },
            'watermark': {
                'enabled': self.watermark_enabled_var.get() if hasattr(self, 'watermark_enabled_var') else False,
                'text': self.watermark_text_var.get() if hasattr(self, 'watermark_text_var') else 'RASCUNHO',
                'opacity': 0.1,
            },
            'qrcode': {
                'enabled': self.qr_enabled_var.get() if hasattr(self, 'qr_enabled_var') else False,
                'content': self.qr_content_var.get() if hasattr(self, 'qr_content_var') else 'nif',
                'size_mm': self.qr_size_var.get() if hasattr(self, 'qr_size_var') else 25,
            },
            'fonts': self._get_fonts_from_ui(),
            'banking': self._get_banking_from_ui(),
            'automation': self._get_automation_from_ui(),
            'recent': self.config.get('recent', {'last_excel_dir': '', 'last_output_dir': ''}),
            'ui': {
                'theme': self.config.get('ui', {}).get('theme', 'light'),
                'notifications_enabled': self.notifications_enabled_var.get()
                    if hasattr(self, 'notifications_enabled_var') else True,
            },
        }

    def _save_config(self):
        """Guarda configurações."""
        self.config = self._get_config_from_ui()
        save_config(self.config)
        messagebox.showinfo("Sucesso", "Configurações guardadas com sucesso!")

    def _reload_config_to_ui(self):
        """Recarrega os valores da config atual para todos os widgets da UI."""
        cfg = self.config
        # PDF
        self.page_size_var.set(cfg['pdf']['page_size'])
        self.orientation_var.set(cfg['pdf']['orientation'])
        self.margin_top_var.set(cfg['pdf']['margin_top'])
        self.margin_bottom_var.set(cfg['pdf']['margin_bottom'])
        self.margin_left_var.set(cfg['pdf']['margin_left'])
        self.margin_right_var.set(cfg['pdf']['margin_right'])
        # Header
        self.show_header_var.set(cfg['header']['show_header'])
        self.company_name_var.set(cfg['header']['company_name'])
        self.company_address_var.set(cfg['header']['company_address'])
        self.company_phone_var.set(cfg['header']['company_phone'])
        self.company_email_var.set(cfg['header']['company_email'])
        self.company_website_var.set(cfg['header'].get('company_website', ''))
        self.company_nif_var.set(cfg['header']['company_nif'])
        self.logo_path_var.set(cfg['header'].get('logo_path', ''))
        # Table
        self.font_size_var.set(cfg['table']['font_size'])
        self.header_font_size_var.set(cfg['table']['header_font_size'])
        self.row_padding_var.set(cfg['table']['row_padding'])
        self.show_grid_var.set(cfg['table']['show_grid'])
        self.alternate_rows_var.set(cfg['table']['alternate_rows'])
        # PDF extras
        if hasattr(self, 'show_iva_summary_var'):
            self.show_iva_summary_var.set(cfg.get('pdf', {}).get('show_iva_summary', True))
        # Footer
        self.show_signatures_var.set(cfg['footer']['show_signatures'])
        self.show_date_var.set(cfg['footer']['show_date'])
        self.show_observations_var.set(cfg['footer']['show_observations'])
        self.custom_footer_var.set(cfg['footer'].get('custom_footer', ''))
        # Output
        self.auto_open_var.set(cfg['output']['auto_open'])
        self.add_timestamp_var.set(cfg['output']['add_timestamp'])
        if hasattr(self, 'filename_template_var'):
            self.filename_template_var.set(cfg.get('output', {}).get('filename_template', ''))
        # Colors
        for key, var in self.color_vars.items():
            if not key.endswith('_btn') and key in cfg.get('colors', {}):
                var.set(cfg['colors'][key])
        # Contabilidade
        contab_cfg = cfg.get('contabilidade', {})
        if hasattr(self, 'contab_colunas_text'):
            self.contab_colunas_text.delete('1.0', tk.END)
            self.contab_colunas_text.insert('1.0', contab_cfg.get('colunas', ''))
        if hasattr(self, 'contab_destacar_total_var'):
            self.contab_destacar_total_var.set(contab_cfg.get('destacar_total', True))
        if hasattr(self, 'contab_destacar_valores_var'):
            self.contab_destacar_valores_var.set(contab_cfg.get('destacar_valores', True))
        if hasattr(self, 'contab_col_widths_vars'):
            col_widths_cfg = contab_cfg.get('col_widths', {})
            for col, var in self.contab_col_widths_vars.items():
                var.set(str(col_widths_cfg.get(col, 0)))
        # Security
        self.pdf_password_var.set(cfg.get('security', {}).get('pdf_password', ''))
        self.watermark_enabled_var.set(cfg.get('watermark', {}).get('enabled', False))
        self.watermark_text_var.set(cfg.get('watermark', {}).get('text', 'RASCUNHO'))
        # Banking
        self.show_banking_var.set(cfg.get('banking', {}).get('show_banking', True))
        self.banking_title_var.set(cfg.get('banking', {}).get('title', 'Nossos Dados Bancários:'))
        # Reload accounts treeview
        for item in self.accounts_tree.get_children():
            self.accounts_tree.delete(item)
        for acc in cfg.get('banking', {}).get('accounts', []):
            default_mark = 'Sim' if acc.get('default', False) else ''
            self.accounts_tree.insert('', 'end', values=(
                acc.get('bank_name', ''), acc.get('iban', ''), default_mark))
        # UI
        theme = cfg.get('ui', {}).get('theme', 'light')
        if self._sv_ttk_available:
            import sv_ttk
            sv_ttk.set_theme(theme)
        self._theme_btn_text.set('Tema: Escuro' if theme == 'light' else 'Tema: Claro')
        if hasattr(self, 'notifications_enabled_var'):
            self.notifications_enabled_var.set(
                cfg.get('ui', {}).get('notifications_enabled', True))
        # QR Code
        if hasattr(self, 'qr_enabled_var'):
            qr_cfg = cfg.get('qrcode', {})
            self.qr_enabled_var.set(qr_cfg.get('enabled', False))
            self.qr_content_var.set(qr_cfg.get('content', 'nif'))
            self.qr_size_var.set(qr_cfg.get('size_mm', 25))
        # Fonts
        if hasattr(self, 'body_font_var'):
            fonts_cfg = cfg.get('fonts', {})
            self.body_font_var.set(fonts_cfg.get('body_font', 'Helvetica'))
            self.header_font_var.set(fonts_cfg.get('header_font', 'Helvetica-Bold'))
        if hasattr(self, 'fonts_tree'):
            for item in self.fonts_tree.get_children():
                self.fonts_tree.delete(item)
            for entry in cfg.get('fonts', {}).get('registered', []):
                self.fonts_tree.insert('', 'end', values=(entry.get('name', ''), entry.get('path', '')))
        # Automação
        auto_cfg = cfg.get('automation', {})
        if hasattr(self, 'watch_folder_var'):
            self.watch_folder_var.set(auto_cfg.get('watch_folder', ''))
        if hasattr(self, 'watch_enabled_var'):
            self.watch_enabled_var.set(auto_cfg.get('watch_enabled', False))
        if hasattr(self, 'watch_mode_var'):
            self.watch_mode_var.set(auto_cfg.get('watch_mode', 'individual'))
        if hasattr(self, 'watch_interval_var'):
            self.watch_interval_var.set(auto_cfg.get('watch_interval', 5))
        if hasattr(self, 'schedules_tree'):
            self._reload_schedules_tree()
        if hasattr(self, 'hooks_tree'):
            self._reload_hooks_tree()

    def run(self):
        """Inicia a aplicação."""
        self.root.mainloop()
