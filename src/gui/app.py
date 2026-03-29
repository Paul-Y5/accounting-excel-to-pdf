#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de interface gráfica do Conversor Excel → PDF.
"""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

from src.config import load_config, save_config, DEFAULT_CONFIG, list_profiles, save_profile, load_profile, delete_profile
from src.converter import ExcelToPDFConverter
from src.nif_validator import validate_nif
from src.excel_exporter import export_to_excel
from src import history
class ConverterApp:
    """Aplicação principal com interface gráfica simples para conversão de Excel para PDF."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor Excel → PDF")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Carregar configurações
        self.config = load_config()
        
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
        if path.lower().endswith(('.xlsx', '.xls')):
            self.excel_path.set(path)
            self.config.setdefault('recent', {})['last_excel_dir'] = os.path.dirname(path)
            save_config(self.config)
            self.status_var.set(f"Ficheiro carregado: {os.path.basename(path)}")
        else:
            messagebox.showwarning("Aviso", "Apenas ficheiros Excel (.xlsx, .xls) são suportados.")
        return event_data

    def _setup_ui(self):
        """Configura a interface."""
        # Estilo
        style = ttk.Style()
        style.configure('TButton', padding=6)
        style.configure('TLabel', padding=2)
        style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))
        
        # Notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Conversão
        self.tab_convert = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_convert, text='Converter')
        self._setup_convert_tab()
        
        # Tab 2: Configurações PDF
        self.tab_pdf = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_pdf, text='Página PDF')
        self._setup_pdf_tab()
        
        # Tab 3: Cabeçalho
        self.tab_header = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_header, text='Cabeçalho')
        self._setup_header_tab()
        
        # Tab 4: Tabela
        self.tab_table = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_table, text='Tabela')
        self._setup_table_tab()
        
        # Tab 5: Cores
        self.tab_colors = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_colors, text='Cores')
        self._setup_colors_tab()
        
        # Tab 6: Contabilidade
        self.tab_contab = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_contab, text='Contabilidade')
        self._setup_contabilidade_tab()
        
        # Tab 7: Dados Bancários
        self.tab_banking = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_banking, text='Dados Bancários')
        self._setup_banking_tab()

        # Tab 8: Perfis
        self.tab_profiles = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_profiles, text='Perfis')
        self._setup_profiles_tab()

        # Tab 9: Histórico
        self.tab_history = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_history, text='Histórico')
        self._setup_history_tab()
    
    def _setup_convert_tab(self):
        """Tab de conversão."""
        frame = ttk.Frame(self.tab_convert, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(frame, text="Conversor Excel → PDF", style='Header.TLabel').pack(pady=(0, 20))
        
        # Ficheiro Excel
        file_frame = ttk.LabelFrame(frame, text="Ficheiro Excel", padding=10)
        file_frame.pack(fill='x', pady=5)
        
        ttk.Entry(file_frame, textvariable=self.excel_path, width=60).pack(side='left', fill='x', expand=True)
        ttk.Button(file_frame, text="Procurar...", command=self._browse_excel).pack(side='right', padx=(10, 0))
        
        # Ficheiro de saída
        output_frame = ttk.LabelFrame(frame, text="Ficheiro PDF de Saída (opcional)", padding=10)
        output_frame.pack(fill='x', pady=5)
        
        ttk.Entry(output_frame, textvariable=self.output_path, width=60).pack(side='left', fill='x', expand=True)
        ttk.Button(output_frame, text="Procurar...", command=self._browse_output).pack(side='right', padx=(10, 0))
        
        # Opções rápidas
        options_frame = ttk.LabelFrame(frame, text="Opções", padding=10)
        options_frame.pack(fill='x', pady=5)
        
        self.auto_open_var = tk.BooleanVar(value=self.config['output']['auto_open'])
        ttk.Checkbutton(options_frame, text="Abrir PDF após conversão", 
                       variable=self.auto_open_var).pack(anchor='w')
        
        self.add_timestamp_var = tk.BooleanVar(value=self.config['output']['add_timestamp'])
        ttk.Checkbutton(options_frame, text="Adicionar data/hora ao nome do ficheiro",
                       variable=self.add_timestamp_var).pack(anchor='w')

        # Segurança
        security_frame = ttk.LabelFrame(frame, text="Segurança", padding=10)
        security_frame.pack(fill='x', pady=5)

        pw_row = ttk.Frame(security_frame)
        pw_row.pack(fill='x')
        ttk.Label(pw_row, text="Password PDF:").pack(side='left')
        self.pdf_password_var = tk.StringVar(value=self.config.get('security', {}).get('pdf_password', ''))
        ttk.Entry(pw_row, textvariable=self.pdf_password_var, width=20, show='*').pack(side='left', padx=5)

        wm_row = ttk.Frame(security_frame)
        wm_row.pack(fill='x', pady=(5, 0))
        self.watermark_enabled_var = tk.BooleanVar(value=self.config.get('watermark', {}).get('enabled', False))
        ttk.Checkbutton(wm_row, text="Marca d'água:", variable=self.watermark_enabled_var).pack(side='left')
        self.watermark_text_var = tk.StringVar(value=self.config.get('watermark', {}).get('text', 'RASCUNHO'))
        wm_combo = ttk.Combobox(wm_row, textvariable=self.watermark_text_var, width=15,
                                values=['RASCUNHO', 'CÓPIA', 'CONFIDENCIAL', 'ORIGINAL'])
        wm_combo.pack(side='left', padx=5)

        # Modo de geração
        mode_frame = ttk.LabelFrame(frame, text="Modo de Geração", padding=10)
        mode_frame.pack(fill='x', pady=5)

        self.generation_mode_var = tk.StringVar(value='individual')  # Default: por linha
        ttk.Radiobutton(mode_frame, text="Por Linha (um PDF por cliente)",
                       variable=self.generation_mode_var, value='individual').pack(anchor='w')
        ttk.Radiobutton(mode_frame, text="Agregado (todos num único PDF)",
                       variable=self.generation_mode_var, value='aggregate').pack(anchor='w')

        # Filtro de clientes
        filter_row = ttk.Frame(mode_frame)
        filter_row.pack(fill='x', pady=(5, 0))
        ttk.Button(filter_row, text="Filtrar Clientes...", command=self._open_client_filter).pack(side='left')
        self.client_filter_label = ttk.Label(filter_row, text="Todos os clientes", foreground='gray')
        self.client_filter_label.pack(side='left', padx=10)
        self._client_filter = None  # None = todos, set() = filtrados

        # Botões
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=20)

        preview_btn = ttk.Button(btn_frame, text="Pré-visualizar",
                                command=self._preview_excel, style='TButton')
        preview_btn.pack(side='left', padx=5)

        generate_btn = ttk.Button(btn_frame, text="Gerar PDF(s)",
                                 command=self._generate, style='TButton')
        generate_btn.pack(side='left', padx=5)

        export_excel_btn = ttk.Button(btn_frame, text="Exportar Excel",
                                      command=self._export_excel, style='TButton')
        export_excel_btn.pack(side='left', padx=5)

        ttk.Button(btn_frame, text="Guardar Configurações",
                  command=self._save_config).pack(side='left', padx=5)

        # Segunda linha de botões
        btn_frame2 = ttk.Frame(frame)
        btn_frame2.pack(pady=(0, 10))

        ttk.Button(btn_frame2, text="Abrir Pasta de Destino",
                  command=self._open_output_folder).pack(side='left', padx=5)
        ttk.Button(btn_frame2, text="Resumo IRS",
                  command=self._show_irs_summary).pack(side='left', padx=5)

        # Barra de progresso
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.pack(fill='x', pady=(5, 0))

        # Status
        self.status_var = tk.StringVar(value="Pronto para converter  |  Ctrl+O: Abrir  Ctrl+G: Gerar  Ctrl+S: Guardar")
        ttk.Label(frame, textvariable=self.status_var, foreground='gray').pack(pady=5)
    
    def _setup_pdf_tab(self):
        """Tab de configurações do PDF."""
        frame = ttk.Frame(self.tab_pdf, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Tamanho da página
        size_frame = ttk.LabelFrame(frame, text="Tamanho da Página", padding=10)
        size_frame.pack(fill='x', pady=5)
        
        self.page_size_var = tk.StringVar(value=self.config['pdf']['page_size'])
        ttk.Label(size_frame, text="Tamanho:").grid(row=0, column=0, sticky='w', padx=5)
        ttk.Combobox(size_frame, textvariable=self.page_size_var, 
                    values=['A4', 'A3', 'Letter'], width=15, state='readonly').grid(row=0, column=1, padx=5)
        
        self.orientation_var = tk.StringVar(value=self.config['pdf']['orientation'])
        ttk.Label(size_frame, text="Orientação:").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Combobox(size_frame, textvariable=self.orientation_var, 
                    values=['portrait', 'landscape'], width=15, state='readonly').grid(row=0, column=3, padx=5)
        
        # Margens
        margin_frame = ttk.LabelFrame(frame, text="Margens (mm)", padding=10)
        margin_frame.pack(fill='x', pady=5)
        
        self.margin_top_var = tk.IntVar(value=self.config['pdf']['margin_top'])
        self.margin_bottom_var = tk.IntVar(value=self.config['pdf']['margin_bottom'])
        self.margin_left_var = tk.IntVar(value=self.config['pdf']['margin_left'])
        self.margin_right_var = tk.IntVar(value=self.config['pdf']['margin_right'])
        
        ttk.Label(margin_frame, text="Superior:").grid(row=0, column=0, sticky='w', padx=5)
        ttk.Spinbox(margin_frame, textvariable=self.margin_top_var, from_=5, to=50, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(margin_frame, text="Inferior:").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Spinbox(margin_frame, textvariable=self.margin_bottom_var, from_=5, to=50, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(margin_frame, text="Esquerda:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(margin_frame, textvariable=self.margin_left_var, from_=5, to=50, width=8).grid(row=1, column=1, padx=5)
        
        ttk.Label(margin_frame, text="Direita:").grid(row=1, column=2, sticky='w', padx=5, pady=5)
        ttk.Spinbox(margin_frame, textvariable=self.margin_right_var, from_=5, to=50, width=8).grid(row=1, column=3, padx=5)
        
        # Botão Guardar
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(pady=20)
    
    def _setup_header_tab(self):
        """Tab de configurações do cabeçalho."""
        frame = ttk.Frame(self.tab_header, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Mostrar cabeçalho
        self.show_header_var = tk.BooleanVar(value=self.config['header']['show_header'])
        ttk.Checkbutton(frame, text="Mostrar cabeçalho no PDF", 
                       variable=self.show_header_var).pack(anchor='w', pady=5)
        
        # Dados da empresa
        company_frame = ttk.LabelFrame(frame, text="Dados da Empresa (valores padrão)", padding=10)
        company_frame.pack(fill='x', pady=10)
        
        self.company_name_var = tk.StringVar(value=self.config['header']['company_name'])
        self.company_address_var = tk.StringVar(value=self.config['header']['company_address'])
        self.company_phone_var = tk.StringVar(value=self.config['header']['company_phone'])
        self.company_email_var = tk.StringVar(value=self.config['header']['company_email'])
        self.company_nif_var = tk.StringVar(value=self.config['header']['company_nif'])
        
        fields = [
            ("Nome:", self.company_name_var),
            ("Morada:", self.company_address_var),
            ("Telefone:", self.company_phone_var),
            ("Email:", self.company_email_var),
            ("NIF:", self.company_nif_var),
        ]
        
        for i, (label, var) in enumerate(fields):
            ttk.Label(company_frame, text=label).grid(row=i, column=0, sticky='w', padx=5, pady=2)
            ttk.Entry(company_frame, textvariable=var, width=50).grid(row=i, column=1, sticky='ew', padx=5, pady=2)
        
        company_frame.columnconfigure(1, weight=1)
        
        # Logo
        logo_frame = ttk.LabelFrame(frame, text="Logo (opcional)", padding=10)
        logo_frame.pack(fill='x', pady=10)
        
        self.logo_path_var = tk.StringVar(value=self.config['header'].get('logo_path', ''))
        ttk.Entry(logo_frame, textvariable=self.logo_path_var, width=50).pack(side='left', fill='x', expand=True)
        ttk.Button(logo_frame, text="Procurar...", command=self._browse_logo).pack(side='right', padx=(10, 0))
        
        # Botão Guardar
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(pady=20)
    
    def _setup_table_tab(self):
        """Tab de configurações da tabela."""
        frame = ttk.Frame(self.tab_table, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Fontes
        font_frame = ttk.LabelFrame(frame, text="Tamanho de Fonte", padding=10)
        font_frame.pack(fill='x', pady=5)
        
        self.font_size_var = tk.IntVar(value=self.config['table']['font_size'])
        self.header_font_size_var = tk.IntVar(value=self.config['table']['header_font_size'])
        self.row_padding_var = tk.IntVar(value=self.config['table']['row_padding'])
        
        ttk.Label(font_frame, text="Texto:").grid(row=0, column=0, sticky='w', padx=5)
        ttk.Spinbox(font_frame, textvariable=self.font_size_var, from_=6, to=14, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(font_frame, text="Cabeçalho:").grid(row=0, column=2, sticky='w', padx=5)
        ttk.Spinbox(font_frame, textvariable=self.header_font_size_var, from_=8, to=16, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(font_frame, text="Espaço:").grid(row=0, column=4, sticky='w', padx=5)
        ttk.Spinbox(font_frame, textvariable=self.row_padding_var, from_=2, to=12, width=8).grid(row=0, column=5, padx=5)
        
        # Opções
        options_frame = ttk.LabelFrame(frame, text="Opções da Tabela", padding=10)
        options_frame.pack(fill='x', pady=5)
        
        self.show_grid_var = tk.BooleanVar(value=self.config['table']['show_grid'])
        self.alternate_rows_var = tk.BooleanVar(value=self.config['table']['alternate_rows'])
        
        ttk.Checkbutton(options_frame, text="Mostrar grelha/bordas", 
                       variable=self.show_grid_var).pack(anchor='w')
        ttk.Checkbutton(options_frame, text="Cores alternadas nas linhas", 
                       variable=self.alternate_rows_var).pack(anchor='w')
        
        # Rodapé
        footer_frame = ttk.LabelFrame(frame, text="Rodapé", padding=10)
        footer_frame.pack(fill='x', pady=5)
        
        self.show_signatures_var = tk.BooleanVar(value=self.config['footer']['show_signatures'])
        self.show_date_var = tk.BooleanVar(value=self.config['footer']['show_date'])
        self.show_observations_var = tk.BooleanVar(value=self.config['footer']['show_observations'])
        
        ttk.Checkbutton(footer_frame, text="Mostrar área de assinaturas", 
                       variable=self.show_signatures_var).pack(anchor='w')
        ttk.Checkbutton(footer_frame, text="Mostrar data de geração", 
                       variable=self.show_date_var).pack(anchor='w')
        ttk.Checkbutton(footer_frame, text="Mostrar observações", 
                       variable=self.show_observations_var).pack(anchor='w')
        
        ttk.Label(footer_frame, text="Texto personalizado no rodapé:").pack(anchor='w', pady=(10, 0))
        self.custom_footer_var = tk.StringVar(value=self.config['footer'].get('custom_footer', ''))
        ttk.Entry(footer_frame, textvariable=self.custom_footer_var, width=60).pack(fill='x', pady=5)
        
        # Botão Guardar
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(pady=20)
    
    def _setup_colors_tab(self):
        """Tab de configurações de cores."""
        frame = ttk.Frame(self.tab_colors, padding=20)
        frame.pack(fill='both', expand=True)
        
        self.color_vars = {}
        
        colors_config = [
            ('header_bg', 'Fundo do cabeçalho da tabela'),
            ('header_text', 'Texto do cabeçalho da tabela'),
            ('row_alt', 'Cor alternada das linhas'),
            ('border', 'Cor das bordas'),
            ('title', 'Cor do título da empresa'),
        ]
        
        for key, label in colors_config:
            row_frame = ttk.Frame(frame)
            row_frame.pack(fill='x', pady=5)
            
            ttk.Label(row_frame, text=label, width=30).pack(side='left')
            
            color_value = self.config['colors'].get(key, '#000000')
            var = tk.StringVar(value=color_value)
            self.color_vars[key] = var
            
            color_entry = ttk.Entry(row_frame, textvariable=var, width=15)
            color_entry.pack(side='left', padx=5)
            
            color_btn = tk.Button(row_frame, text="  ", bg=color_value, width=3,
                                 command=lambda k=key, v=var, b=None: self._pick_color(k, v))
            color_btn.pack(side='left')
            self.color_vars[f'{key}_btn'] = color_btn
        
        # Botão Guardar
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(pady=20)
    
    def _setup_contabilidade_tab(self):
        """Tab de configurações de contabilidade."""
        frame = ttk.Frame(self.tab_contab, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(frame, text="Configurações de Contabilidade", style='Header.TLabel').pack(pady=(0, 15))
        
        # Descrição
        desc_text = "Configure quais colunas do Excel serão incluídas no PDF de contabilidade.\nSepare as colunas por vírgula, na ordem desejada."
        ttk.Label(frame, text=desc_text, foreground='gray').pack(pady=(0, 10))
        
        # Colunas
        colunas_frame = ttk.LabelFrame(frame, text="Colunas a Incluir", padding=10)
        colunas_frame.pack(fill='x', pady=10)
        
        contab_cfg = self.config.get('contabilidade', {})
        default_colunas = 'Nr., SIGLA, Cliente, CONTAB, Iva, Subtotal, Extras, Duodécimos, S.Social GER, S.Soc Emp, Ret. IRS, Ret. IRS EXT, SbTx/Fcomp, Outro, TOTAL'
        
        self.contab_colunas_var = tk.StringVar(value=contab_cfg.get('colunas', default_colunas))
        
        ttk.Label(colunas_frame, text="Lista de colunas (separadas por vírgula):").pack(anchor='w', pady=(0, 5))
        
        # Text widget para permitir múltiplas linhas
        self.contab_colunas_text = tk.Text(colunas_frame, height=4, width=70, wrap='word')
        self.contab_colunas_text.pack(fill='x', pady=5)
        self.contab_colunas_text.insert('1.0', self.contab_colunas_var.get())
        
        # Botão para restaurar padrão
        def reset_colunas():
            self.contab_colunas_text.delete('1.0', tk.END)
            self.contab_colunas_text.insert('1.0', default_colunas)
        
        ttk.Button(colunas_frame, text="Restaurar Padrão", command=reset_colunas).pack(anchor='e', pady=5)
        
        # Opções de destaque
        options_frame = ttk.LabelFrame(frame, text="Opções de Formatação", padding=10)
        options_frame.pack(fill='x', pady=10)
        
        self.contab_destacar_total_var = tk.BooleanVar(value=contab_cfg.get('destacar_total', True))
        ttk.Checkbutton(options_frame, text="Destacar coluna TOTAL com cor de fundo", 
                       variable=self.contab_destacar_total_var).pack(anchor='w')
        
        self.contab_destacar_valores_var = tk.BooleanVar(value=contab_cfg.get('destacar_valores', True))
        ttk.Checkbutton(options_frame, text="Destacar valores (positivos/negativos)", 
                       variable=self.contab_destacar_valores_var).pack(anchor='w')
        
        # Exemplos de colunas possíveis
        examples_frame = ttk.LabelFrame(frame, text="Colunas Disponíveis (exemplos)", padding=10)
        examples_frame.pack(fill='x', pady=10)
        
        examples = [
            "Nr. - Número do cliente",
            "SIGLA - Sigla do cliente",
            "Cliente - Nome do cliente",
            "CONTAB - Valor de contabilidade",
            "Iva - Valor do IVA",
            "Subtotal - Subtotal",
            "Extras - Valores extras",
            "Duodécimos - Duodécimos",
            "S.Social GER - Segurança Social (Gerente)",
            "S.Soc Emp - Segurança Social (Empresa)",
            "Ret. IRS - Retenção IRS",
            "Ret. IRS EXT - Retenção IRS Exterior",
            "SbTx/Fcomp - Subsídios/Férias",
            "Outro - Outros valores",
            "TOTAL - Total calculado",
        ]
        
        examples_text = "\n".join(examples)
        ttk.Label(examples_frame, text=examples_text, foreground='gray', justify='left').pack(anchor='w')
        
        # Botão Guardar
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(pady=20)
    
    def _setup_banking_tab(self):
        """Tab de configurações de dados bancários (múltiplas contas)."""
        frame = ttk.Frame(self.tab_banking, padding=20)
        frame.pack(fill='both', expand=True)

        # Título
        ttk.Label(frame, text="Dados Bancários", style='Header.TLabel').pack(pady=(0, 15))

        desc_text = "Configure as contas bancárias que aparecerão no rodapé do PDF.\nA conta marcada como predefinida será usada na geração."
        ttk.Label(frame, text=desc_text, foreground='gray').pack(pady=(0, 10))

        # Mostrar dados bancários
        banking_cfg = self.config.get('banking', {})
        self.show_banking_var = tk.BooleanVar(value=banking_cfg.get('show_banking', True))
        ttk.Checkbutton(frame, text="Mostrar dados bancários no PDF",
                       variable=self.show_banking_var).pack(anchor='w', pady=5)

        # Título bancário
        title_row = ttk.Frame(frame)
        title_row.pack(fill='x', pady=5)
        ttk.Label(title_row, text="Título:").pack(side='left')
        self.banking_title_var = tk.StringVar(value=banking_cfg.get('title', 'Nossos Dados Bancários:'))
        ttk.Entry(title_row, textvariable=self.banking_title_var, width=40).pack(side='left', padx=5)

        # Lista de contas
        accounts_frame = ttk.LabelFrame(frame, text="Contas Bancárias", padding=10)
        accounts_frame.pack(fill='both', expand=True, pady=10)

        # Treeview para contas
        tree_frame = ttk.Frame(accounts_frame)
        tree_frame.pack(fill='both', expand=True)

        cols = ('banco', 'iban', 'predefinida')
        self.accounts_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=5)
        self.accounts_tree.heading('banco', text='Banco')
        self.accounts_tree.heading('iban', text='IBAN')
        self.accounts_tree.heading('predefinida', text='Predefinida')
        self.accounts_tree.column('banco', width=120)
        self.accounts_tree.column('iban', width=300)
        self.accounts_tree.column('predefinida', width=80)
        self.accounts_tree.pack(fill='both', expand=True)

        # Carregar contas existentes
        accounts = banking_cfg.get('accounts', [])
        for acc in accounts:
            default_mark = 'Sim' if acc.get('default', False) else ''
            self.accounts_tree.insert('', 'end', values=(
                acc.get('bank_name', ''),
                acc.get('iban', ''),
                default_mark,
            ))

        # Botões de gestão de contas
        acc_btn_frame = ttk.Frame(accounts_frame)
        acc_btn_frame.pack(fill='x', pady=(10, 0))

        ttk.Button(acc_btn_frame, text="Adicionar", command=self._add_bank_account).pack(side='left', padx=5)
        ttk.Button(acc_btn_frame, text="Remover", command=self._remove_bank_account).pack(side='left', padx=5)
        ttk.Button(acc_btn_frame, text="Definir como Predefinida", command=self._set_default_account).pack(side='left', padx=5)

        # Botão Guardar
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(pady=20)
    
    def _add_bank_account(self):
        """Adiciona uma nova conta bancária via popup."""
        popup = tk.Toplevel(self.root)
        popup.title("Adicionar Conta Bancária")
        popup.geometry("400x180")
        popup.transient(self.root)
        popup.grab_set()

        f = ttk.Frame(popup, padding=15)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text="Nome do Banco:").grid(row=0, column=0, sticky='w', pady=5)
        bank_var = tk.StringVar()
        ttk.Entry(f, textvariable=bank_var, width=35).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(f, text="IBAN:").grid(row=1, column=0, sticky='w', pady=5)
        iban_var = tk.StringVar()
        ttk.Entry(f, textvariable=iban_var, width=35).grid(row=1, column=1, padx=5, pady=5)

        def confirm():
            bank = bank_var.get().strip()
            iban = iban_var.get().strip()
            if not bank or not iban:
                messagebox.showwarning("Aviso", "Preencha o nome do banco e o IBAN.", parent=popup)
                return
            self.accounts_tree.insert('', 'end', values=(bank, iban, ''))
            popup.destroy()

        ttk.Button(f, text="Adicionar", command=confirm).grid(row=2, column=1, sticky='e', pady=15)

    def _remove_bank_account(self):
        """Remove a conta bancária selecionada."""
        selected = self.accounts_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma conta para remover.")
            return
        for item in selected:
            self.accounts_tree.delete(item)

    def _set_default_account(self):
        """Define a conta selecionada como predefinida."""
        selected = self.accounts_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma conta para definir como predefinida.")
            return
        # Limpar todas as marcas de predefinida
        for item in self.accounts_tree.get_children():
            vals = list(self.accounts_tree.item(item, 'values'))
            vals[2] = ''
            self.accounts_tree.item(item, values=vals)
        # Marcar a selecionada
        vals = list(self.accounts_tree.item(selected[0], 'values'))
        vals[2] = 'Sim'
        self.accounts_tree.item(selected[0], values=vals)

    def _open_client_filter(self):
        """Abre janela para selecionar clientes a incluir no PDF."""
        excel_path = self.excel_path.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Erro", "Selecione um ficheiro Excel primeiro.")
            return

        try:
            config = self._get_config_from_ui()
            converter = ExcelToPDFConverter(excel_path, None, config)
            data = converter.read_excel_data()
            itens = data.get('itens', [])
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler Excel:\n{e}")
            return

        if not itens:
            messagebox.showwarning("Aviso", "Sem dados no Excel.")
            return

        # Obter lista de clientes únicos
        clients = []
        seen = set()
        for item in itens:
            name = item.get('Cliente', '')
            if name and name not in seen:
                clients.append(name)
                seen.add(name)

        # Popup de seleção
        popup = tk.Toplevel(self.root)
        popup.title("Filtrar Clientes")
        popup.geometry("450x500")
        popup.transient(self.root)
        popup.grab_set()

        f = ttk.Frame(popup, padding=10)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text=f"{len(clients)} clientes encontrados. Selecione os que deseja incluir:",
                 font=('Helvetica', 10)).pack(anchor='w', pady=(0, 10))

        # Botões selecionar/desselecionar todos
        sel_frame = ttk.Frame(f)
        sel_frame.pack(fill='x', pady=(0, 5))

        check_vars = {}
        list_frame = ttk.Frame(f)
        list_frame.pack(fill='both', expand=True)

        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        scroll_content = ttk.Frame(canvas)

        scroll_content.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_content, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        for client in clients:
            var = tk.BooleanVar(value=(self._client_filter is None or client in self._client_filter))
            check_vars[client] = var
            ttk.Checkbutton(scroll_content, text=client, variable=var).pack(anchor='w', padx=5, pady=1)

        def select_all():
            for v in check_vars.values():
                v.set(True)

        def deselect_all():
            for v in check_vars.values():
                v.set(False)

        ttk.Button(sel_frame, text="Selecionar Todos", command=select_all).pack(side='left', padx=5)
        ttk.Button(sel_frame, text="Desselecionar Todos", command=deselect_all).pack(side='left', padx=5)

        def apply_filter():
            selected = {name for name, var in check_vars.items() if var.get()}
            if len(selected) == len(clients):
                self._client_filter = None
                self.client_filter_label.config(text="Todos os clientes")
            elif len(selected) == 0:
                messagebox.showwarning("Aviso", "Selecione pelo menos um cliente.", parent=popup)
                return
            else:
                self._client_filter = selected
                self.client_filter_label.config(text=f"{len(selected)} de {len(clients)} clientes")
            popup.destroy()

        btn_frame = ttk.Frame(f)
        btn_frame.pack(fill='x', pady=(10, 0))
        ttk.Button(btn_frame, text="Aplicar", command=apply_filter).pack(side='right', padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=popup.destroy).pack(side='right', padx=5)

    def _open_output_folder(self):
        """Abre a pasta de destino no explorador de ficheiros."""
        output_path = self.output_path.get()
        if output_path:
            folder = os.path.dirname(output_path)
        else:
            folder = self.config.get('recent', {}).get('last_output_dir', '')

        if not folder or not os.path.isdir(folder):
            messagebox.showinfo("Info", "Nenhuma pasta de destino definida.\nGere um PDF primeiro ou defina o caminho de saída.")
            return

        if sys.platform == 'linux':
            subprocess.Popen(['xdg-open', folder])
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', folder])
        else:
            os.startfile(folder)

    def _pick_color(self, key, var):
        """Abre seletor de cor."""
        color = colorchooser.askcolor(initialcolor=var.get())
        if color[1]:
            var.set(color[1])
            if f'{key}_btn' in self.color_vars:
                self.color_vars[f'{key}_btn'].configure(bg=color[1])
    
    def _browse_excel(self):
        """Seleciona ficheiro Excel, lembrando a última pasta usada."""
        initial_dir = self.config.get('recent', {}).get('last_excel_dir', '')
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser('~')

        path = filedialog.askopenfilename(
            title="Selecionar ficheiro Excel",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.excel_path.set(path)
            # Guardar última pasta
            self.config.setdefault('recent', {})['last_excel_dir'] = os.path.dirname(path)
            save_config(self.config)
    
    def _browse_output(self):
        """Seleciona ficheiro de saída, lembrando a última pasta."""
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
    
    def _browse_logo(self):
        """Seleciona ficheiro de logo."""
        path = filedialog.askopenfilename(
            title="Selecionar logo",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif"), ("All files", "*.*")]
        )
        if path:
            self.logo_path_var.set(path)
    
    def _load_config_to_ui(self):
        """Carrega configurações para a UI."""
        # Já feito nos setup_*_tab através dos valores padrão
        pass
    
    def _get_config_from_ui(self) -> dict:
        """Obtém configurações da UI."""
        # Obter texto das colunas de contabilidade
        contab_colunas = self.contab_colunas_text.get('1.0', tk.END).strip() if hasattr(self, 'contab_colunas_text') else ''
        
        return {
            'pdf': {
                'page_size': self.page_size_var.get(),
                'orientation': self.orientation_var.get(),
                'margin_top': self.margin_top_var.get(),
                'margin_bottom': self.margin_bottom_var.get(),
                'margin_left': self.margin_left_var.get(),
                'margin_right': self.margin_right_var.get(),
            },
            'header': {
                'show_header': self.show_header_var.get(),
                'company_name': self.company_name_var.get(),
                'company_address': self.company_address_var.get(),
                'company_phone': self.company_phone_var.get(),
                'company_email': self.company_email_var.get(),
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
            },
            'contabilidade': {
                'enabled': True,
                'colunas': contab_colunas,
                'destacar_total': self.contab_destacar_total_var.get() if hasattr(self, 'contab_destacar_total_var') else True,
                'destacar_valores': self.contab_destacar_valores_var.get() if hasattr(self, 'contab_destacar_valores_var') else True,
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
            'banking': self._get_banking_from_ui(),
            'recent': self.config.get('recent', {'last_excel_dir': '', 'last_output_dir': ''}),
        }
    
    def _get_banking_from_ui(self) -> dict:
        """Lê as contas bancárias do Treeview."""
        accounts = []
        if hasattr(self, 'accounts_tree'):
            for item in self.accounts_tree.get_children():
                vals = self.accounts_tree.item(item, 'values')
                accounts.append({
                    'bank_name': vals[0],
                    'iban': vals[1],
                    'default': vals[2] == 'Sim',
                })
        if not accounts:
            accounts = [{'bank_name': 'ABANCA', 'iban': 'PT50 0170 3782 0304 0053 5672 9', 'default': True}]
        return {
            'show_banking': self.show_banking_var.get() if hasattr(self, 'show_banking_var') else True,
            'title': self.banking_title_var.get() if hasattr(self, 'banking_title_var') else 'Nossos Dados Bancários:',
            'accounts': accounts,
        }

    def _save_config(self):
        """Guarda configurações."""
        self.config = self._get_config_from_ui()
        save_config(self.config)
        messagebox.showinfo("Sucesso", "Configurações guardadas com sucesso!")
    
    def _generate(self):
        """Executa a geração conforme o modo selecionado."""
        mode = self.generation_mode_var.get()
        
        if mode == 'individual':
            self._convert_individual()
        else:
            self._convert()
    
    def _convert(self):
        """Executa a conversão (modo agregado)."""
        excel_path = self.excel_path.get()

        if not excel_path:
            messagebox.showerror("Erro", "Por favor, selecione um ficheiro Excel.")
            return

        if not os.path.exists(excel_path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado: {excel_path}")
            return

        config = self._get_config_from_ui()
        output_path = self.output_path.get() or None

        converter = ExcelToPDFConverter(excel_path, output_path, config)

        # Confirmar sobrescrita
        if os.path.exists(converter.output_pdf_path):
            if not messagebox.askyesno("Confirmar",
                    f"O ficheiro já existe:\n{converter.output_pdf_path}\n\nDeseja substituir?"):
                self.status_var.set("Conversão cancelada")
                return

        self.progress_var.set(10)
        self.status_var.set("A ler dados do Excel...")
        self.root.update()

        def task():
            try:
                data = converter.read_excel_data()
                self.root.after(0, lambda: self.progress_var.set(40))
                clients_count = len(data.get('itens', []))

                self.root.after(0, lambda: self.status_var.set("A gerar PDF..."))
                self.root.after(0, lambda: self.progress_var.set(60))
                result_path = converter.generate_pdf(client_filter=self._client_filter)

                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(0, lambda: self.status_var.set(
                    f"PDF gerado: {os.path.basename(result_path)} ({clients_count} clientes)"))

                history.add_entry(excel_path, result_path, 'aggregate', clients_count, True)

                self.root.after(0, lambda: messagebox.showinfo("Sucesso",
                    f"PDF gerado com sucesso!\n\n{result_path}\n\nClientes: {clients_count}"))

                if config['output'].get('auto_open', True):
                    if sys.platform == 'linux':
                        subprocess.Popen(['xdg-open', result_path])
                    elif sys.platform == 'darwin':
                        subprocess.Popen(['open', result_path])
                    else:
                        os.startfile(result_path)

                self.root.after(1500, lambda: self.progress_var.set(0))

            except Exception as e:
                self.root.after(0, lambda: self.progress_var.set(0))
                self.root.after(0, lambda: self.status_var.set("Erro na conversão"))
                history.add_entry(excel_path, output_path or '', 'aggregate', 0, False, str(e))
                self.root.after(0, lambda: messagebox.showerror("Erro",
                    f"Erro durante a conversão:\n\n{str(e)}"))

        threading.Thread(target=task, daemon=True).start()
    
    def _convert_individual(self):
        """Gera PDFs individuais para cada cliente."""
        excel_path = self.excel_path.get()

        if not excel_path:
            messagebox.showerror("Erro", "Por favor, selecione um ficheiro Excel.")
            return

        if not os.path.exists(excel_path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado: {excel_path}")
            return

        config = self._get_config_from_ui()
        self.progress_var.set(10)
        self.status_var.set("A gerar PDFs individuais...")
        self.root.update()

        def task():
            try:
                converter = ExcelToPDFConverter(excel_path, None, config)

                self.root.after(0, lambda: self.progress_var.set(30))
                result_files = converter.generate_individual_pdfs(client_filter=self._client_filter)

                self.root.after(0, lambda: self.progress_var.set(100))

                if result_files:
                    folder = os.path.dirname(result_files[0])
                    self.root.after(0, lambda: self.status_var.set(
                        f"{len(result_files)} PDFs gerados!"))

                    history.add_entry(excel_path, folder, 'individual', len(result_files), True)

                    self.root.after(0, lambda: messagebox.showinfo("Sucesso",
                        f"Gerados {len(result_files)} PDFs individuais!\n\n"
                        f"Pasta: {folder}"))

                    if config['output'].get('auto_open', True):
                        if sys.platform == 'linux':
                            subprocess.Popen(['xdg-open', folder])
                        elif sys.platform == 'darwin':
                            subprocess.Popen(['open', folder])
                        else:
                            os.startfile(folder)
                else:
                    self.root.after(0, lambda: self.status_var.set("Nenhum PDF gerado"))
                    self.root.after(0, lambda: messagebox.showwarning("Aviso",
                        "Nenhum item encontrado para gerar PDFs."))

                self.root.after(1500, lambda: self.progress_var.set(0))

            except Exception as e:
                self.root.after(0, lambda: self.progress_var.set(0))
                self.root.after(0, lambda: self.status_var.set("Erro na conversão"))
                history.add_entry(excel_path, '', 'individual', 0, False, str(e))
                self.root.after(0, lambda: messagebox.showerror("Erro",
                    f"Erro durante a geração:\n\n{str(e)}"))

        threading.Thread(target=task, daemon=True).start()
    
    def _preview_excel(self):
        """Mostra pré-visualização dos dados do Excel antes de gerar PDF."""
        excel_path = self.excel_path.get()
        
        if not excel_path:
            messagebox.showerror("Erro", "Por favor, selecione um ficheiro Excel.")
            return
        
        if not os.path.exists(excel_path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado: {excel_path}")
            return
        
        try:
            self.status_var.set("A carregar pré-visualização...")
            self.root.update()
            
            # Ler dados do Excel
            config = self._get_config_from_ui()
            converter = ExcelToPDFConverter(excel_path, None, config)
            data = converter.read_excel_data()
            itens = data.get('itens', [])
            
            if not itens:
                messagebox.showwarning("Aviso", "O ficheiro Excel não contém dados para converter.")
                self.status_var.set("Pronto para converter")
                return
            
            # Criar janela de pré-visualização
            preview_window = tk.Toplevel(self.root)
            preview_window.title(f"Pré-visualização: {os.path.basename(excel_path)}")
            preview_window.geometry("900x600")
            preview_window.transient(self.root)
            preview_window.grab_set()
            
            # Frame principal
            main_frame = ttk.Frame(preview_window, padding=10)
            main_frame.pack(fill='both', expand=True)
            
            # Resumo
            summary_frame = ttk.LabelFrame(main_frame, text="Resumo", padding=10)
            summary_frame.pack(fill='x', pady=(0, 10))
            
            # Obter colunas
            all_cols = set()
            for item in itens:
                all_cols.update(item.keys())
            
            mes_ref = data.get('mes_referencia', 'N/A')
            mode_text = "Individual (1 PDF por linha)" if self.generation_mode_var.get() == 'individual' else "Agregado (1 único PDF)"
            
            # === VALIDAÇÃO DE DADOS ===
            warnings = []
            rows_with_issues = []

            for idx, item in enumerate(itens):
                row_issues = []

                # Verificar Cliente vazio
                cliente = item.get('Cliente', '')
                if not cliente or str(cliente).strip() == '':
                    row_issues.append("Cliente vazio")

                # Verificar SIGLA vazia
                sigla = item.get('SIGLA', '')
                if not sigla or str(sigla).strip() == '':
                    row_issues.append("SIGLA vazia")

                # Validação de NIF
                nif = item.get('NIF', '')
                if nif and str(nif).strip():
                    is_valid, nif_msg = validate_nif(str(nif))
                    if not is_valid:
                        row_issues.append(f"NIF inválido ({nif_msg})")

                # Verificar TOTAL = 0 ou vazio
                total = item.get('TOTAL', 0)
                if total == 0 or total == '' or total is None:
                    row_issues.append("TOTAL é 0 ou vazio")

                # Verificar valores negativos inesperados
                for field in ['CONTAB', 'Subtotal']:
                    val = item.get(field, 0)
                    if isinstance(val, (int, float)) and val < 0:
                        row_issues.append(f"{field} negativo")
                
                if row_issues:
                    nr = item.get('Nr.', idx + 1)
                    # Mostrar identificação mais clara: Nr + SIGLA ou Cliente
                    sigla_display = item.get('SIGLA', '') or ''
                    cliente_display = item.get('Cliente', '') or ''
                    
                    if sigla_display:
                        identificador = f"{nr} ({sigla_display})"
                    elif cliente_display:
                        # Truncar nome se muito longo
                        nome_curto = cliente_display[:25] + "..." if len(cliente_display) > 25 else cliente_display
                        identificador = f"{nr} - {nome_curto}"
                    else:
                        identificador = str(nr)
                    
                    warnings.append(f"{identificador}: {', '.join(row_issues)}")
                    rows_with_issues.append(idx)
            
            summary_text = f"📊 Total de registos: {len(itens)}  |  📋 Colunas: {len(all_cols)}  |  📅 Mês: {mes_ref}  |  📄 Modo: {mode_text}"
            ttk.Label(summary_frame, text=summary_text, font=('Helvetica', 10)).pack(anchor='w')
            
            # Mostrar alertas de validação (se houver)
            if warnings:
                warning_frame = ttk.LabelFrame(main_frame, text=f"⚠️ Alertas de Validação ({len(warnings)})", padding=10)
                warning_frame.pack(fill='x', pady=(0, 10))
                
                # Mostrar até 5 avisos
                warning_display = warnings[:5]
                warning_text = "\n".join(warning_display)
                warning_label = ttk.Label(warning_frame, text=warning_text, foreground='#b45309', 
                                         font=('Helvetica', 9), justify='left')
                warning_label.pack(anchor='w')
                
                # Se houver mais de 5, mostrar link clicável
                if len(warnings) > 5:
                    # Capturar warnings numa variável local para o closure
                    all_warnings_list = list(warnings)
                    
                    def show_all_warnings(warnings_to_show=all_warnings_list):
                        """Mostra todos os alertas numa janela popup."""
                        popup = tk.Toplevel(preview_window)
                        popup.title(f"Todos os Alertas ({len(warnings_to_show)})")
                        popup.geometry("600x400")
                        popup.transient(preview_window)
                        
                        # Frame principal
                        popup_frame = tk.Frame(popup, bg='#fffbeb', padx=10, pady=10)
                        popup_frame.pack(fill='both', expand=True)
                        
                        # Label título
                        tk.Label(popup_frame, text=f"⚠️ {len(warnings_to_show)} alertas encontrados:", 
                                font=('Helvetica', 11, 'bold'), bg='#fffbeb', fg='#92400e').pack(anchor='w', pady=(0, 10))
                        
                        # Frame para lista + scrollbar
                        list_frame = tk.Frame(popup_frame, bg='#fffbeb')
                        list_frame.pack(fill='both', expand=True)
                        
                        # Scrollbar
                        scrollbar = tk.Scrollbar(list_frame)
                        scrollbar.pack(side='right', fill='y')
                        
                        # Listbox (mais fiável que Text widget)
                        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set,
                                            font=('Helvetica', 10), fg='#92400e', bg='#fffbeb',
                                            selectbackground='#fcd34d', highlightthickness=0,
                                            relief='flat', activestyle='none')
                        listbox.pack(side='left', fill='both', expand=True)
                        scrollbar.config(command=listbox.yview)
                        
                        # Inserir todos os warnings
                        for i, w in enumerate(warnings_to_show, 1):
                            listbox.insert('end', f"  {i}. {w}")
                        
                        # Botão fechar
                        tk.Button(popup_frame, text="Fechar", command=popup.destroy, 
                                 bg='#f59e0b', fg='white', font=('Helvetica', 10),
                                 padx=20, pady=5, relief='flat', cursor='hand2').pack(pady=(10, 0))
                        
                        popup.grab_set()
                        popup.update()
                    
                    # Link clicável
                    more_link = tk.Label(warning_frame, text=f"👆 Ver todos os {len(warnings)} alertas...", 
                                        fg='#2563eb', cursor='hand2', font=('Helvetica', 9, 'underline'))
                    more_link.pack(anchor='w', pady=(5, 0))
                    more_link.bind('<Button-1>', lambda e: show_all_warnings())
                
                ttk.Label(warning_frame, text="ℹ️ Pode gerar os PDFs mesmo com alertas.", 
                         foreground='gray', font=('Helvetica', 8)).pack(anchor='w', pady=(5, 0))
            
            # Criar Treeview para mostrar dados
            tree_frame = ttk.Frame(main_frame)
            tree_frame.pack(fill='both', expand=True)
            
            # Scrollbars
            y_scroll = ttk.Scrollbar(tree_frame, orient='vertical')
            y_scroll.pack(side='right', fill='y')
            
            x_scroll = ttk.Scrollbar(tree_frame, orient='horizontal')
            x_scroll.pack(side='bottom', fill='x')
            
            # Ordenar colunas
            col_order = ['Nr.', 'SIGLA', 'Cliente', 'CONTAB', 'Iva', 'Subtotal', 'Extras', 
                        'Duodécimos', 'S.Social GER', 'S.Soc Emp', 'Ret. IRS', 'Ret. IRS EXT',
                        'SbTx/Fcomp', 'Outro', 'TOTAL']
            columns = [c for c in col_order if c in all_cols]
            columns += [c for c in all_cols if c not in columns]
            
            tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                               yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
            
            y_scroll.config(command=tree.yview)
            x_scroll.config(command=tree.xview)
            
            # Configurar colunas
            for col in columns:
                tree.heading(col, text=col)
                # Largura baseada no tipo de coluna
                if col == 'Cliente':
                    tree.column(col, width=200, minwidth=100)
                elif col in ['Nr.', 'SIGLA']:
                    tree.column(col, width=60, minwidth=40)
                else:
                    tree.column(col, width=80, minwidth=50)
            
            # Tags para linhas com problemas
            tree.tag_configure('warning', background='#fef3c7')
            tree.tag_configure('normal', background='white')
            
            # Inserir dados
            for idx, item in enumerate(itens):
                values = []
                for col in columns:
                    val = item.get(col, '')
                    if isinstance(val, (int, float)) and col not in ['Nr.']:
                        if val != 0:
                            values.append(f"{val:.2f}€" if col in ['CONTAB', 'Iva', 'Subtotal', 
                                         'Extras', 'Duodécimos', 'S.Social GER', 'S.Soc Emp',
                                         'Ret. IRS', 'Ret. IRS EXT', 'SbTx/Fcomp', 'Outro', 'TOTAL'] else str(val))
                        else:
                            values.append('')
                    else:
                        values.append(str(val) if val else '')
                
                # Aplicar tag de warning se linha tem problemas
                tag = 'warning' if idx in rows_with_issues else 'normal'
                tree.insert('', 'end', values=values, tags=(tag,))
            
            tree.pack(fill='both', expand=True)
            
            # Botões
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill='x', pady=(10, 0))
            
            def generate_and_close():
                preview_window.destroy()
                self._generate()
            
            ttk.Button(btn_frame, text="✅ Gerar PDF(s)", 
                      command=generate_and_close).pack(side='right', padx=5)
            ttk.Button(btn_frame, text="❌ Cancelar", 
                      command=preview_window.destroy).pack(side='right', padx=5)
            
            self.status_var.set("Pré-visualização aberta")
            
        except Exception as e:
            self.status_var.set("❌ Erro na pré-visualização")
            messagebox.showerror("Erro", f"Erro ao carregar pré-visualização:\n\n{str(e)}")
    
    def _setup_profiles_tab(self):
        """Tab de gestão de perfis de configuração."""
        frame = ttk.Frame(self.tab_profiles, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Perfis de Configuração", style='Header.TLabel').pack(pady=(0, 10))
        ttk.Label(frame, text="Guarde diferentes configurações como perfis reutilizáveis.",
                 foreground='gray').pack(pady=(0, 10))

        # Lista de perfis
        list_frame = ttk.LabelFrame(frame, text="Perfis Guardados", padding=10)
        list_frame.pack(fill='both', expand=True, pady=5)

        self.profiles_listbox = tk.Listbox(list_frame, height=8, font=('Helvetica', 10))
        self.profiles_listbox.pack(fill='both', expand=True)

        # Botões
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=10)

        ttk.Button(btn_frame, text="Guardar Perfil Atual", command=self._save_profile).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Carregar Perfil", command=self._load_profile).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Apagar Perfil", command=self._delete_profile).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Atualizar Lista", command=self._refresh_profiles).pack(side='right', padx=5)

        self._refresh_profiles()

    def _refresh_profiles(self):
        """Atualiza a lista de perfis."""
        self.profiles_listbox.delete(0, tk.END)
        for name in list_profiles():
            self.profiles_listbox.insert(tk.END, name)

    def _save_profile(self):
        """Guarda a configuração atual como perfil."""
        popup = tk.Toplevel(self.root)
        popup.title("Guardar Perfil")
        popup.geometry("350x120")
        popup.transient(self.root)
        popup.grab_set()

        f = ttk.Frame(popup, padding=15)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text="Nome do perfil:").pack(anchor='w')
        name_var = tk.StringVar()
        ttk.Entry(f, textvariable=name_var, width=40).pack(fill='x', pady=5)

        def confirm():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Aviso", "Introduza um nome para o perfil.", parent=popup)
                return
            config = self._get_config_from_ui()
            save_profile(name, config)
            popup.destroy()
            self._refresh_profiles()
            messagebox.showinfo("Sucesso", f"Perfil '{name}' guardado!")

        ttk.Button(f, text="Guardar", command=confirm).pack(anchor='e', pady=5)

    def _load_profile(self):
        """Carrega o perfil selecionado."""
        sel = self.profiles_listbox.curselection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um perfil para carregar.")
            return
        name = self.profiles_listbox.get(sel[0])
        config = load_profile(name)
        if config:
            self.config = config
            # Recarregar UI com nova config
            self._reload_config_to_ui()
            messagebox.showinfo("Sucesso", f"Perfil '{name}' carregado!")
        else:
            messagebox.showerror("Erro", f"Não foi possível carregar o perfil '{name}'.")

    def _delete_profile(self):
        """Apaga o perfil selecionado."""
        sel = self.profiles_listbox.curselection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um perfil para apagar.")
            return
        name = self.profiles_listbox.get(sel[0])
        if messagebox.askyesno("Confirmar", f"Apagar o perfil '{name}'?"):
            delete_profile(name)
            self._refresh_profiles()

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
        self.company_nif_var.set(cfg['header']['company_nif'])
        self.logo_path_var.set(cfg['header'].get('logo_path', ''))
        # Table
        self.font_size_var.set(cfg['table']['font_size'])
        self.header_font_size_var.set(cfg['table']['header_font_size'])
        self.row_padding_var.set(cfg['table']['row_padding'])
        self.show_grid_var.set(cfg['table']['show_grid'])
        self.alternate_rows_var.set(cfg['table']['alternate_rows'])
        # Footer
        self.show_signatures_var.set(cfg['footer']['show_signatures'])
        self.show_date_var.set(cfg['footer']['show_date'])
        self.show_observations_var.set(cfg['footer']['show_observations'])
        self.custom_footer_var.set(cfg['footer'].get('custom_footer', ''))
        # Output
        self.auto_open_var.set(cfg['output']['auto_open'])
        self.add_timestamp_var.set(cfg['output']['add_timestamp'])
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

    def _show_irs_summary(self):
        """Mostra resumo de IRS com totais por coluna."""
        excel_path = self.excel_path.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Erro", "Selecione um ficheiro Excel primeiro.")
            return

        try:
            config = self._get_config_from_ui()
            converter = ExcelToPDFConverter(excel_path, None, config)
            data = converter.read_excel_data()
            itens = data.get('itens', [])
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler Excel:\n{e}")
            return

        if not itens:
            messagebox.showwarning("Aviso", "Sem dados no Excel.")
            return

        # Calcular totais
        irs_cols = ['Ret. IRS', 'Ret. IRS EXT']
        summary_cols = ['CONTAB', 'Iva', 'Subtotal', 'Extras', 'Duodécimos',
                        'S.Social GER', 'S.Soc Emp', 'Ret. IRS', 'Ret. IRS EXT',
                        'SbTx/Fcomp', 'Outro', 'TOTAL']

        totals = {}
        for col in summary_cols:
            total = sum(item.get(col, 0) for item in itens if isinstance(item.get(col, 0), (int, float)))
            totals[col] = total

        # Popup
        popup = tk.Toplevel(self.root)
        popup.title("Resumo IRS / Totais")
        popup.geometry("500x450")
        popup.transient(self.root)
        popup.grab_set()

        f = ttk.Frame(popup, padding=15)
        f.pack(fill='both', expand=True)

        mes_ref = data.get('mes_referencia', 'N/A')
        ttk.Label(f, text=f"Resumo — {mes_ref}", font=('Helvetica', 12, 'bold')).pack(anchor='w', pady=(0, 10))
        ttk.Label(f, text=f"Total de clientes: {len(itens)}", foreground='gray').pack(anchor='w')

        # Tabela de totais
        tree_frame = ttk.Frame(f)
        tree_frame.pack(fill='both', expand=True, pady=10)

        cols = ('coluna', 'total')
        tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=12)
        tree.heading('coluna', text='Coluna')
        tree.heading('total', text='Total')
        tree.column('coluna', width=250)
        tree.column('total', width=150, anchor='e')

        tree.tag_configure('irs', background='#fef3c7', foreground='#92400e')
        tree.tag_configure('total_row', background='#e2e8f0', font=('Helvetica', 10, 'bold'))

        for col in summary_cols:
            val = totals[col]
            val_str = f"{val:,.2f}€" if val != 0 else "-"
            tag = 'total_row' if col == 'TOTAL' else ('irs' if col in irs_cols else '')
            tree.insert('', 'end', values=(col, val_str), tags=(tag,) if tag else ())

        tree.pack(fill='both', expand=True)

        # IRS total destacado
        irs_total = totals.get('Ret. IRS', 0) + totals.get('Ret. IRS EXT', 0)
        ttk.Label(f, text=f"Total IRS (Ret. IRS + Ret. IRS EXT): {irs_total:,.2f}€",
                 font=('Helvetica', 11, 'bold'), foreground='#92400e').pack(anchor='w', pady=(5, 0))

        ttk.Button(f, text="Fechar", command=popup.destroy).pack(anchor='e', pady=10)

    def _setup_history_tab(self):
        """Tab de histórico de conversões."""
        frame = ttk.Frame(self.tab_history, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Histórico de Conversões", style='Header.TLabel').pack(pady=(0, 10))

        # Treeview
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill='both', expand=True)

        y_scroll = ttk.Scrollbar(tree_frame, orient='vertical')
        y_scroll.pack(side='right', fill='y')

        columns = ('data', 'ficheiro', 'modo', 'clientes', 'resultado')
        self.history_tree = ttk.Treeview(tree_frame, columns=columns, show='headings',
                                         yscrollcommand=y_scroll.set)
        y_scroll.config(command=self.history_tree.yview)

        self.history_tree.heading('data', text='Data/Hora')
        self.history_tree.heading('ficheiro', text='Ficheiro')
        self.history_tree.heading('modo', text='Modo')
        self.history_tree.heading('clientes', text='Clientes')
        self.history_tree.heading('resultado', text='Resultado')

        self.history_tree.column('data', width=140, minwidth=120)
        self.history_tree.column('ficheiro', width=250, minwidth=150)
        self.history_tree.column('modo', width=100, minwidth=80)
        self.history_tree.column('clientes', width=70, minwidth=50)
        self.history_tree.column('resultado', width=80, minwidth=60)

        self.history_tree.tag_configure('success', foreground='#38A169')
        self.history_tree.tag_configure('error', foreground='#E53E3E')

        self.history_tree.pack(fill='both', expand=True)

        # Botões
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=(10, 0))

        ttk.Button(btn_frame, text="Atualizar", command=self._refresh_history).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Limpar Histórico", command=self._clear_history).pack(side='left', padx=5)

        self._refresh_history()

    def _refresh_history(self):
        """Atualiza a lista de histórico."""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        entries = history.get_history(limit=100)
        for entry in entries:
            try:
                dt = entry['timestamp'][:16].replace('T', ' ')
            except (KeyError, TypeError):
                dt = '?'

            tag = 'success' if entry.get('success', False) else 'error'
            mode_label = 'Individual' if entry.get('mode') == 'individual' else 'Agregado'
            result_label = 'OK' if entry.get('success', False) else 'Erro'

            self.history_tree.insert('', 'end', values=(
                dt,
                entry.get('source_file', '?'),
                mode_label,
                entry.get('clients_count', 0),
                result_label,
            ), tags=(tag,))

    def _clear_history(self):
        """Limpa o histórico de conversões."""
        if messagebox.askyesno("Confirmar", "Tem a certeza que deseja limpar todo o histórico?"):
            history.clear_history()
            self._refresh_history()

    def _export_excel(self):
        """Exporta os dados para Excel formatado."""
        excel_path = self.excel_path.get()

        if not excel_path:
            messagebox.showerror("Erro", "Por favor, selecione um ficheiro Excel.")
            return

        if not os.path.exists(excel_path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado: {excel_path}")
            return

        # Escolher destino
        initial_dir = self.config.get('recent', {}).get('last_output_dir', os.path.dirname(excel_path))
        base_name = os.path.splitext(os.path.basename(excel_path))[0]

        output_path = filedialog.asksaveasfilename(
            title="Guardar Excel formatado como",
            initialdir=initial_dir,
            initialfile=f"{base_name}_formatado.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not output_path:
            return

        try:
            self.status_var.set("A exportar Excel...")
            self.root.update()

            config = self._get_config_from_ui()
            converter = ExcelToPDFConverter(excel_path, None, config)
            data = converter.read_excel_data()

            result_path = export_to_excel(data, output_path, config)
            clients_count = len(data.get('itens', []))

            self.status_var.set(f"Excel exportado: {os.path.basename(result_path)} ({clients_count} clientes)")

            # Registar no histórico
            history.add_entry(excel_path, result_path, 'excel_export', clients_count, True)

            # Guardar última pasta
            self.config.setdefault('recent', {})['last_output_dir'] = os.path.dirname(output_path)
            save_config(self.config)

            messagebox.showinfo("Sucesso",
                f"Excel formatado gerado com sucesso!\n\n"
                f"{result_path}\n\n"
                f"Clientes: {clients_count}")

            # Abrir ficheiro
            if config['output'].get('auto_open', True):
                if sys.platform == 'linux':
                    subprocess.Popen(['xdg-open', result_path])
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', result_path])
                else:
                    os.startfile(result_path)

        except Exception as e:
            self.status_var.set("Erro na exportação Excel")
            history.add_entry(excel_path, output_path, 'excel_export', 0, False, str(e))
            messagebox.showerror("Erro", f"Erro durante a exportação:\n\n{str(e)}")

    def run(self):
        """Inicia a aplicação."""
        self.root.mainloop()


# ============================================
# PONTO DE ENTRADA
