#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de interface gráfica do Conversor Excel → PDF.
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

from src.config import load_config, save_config, DEFAULT_CONFIG
from src.converter import ExcelToPDFConverter
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
        
        # Modo de geração
        mode_frame = ttk.LabelFrame(frame, text="Modo de Geração", padding=10)
        mode_frame.pack(fill='x', pady=5)
        
        self.generation_mode_var = tk.StringVar(value='individual')  # Default: por linha
        ttk.Radiobutton(mode_frame, text="Por Linha (um PDF por cliente)", 
                       variable=self.generation_mode_var, value='individual').pack(anchor='w')
        ttk.Radiobutton(mode_frame, text="Agregado (todos num único PDF)", 
                       variable=self.generation_mode_var, value='aggregate').pack(anchor='w')
        
        # Botões
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=30)
        
        preview_btn = ttk.Button(btn_frame, text="👁 Pré-visualizar", 
                                command=self._preview_excel, style='TButton')
        preview_btn.pack(side='left', padx=5)
        
        generate_btn = ttk.Button(btn_frame, text="Gerar PDF(s)", 
                                 command=self._generate, style='TButton')
        generate_btn.pack(side='left', padx=5)
        
        ttk.Button(btn_frame, text="Guardar Configurações", 
                  command=self._save_config).pack(side='left', padx=5)
        
        # Status
        self.status_var = tk.StringVar(value="Pronto para converter")
        ttk.Label(frame, textvariable=self.status_var, foreground='gray').pack(pady=10)
    
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
        """Tab de configurações de dados bancários."""
        frame = ttk.Frame(self.tab_banking, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(frame, text="Dados Bancários", style='Header.TLabel').pack(pady=(0, 15))
        
        # Descrição
        desc_text = "Configure os dados bancários que aparecerão no rodapé do PDF."
        ttk.Label(frame, text=desc_text, foreground='gray').pack(pady=(0, 10))
        
        # Mostrar dados bancários
        banking_cfg = self.config.get('banking', {})
        self.show_banking_var = tk.BooleanVar(value=banking_cfg.get('show_banking', True))
        ttk.Checkbutton(frame, text="Mostrar dados bancários no PDF", 
                       variable=self.show_banking_var).pack(anchor='w', pady=5)
        
        # Dados do banco
        bank_frame = ttk.LabelFrame(frame, text="Informação Bancária", padding=10)
        bank_frame.pack(fill='x', pady=10)
        
        self.banking_title_var = tk.StringVar(value=banking_cfg.get('title', 'Nossos Dados Bancários:'))
        self.bank_name_var = tk.StringVar(value=banking_cfg.get('bank_name', 'ABANCA'))
        self.iban_var = tk.StringVar(value=banking_cfg.get('iban', 'PT50 0170 3782 0304 0053 5672 9'))
        
        ttk.Label(bank_frame, text="Título:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(bank_frame, textvariable=self.banking_title_var, width=40).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        
        ttk.Label(bank_frame, text="Nome do Banco:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(bank_frame, textvariable=self.bank_name_var, width=40).grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        
        ttk.Label(bank_frame, text="IBAN:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(bank_frame, textvariable=self.iban_var, width=40).grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        
        bank_frame.columnconfigure(1, weight=1)
        
        # Botão Guardar
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(pady=20)
    
    def _pick_color(self, key, var):
        """Abre seletor de cor."""
        color = colorchooser.askcolor(initialcolor=var.get())
        if color[1]:
            var.set(color[1])
            if f'{key}_btn' in self.color_vars:
                self.color_vars[f'{key}_btn'].configure(bg=color[1])
    
    def _browse_excel(self):
        """Seleciona ficheiro Excel."""
        path = filedialog.askopenfilename(
            title="Selecionar ficheiro Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.excel_path.set(path)
    
    def _browse_output(self):
        """Seleciona ficheiro de saída."""
        path = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.output_path.set(path)
    
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
            'banking': {
                'show_banking': self.show_banking_var.get() if hasattr(self, 'show_banking_var') else True,
                'title': self.banking_title_var.get() if hasattr(self, 'banking_title_var') else 'Nossos Dados Bancários:',
                'bank_name': self.bank_name_var.get() if hasattr(self, 'bank_name_var') else 'ABANCA',
                'iban': self.iban_var.get() if hasattr(self, 'iban_var') else 'PT50 0170 3782 0304 0053 5672 9',
            }
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
        """Executa a conversão."""
        excel_path = self.excel_path.get()
        
        if not excel_path:
            messagebox.showerror("Erro", "Por favor, selecione um ficheiro Excel.")
            return
        
        if not os.path.exists(excel_path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado: {excel_path}")
            return
        
        try:
            self.status_var.set("A converter...")
            self.root.update()
            
            config = self._get_config_from_ui()
            output_path = self.output_path.get() or None
            
            converter = ExcelToPDFConverter(excel_path, output_path, config)
            result_path = converter.generate_pdf()
            
            self.status_var.set(f"✅ PDF gerado: {os.path.basename(result_path)}")
            
            messagebox.showinfo("Sucesso", f"PDF gerado com sucesso!\n\n{result_path}")
            
            # Abrir PDF
            if config['output'].get('auto_open', True):
                import subprocess
                if sys.platform == 'linux':
                    subprocess.Popen(['xdg-open', result_path])
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', result_path])
                else:
                    os.startfile(result_path)
                    
        except Exception as e:
            self.status_var.set("❌ Erro na conversão")
            messagebox.showerror("Erro", f"Erro durante a conversão:\n\n{str(e)}")
    
    def _convert_individual(self):
        """Gera PDFs individuais para cada cliente."""
        excel_path = self.excel_path.get()
        
        if not excel_path:
            messagebox.showerror("Erro", "Por favor, selecione um ficheiro Excel.")
            return
        
        if not os.path.exists(excel_path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado: {excel_path}")
            return
        
        try:
            self.status_var.set("A gerar PDFs individuais...")
            self.root.update()
            
            config = self._get_config_from_ui()
            
            converter = ExcelToPDFConverter(excel_path, None, config)
            result_files = converter.generate_individual_pdfs()
            
            if result_files:
                folder = os.path.dirname(result_files[0])
                self.status_var.set(f"✅ {len(result_files)} PDFs gerados!")
                
                messagebox.showinfo("Sucesso", 
                    f"Gerados {len(result_files)} PDFs individuais!\n\n"
                    f"Pasta: {folder}")
                
                # Abrir pasta de destino
                if config['output'].get('auto_open', True):
                    import subprocess
                    if sys.platform == 'linux':
                        subprocess.Popen(['xdg-open', folder])
                    elif sys.platform == 'darwin':
                        subprocess.Popen(['open', folder])
                    else:
                        os.startfile(folder)
            else:
                self.status_var.set("❌ Nenhum PDF gerado")
                messagebox.showwarning("Aviso", "Nenhum item encontrado para gerar PDFs.")
                
        except Exception as e:
            self.status_var.set("❌ Erro na conversão")
            messagebox.showerror("Erro", f"Erro durante a geração:\n\n{str(e)}")
    
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
            
            # Campos importantes que devem ter valor
            important_fields = ['Cliente', 'SIGLA', 'TOTAL']
            
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
    
    def run(self):
        """Inicia a aplicação."""
        self.root.mainloop()


# ============================================
# PONTO DE ENTRADA
