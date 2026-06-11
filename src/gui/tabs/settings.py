# -*- coding: utf-8 -*-
"""Tab Definições — sub-notebook com Página PDF, Cabeçalho, Tabela/Rodapé e Cores."""

import tkinter as tk
from tkinter import ttk, filedialog, colorchooser


class SettingsTabMixin:
    """Métodos da tab Definições e das sub-tabs PDF, Cabeçalho, Tabela e Cores."""

    def _setup_settings_tab(self):
        """Tab de definições com sub-notebook para todas as configurações."""
        settings_nb = ttk.Notebook(self.tab_settings)
        settings_nb.pack(fill='both', expand=True, padx=5, pady=5)

        # Sub-tab: Página PDF
        self.tab_pdf = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_pdf, text='Página PDF')
        self._setup_pdf_tab()

        # Sub-tab: Cabeçalho
        self.tab_header = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_header, text='Cabeçalho')
        self._setup_header_tab()

        # Sub-tab: Tabela e Rodapé
        self.tab_table = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_table, text='Tabela e Rodapé')
        self._setup_table_tab()

        # Sub-tab: Cores
        self.tab_colors = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_colors, text='Cores')
        self._setup_colors_tab()

        # Sub-tab: Contabilidade
        self.tab_contab = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_contab, text='Contabilidade')
        self._setup_contabilidade_tab()

        # Sub-tab: Dados Bancários
        self.tab_banking = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_banking, text='Dados Bancários')
        self._setup_banking_tab()

        # Sub-tab: QR Code
        self.tab_qrcode = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_qrcode, text='QR Code')
        self._setup_qrcode_tab()

        # Sub-tab: Fontes
        self.tab_fonts = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_fonts, text='Fontes')
        self._setup_fonts_tab()

        # Sub-tab: Nº Documentos
        self.tab_doc_seq = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_doc_seq, text='Nº Documentos')
        self._setup_doc_sequence_tab()

        # Sub-tab: Automação
        self.tab_automation = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_automation, text='Automação')
        self._setup_automation_tab()

    def _setup_pdf_tab(self):
        """Tab de configurações do PDF."""
        frame = ttk.Frame(self.tab_pdf, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        # Tamanho da página
        size_frame = ttk.LabelFrame(frame, text="Tamanho da Página", padding=self._PAD_INNER)
        size_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.page_size_var = tk.StringVar(value=self.config['pdf']['page_size'])
        ttk.Label(size_frame, text="Tamanho:").grid(row=0, column=0, sticky='e', padx=(0, 8), pady=4)
        ttk.Combobox(size_frame, textvariable=self.page_size_var,
                    values=['A4', 'A3', 'Letter'], width=15, state='readonly').grid(row=0, column=1, padx=(0, 20), pady=4)

        self.orientation_var = tk.StringVar(value=self.config['pdf']['orientation'])
        ttk.Label(size_frame, text="Orientação:").grid(row=0, column=2, sticky='e', padx=(0, 8), pady=4)
        ttk.Combobox(size_frame, textvariable=self.orientation_var,
                    values=['portrait', 'landscape'], width=15, state='readonly').grid(row=0, column=3, pady=4)

        # Margens
        margin_frame = ttk.LabelFrame(frame, text="Margens (mm)", padding=self._PAD_INNER)
        margin_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.margin_top_var = tk.IntVar(value=self.config['pdf']['margin_top'])
        self.margin_bottom_var = tk.IntVar(value=self.config['pdf']['margin_bottom'])
        self.margin_left_var = tk.IntVar(value=self.config['pdf']['margin_left'])
        self.margin_right_var = tk.IntVar(value=self.config['pdf']['margin_right'])

        for i, (label, var) in enumerate([
            ("Superior:", self.margin_top_var), ("Inferior:", self.margin_bottom_var),
            ("Esquerda:", self.margin_left_var), ("Direita:", self.margin_right_var),
        ]):
            row, col = divmod(i, 2)
            ttk.Label(margin_frame, text=label).grid(row=row, column=col*2, sticky='e', padx=(0, 8), pady=4)
            ttk.Spinbox(margin_frame, textvariable=var, from_=5, to=50, width=8).grid(
                row=row, column=col*2+1, padx=(0, 20), pady=4)

        # Interface
        ui_frame = ttk.LabelFrame(frame, text="Interface", padding=self._PAD_INNER)
        ui_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.notifications_enabled_var = tk.BooleanVar(
            value=self.config.get('ui', {}).get('notifications_enabled', True))
        ttk.Checkbutton(ui_frame, text="Ativar notificações desktop após conversão",
                        variable=self.notifications_enabled_var).pack(anchor='w')

        # Nome do ficheiro de saída
        name_frame = ttk.LabelFrame(frame, text="Nome do Ficheiro de Saída", padding=self._PAD_INNER)
        name_frame.pack(fill='x', pady=self._PAD_SECTION)
        name_frame.columnconfigure(1, weight=1)

        self.filename_template_var = tk.StringVar(
            value=self.config.get('output', {}).get('filename_template', ''))
        ttk.Label(name_frame, text="Template:").grid(row=0, column=0, sticky='e', padx=(0, 8), pady=4)
        ttk.Entry(name_frame, textvariable=self.filename_template_var).grid(
            row=0, column=1, sticky='ew', pady=4)
        ttk.Label(name_frame,
                  text="Tokens: {empresa}  {mes}  {nr}  {data}  {sigla}  {cliente}",
                  foreground='#666666', style='Status.TLabel').grid(
            row=1, column=0, columnspan=2, sticky='w', pady=(0, 2))
        ttk.Label(name_frame,
                  text="Exemplo: {empresa}_{mes}   Deixe em branco para usar o nome do ficheiro Excel.",
                  foreground='#666666', style='Status.TLabel').grid(
            row=2, column=0, columnspan=2, sticky='w')

        # Botão Guardar
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')

    def _setup_header_tab(self):
        """Tab de configurações do cabeçalho."""
        frame = ttk.Frame(self.tab_header, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        # Mostrar cabeçalho
        self.show_header_var = tk.BooleanVar(value=self.config['header']['show_header'])
        ttk.Checkbutton(frame, text="Mostrar cabeçalho no PDF",
                       variable=self.show_header_var).pack(anchor='w', pady=(0, 8))

        # Dados da empresa
        company_frame = ttk.LabelFrame(frame, text="Dados da Empresa", padding=self._PAD_INNER)
        company_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.company_name_var = tk.StringVar(value=self.config['header']['company_name'])
        self.company_address_var = tk.StringVar(value=self.config['header']['company_address'])
        self.company_phone_var = tk.StringVar(value=self.config['header']['company_phone'])
        self.company_email_var = tk.StringVar(value=self.config['header']['company_email'])
        self.company_website_var = tk.StringVar(value=self.config['header'].get('company_website', ''))
        self.company_nif_var = tk.StringVar(value=self.config['header']['company_nif'])

        fields = [
            ("Nome:", self.company_name_var),
            ("Morada:", self.company_address_var),
            ("Telefone:", self.company_phone_var),
            ("Email:", self.company_email_var),
            ("Website:", self.company_website_var),
            ("NIF:", self.company_nif_var),
        ]

        for i, (label, var) in enumerate(fields):
            ttk.Label(company_frame, text=label).grid(row=i, column=0, sticky='e', padx=(0, 8), pady=4)
            ttk.Entry(company_frame, textvariable=var).grid(row=i, column=1, sticky='ew', pady=4)

        company_frame.columnconfigure(1, weight=1)

        # Logo
        logo_frame = ttk.LabelFrame(frame, text="Logo (opcional)", padding=self._PAD_INNER)
        logo_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.logo_path_var = tk.StringVar(value=self.config['header'].get('logo_path', ''))
        ttk.Entry(logo_frame, textvariable=self.logo_path_var).pack(side='left', fill='x', expand=True)
        ttk.Button(logo_frame, text="Procurar...", command=self._browse_logo).pack(side='right', padx=(8, 0))

        # Botão Guardar
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')

    def _setup_table_tab(self):
        """Tab de configurações da tabela e rodapé."""
        frame = ttk.Frame(self.tab_table, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        # Fontes
        font_frame = ttk.LabelFrame(frame, text="Fontes e Espaçamento", padding=self._PAD_INNER)
        font_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.font_size_var = tk.IntVar(value=self.config['table']['font_size'])
        self.header_font_size_var = tk.IntVar(value=self.config['table']['header_font_size'])
        self.row_padding_var = tk.IntVar(value=self.config['table']['row_padding'])

        for i, (label, var, rng) in enumerate([
            ("Texto:", self.font_size_var, (6, 14)),
            ("Cabeçalho:", self.header_font_size_var, (8, 16)),
            ("Espaçamento:", self.row_padding_var, (2, 12)),
        ]):
            ttk.Label(font_frame, text=label).grid(row=0, column=i*2, sticky='e', padx=(0, 8), pady=4)
            ttk.Spinbox(font_frame, textvariable=var, from_=rng[0], to=rng[1], width=6).grid(
                row=0, column=i*2+1, padx=(0, 16), pady=4)

        # Opções da tabela
        options_frame = ttk.LabelFrame(frame, text="Opções da Tabela", padding=self._PAD_INNER)
        options_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.show_grid_var = tk.BooleanVar(value=self.config['table']['show_grid'])
        self.alternate_rows_var = tk.BooleanVar(value=self.config['table']['alternate_rows'])

        ttk.Checkbutton(options_frame, text="Mostrar grelha/bordas",
                       variable=self.show_grid_var).pack(anchor='w', pady=2)
        ttk.Checkbutton(options_frame, text="Cores alternadas nas linhas",
                       variable=self.alternate_rows_var).pack(anchor='w', pady=2)

        # Rodapé
        footer_frame = ttk.LabelFrame(frame, text="Rodapé", padding=self._PAD_INNER)
        footer_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.show_signatures_var = tk.BooleanVar(value=self.config['footer']['show_signatures'])
        self.show_date_var = tk.BooleanVar(value=self.config['footer']['show_date'])
        self.show_observations_var = tk.BooleanVar(value=self.config['footer']['show_observations'])

        self.show_iva_summary_var = tk.BooleanVar(
            value=self.config.get('pdf', {}).get('show_iva_summary', True))
        ttk.Checkbutton(footer_frame, text="Mostrar resumo de IVA",
                       variable=self.show_iva_summary_var).pack(anchor='w', pady=2)

        ttk.Checkbutton(footer_frame, text="Mostrar área de assinaturas",
                       variable=self.show_signatures_var).pack(anchor='w', pady=2)
        ttk.Checkbutton(footer_frame, text="Mostrar data de geração",
                       variable=self.show_date_var).pack(anchor='w', pady=2)
        ttk.Checkbutton(footer_frame, text="Mostrar observações",
                       variable=self.show_observations_var).pack(anchor='w', pady=2)

        footer_text_frame = ttk.Frame(footer_frame)
        footer_text_frame.pack(fill='x', pady=(8, 0))
        ttk.Label(footer_text_frame, text="Texto personalizado:").pack(side='left', padx=(0, 8))
        self.custom_footer_var = tk.StringVar(value=self.config['footer'].get('custom_footer', ''))
        ttk.Entry(footer_text_frame, textvariable=self.custom_footer_var).pack(side='left', fill='x', expand=True)

        # Botão Guardar
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')

    def _setup_colors_tab(self):
        """Tab de configurações de cores."""
        frame = ttk.Frame(self.tab_colors, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        colors_frame = ttk.LabelFrame(frame, text="Cores do PDF", padding=self._PAD_INNER)
        colors_frame.pack(fill='x', pady=self._PAD_SECTION)
        colors_frame.columnconfigure(1, weight=1)

        self.color_vars = {}

        colors_config = [
            ('header_bg', 'Fundo do cabeçalho'),
            ('header_text', 'Texto do cabeçalho'),
            ('row_alt', 'Linhas alternadas'),
            ('border', 'Bordas'),
            ('title', 'Título da empresa'),
        ]

        for i, (key, label) in enumerate(colors_config):
            color_value = self.config['colors'].get(key, '#000000')
            var = tk.StringVar(value=color_value)
            self.color_vars[key] = var

            ttk.Label(colors_frame, text=label).grid(row=i, column=0, sticky='e', padx=(0, 12), pady=6)
            ttk.Entry(colors_frame, textvariable=var, width=10).grid(row=i, column=1, sticky='w', pady=6)

            color_btn = tk.Button(colors_frame, text="     ", bg=color_value, width=4,
                                 relief='flat', borderwidth=1,
                                 command=lambda k=key, v=var: self._pick_color(k, v))
            color_btn.grid(row=i, column=2, padx=(8, 0), pady=6)
            self.color_vars[f'{key}_btn'] = color_btn

        # Botão Guardar
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')

    def _pick_color(self, key, var):
        """Abre seletor de cor."""
        color = colorchooser.askcolor(initialcolor=var.get())
        if color[1]:
            var.set(color[1])
            if f'{key}_btn' in self.color_vars:
                self.color_vars[f'{key}_btn'].configure(bg=color[1])

    def _browse_logo(self):
        """Seleciona ficheiro de logo."""
        path = filedialog.askopenfilename(
            title="Selecionar logo",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif"), ("All files", "*.*")]
        )
        if path:
            self.logo_path_var.set(path)
