# -*- coding: utf-8 -*-
"""Sub-tab Contabilidade — colunas, formatação e larguras do mapa de contabilidade."""

import tkinter as tk
from tkinter import ttk


class ContabilidadeTabMixin:
    """Métodos da sub-tab Contabilidade."""

    def _setup_contabilidade_tab(self):
        """Tab de configurações de contabilidade."""
        frame = ttk.Frame(self.tab_contab, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Separe as colunas por vírgula, na ordem desejada.",
                  foreground='#666666', style='Status.TLabel').pack(anchor='w', pady=(0, 8))

        # Colunas
        colunas_frame = ttk.LabelFrame(frame, text="Colunas a Incluir", padding=self._PAD_INNER)
        colunas_frame.pack(fill='x', pady=self._PAD_SECTION)

        contab_cfg = self.config.get('contabilidade', {})
        default_colunas = 'Nr., SIGLA, Cliente, CONTAB, Iva, Subtotal, Extras, Duodécimos, S.Social GER, S.Soc Emp, Ret. IRS, Ret. IRS EXT, SbTx/Fcomp, Outro, TOTAL'

        self.contab_colunas_var = tk.StringVar(value=contab_cfg.get('colunas', default_colunas))

        self.contab_colunas_text = tk.Text(colunas_frame, height=3, wrap='word',
                                           font=(self._FONT_FAMILY, self._FONT_SIZE))
        self.contab_colunas_text.pack(fill='x', pady=(0, 8))
        self.contab_colunas_text.insert('1.0', self.contab_colunas_var.get())

        def reset_colunas():
            self.contab_colunas_text.delete('1.0', tk.END)
            self.contab_colunas_text.insert('1.0', default_colunas)

        ttk.Button(colunas_frame, text="Restaurar Padrão", command=reset_colunas).pack(anchor='e')

        # Opções de destaque
        options_frame = ttk.LabelFrame(frame, text="Formatação", padding=self._PAD_INNER)
        options_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.contab_destacar_total_var = tk.BooleanVar(value=contab_cfg.get('destacar_total', True))
        ttk.Checkbutton(options_frame, text="Destacar coluna TOTAL com cor de fundo",
                       variable=self.contab_destacar_total_var).pack(anchor='w', pady=2)

        self.contab_destacar_valores_var = tk.BooleanVar(value=contab_cfg.get('destacar_valores', True))
        ttk.Checkbutton(options_frame, text="Destacar valores (positivos/negativos)",
                       variable=self.contab_destacar_valores_var).pack(anchor='w', pady=2)

        # Larguras de colunas configuráveis
        widths_frame = ttk.LabelFrame(frame, text="Larguras das Colunas (mm, 0 = automático)", padding=self._PAD_INNER)
        widths_frame.pack(fill='x', pady=self._PAD_SECTION)

        col_widths_cfg = contab_cfg.get('col_widths', {})
        all_cols = [
            'Nr.', 'SIGLA', 'Cliente', 'CONTAB', 'Iva', 'Subtotal',
            'Extras', 'Duodécimos', 'S.Social GER', 'S.Soc Emp',
            'Ret. IRS', 'Ret. IRS EXT', 'SbTx/Fcomp', 'Outro', 'TOTAL',
        ]
        self.contab_col_widths_vars = {}
        grid = ttk.Frame(widths_frame)
        grid.pack(fill='x')
        for i, col in enumerate(all_cols):
            row, gcol = divmod(i, 3)
            val = str(col_widths_cfg.get(col, 0))
            var = tk.StringVar(value=val)
            self.contab_col_widths_vars[col] = var
            cell = ttk.Frame(grid)
            cell.grid(row=row, column=gcol, sticky='w', padx=(0, 12), pady=2)
            ttk.Label(cell, text=f"{col}:", width=14, anchor='w').pack(side='left')
            ttk.Spinbox(cell, from_=0, to=200, textvariable=var, width=5).pack(side='left')

        # Referência de colunas (colapsável via expander)
        ref_frame = ttk.LabelFrame(frame, text="Referência de Colunas", padding=self._PAD_INNER)
        ref_frame.pack(fill='x', pady=self._PAD_SECTION)

        ref_cols = [
            ("Nr.", "Número"),        ("SIGLA", "Sigla"),
            ("Cliente", "Nome"),      ("CONTAB", "Contabilidade"),
            ("Iva", "IVA"),           ("Subtotal", "Subtotal"),
            ("Extras", "Extras"),     ("Duodécimos", "Duodécimos"),
            ("S.Social GER", "SS Gerente"), ("S.Soc Emp", "SS Empresa"),
            ("Ret. IRS", "IRS"),      ("Ret. IRS EXT", "IRS Ext."),
            ("SbTx/Fcomp", "Sub/Férias"), ("Outro", "Outros"),
            ("TOTAL", "Total"),
        ]

        ref_grid = ttk.Frame(ref_frame)
        ref_grid.pack(fill='x')
        for i, (code, desc) in enumerate(ref_cols):
            row, col = divmod(i, 3)
            ttk.Label(ref_grid, text=f"{code} — {desc}", foreground='#666666',
                      style='Status.TLabel').grid(row=row, column=col, sticky='w', padx=(0, 20), pady=1)

        # Botão Guardar
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')
