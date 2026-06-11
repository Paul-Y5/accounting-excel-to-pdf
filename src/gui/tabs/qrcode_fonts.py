# -*- coding: utf-8 -*-
"""Sub-tabs QR Code e Fontes — configuração de QR Code e fontes .ttf personalizadas."""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class QrcodeFontsTabMixin:
    """Métodos das sub-tabs QR Code e Fontes."""

    def _setup_qrcode_tab(self):
        """Tab de configurações de QR Code."""
        frame = ttk.Frame(self.tab_qrcode, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Adiciona um QR Code ao final do PDF com o NIF ou IBAN da empresa.",
                  foreground='#666666', style='Status.TLabel').pack(anchor='w', pady=(0, 8))

        qr_cfg = self.config.get('qrcode', {})

        self.qr_enabled_var = tk.BooleanVar(value=qr_cfg.get('enabled', False))
        ttk.Checkbutton(frame, text="Incluir QR Code no PDF",
                        variable=self.qr_enabled_var).pack(anchor='w', pady=2)

        opts_frame = ttk.LabelFrame(frame, text="Opções do QR Code", padding=self._PAD_INNER)
        opts_frame.pack(fill='x', pady=self._PAD_SECTION)

        content_row = ttk.Frame(opts_frame)
        content_row.pack(fill='x', pady=2)
        ttk.Label(content_row, text="Conteúdo:").pack(side='left', padx=(0, 8))
        self.qr_content_var = tk.StringVar(value=qr_cfg.get('content', 'nif'))
        ttk.Combobox(content_row, textvariable=self.qr_content_var, width=12,
                     values=['nif', 'iban'], state='readonly').pack(side='left')

        size_row = ttk.Frame(opts_frame)
        size_row.pack(fill='x', pady=2)
        ttk.Label(size_row, text="Tamanho (mm):").pack(side='left', padx=(0, 8))
        self.qr_size_var = tk.IntVar(value=qr_cfg.get('size_mm', 25))
        ttk.Spinbox(size_row, from_=10, to=80, textvariable=self.qr_size_var, width=5).pack(side='left')

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')

    def _setup_fonts_tab(self):
        """Tab de configurações de fontes personalizadas."""
        frame = ttk.Frame(self.tab_fonts, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Permite usar fontes .ttf personalizadas no PDF.",
                  foreground='#666666', style='Status.TLabel').pack(anchor='w', pady=(0, 8))

        fonts_cfg = self.config.get('fonts', {})

        # Fonte de corpo
        body_row = ttk.Frame(frame)
        body_row.pack(fill='x', pady=4)
        ttk.Label(body_row, text="Fonte de Corpo:").pack(side='left', padx=(0, 8))
        self.body_font_var = tk.StringVar(value=fonts_cfg.get('body_font', 'Helvetica'))
        ttk.Entry(body_row, textvariable=self.body_font_var, width=25).pack(side='left')

        # Fonte de cabeçalho
        header_row = ttk.Frame(frame)
        header_row.pack(fill='x', pady=4)
        ttk.Label(header_row, text="Fonte de Cabeçalho:").pack(side='left', padx=(0, 8))
        self.header_font_var = tk.StringVar(value=fonts_cfg.get('header_font', 'Helvetica-Bold'))
        ttk.Entry(header_row, textvariable=self.header_font_var, width=25).pack(side='left')

        # Fontes registadas
        reg_frame = ttk.LabelFrame(frame, text="Fontes Registadas (.ttf)", padding=self._PAD_INNER)
        reg_frame.pack(fill='both', expand=True, pady=self._PAD_SECTION)

        cols = ('nome', 'caminho')
        self.fonts_tree = ttk.Treeview(reg_frame, columns=cols, show='headings', height=4)
        self.fonts_tree.heading('nome', text='Nome')
        self.fonts_tree.heading('caminho', text='Caminho')
        self.fonts_tree.column('nome', width=120)
        self.fonts_tree.column('caminho', width=350)
        self.fonts_tree.pack(fill='both', expand=True)

        for entry in fonts_cfg.get('registered', []):
            self.fonts_tree.insert('', 'end', values=(entry.get('name', ''), entry.get('path', '')))

        btn_frame = ttk.Frame(reg_frame)
        btn_frame.pack(fill='x', pady=(8, 0))
        ttk.Button(btn_frame, text="Adicionar .ttf", command=self._add_font).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text="Remover", command=self._remove_font).pack(side='left', padx=6)

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')

    def _add_font(self):
        """Adiciona uma fonte .ttf via diálogo de ficheiro."""
        path = filedialog.askopenfilename(
            title="Selecionar fonte .ttf",
            filetypes=[("TrueType Font", "*.ttf"), ("All files", "*.*")]
        )
        if not path:
            return
        name = os.path.splitext(os.path.basename(path))[0]
        self.fonts_tree.insert('', 'end', values=(name, path))

    def _remove_font(self):
        """Remove a fonte selecionada."""
        selected = self.fonts_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma fonte para remover.")
            return
        for item in selected:
            self.fonts_tree.delete(item)

    def _get_fonts_from_ui(self) -> dict:
        """Lê a configuração de fontes da UI."""
        registered = []
        if hasattr(self, 'fonts_tree'):
            for item in self.fonts_tree.get_children():
                vals = self.fonts_tree.item(item, 'values')
                registered.append({'name': vals[0], 'path': vals[1]})
        return {
            'body_font': self.body_font_var.get() if hasattr(self, 'body_font_var') else 'Helvetica',
            'header_font': self.header_font_var.get() if hasattr(self, 'header_font_var') else 'Helvetica-Bold',
            'registered': registered,
        }
