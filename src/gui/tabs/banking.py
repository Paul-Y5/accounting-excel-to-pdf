# -*- coding: utf-8 -*-
"""Sub-tab Dados Bancários — gestão de contas bancárias mostradas no PDF."""

import tkinter as tk
from tkinter import ttk, messagebox

from src.iban_validator import validate_iban, format_iban


class BankingTabMixin:
    """Métodos da sub-tab Dados Bancários."""

    def _setup_banking_tab(self):
        """Tab de configurações de dados bancários (múltiplas contas)."""
        frame = ttk.Frame(self.tab_banking, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="A conta predefinida será usada na geração do PDF.",
                  foreground='#666666', style='Status.TLabel').pack(anchor='w', pady=(0, 8))

        # Mostrar dados bancários
        banking_cfg = self.config.get('banking', {})
        self.show_banking_var = tk.BooleanVar(value=banking_cfg.get('show_banking', True))
        ttk.Checkbutton(frame, text="Mostrar dados bancários no PDF",
                       variable=self.show_banking_var).pack(anchor='w', pady=(0, 4))

        # Título bancário
        title_row = ttk.Frame(frame)
        title_row.pack(fill='x', pady=(0, 8))
        ttk.Label(title_row, text="Título:").pack(side='left', padx=(0, 8))
        self.banking_title_var = tk.StringVar(value=banking_cfg.get('title', 'Nossos Dados Bancários:'))
        ttk.Entry(title_row, textvariable=self.banking_title_var, width=40).pack(side='left')

        # Lista de contas
        accounts_frame = ttk.LabelFrame(frame, text="Contas Bancárias", padding=self._PAD_INNER)
        accounts_frame.pack(fill='both', expand=True, pady=self._PAD_SECTION)

        cols = ('banco', 'iban', 'predefinida')
        self.accounts_tree = ttk.Treeview(accounts_frame, columns=cols, show='headings', height=5)
        self.accounts_tree.heading('banco', text='Banco')
        self.accounts_tree.heading('iban', text='IBAN')
        self.accounts_tree.heading('predefinida', text='Predefinida')
        self.accounts_tree.column('banco', width=120)
        self.accounts_tree.column('iban', width=300)
        self.accounts_tree.column('predefinida', width=80)
        self.accounts_tree.pack(fill='both', expand=True)

        accounts = banking_cfg.get('accounts', [])
        for acc in accounts:
            default_mark = 'Sim' if acc.get('default', False) else ''
            self.accounts_tree.insert('', 'end', values=(
                acc.get('bank_name', ''),
                acc.get('iban', ''),
                default_mark,
            ))

        acc_btn_frame = ttk.Frame(accounts_frame)
        acc_btn_frame.pack(fill='x', pady=(8, 0))

        ttk.Button(acc_btn_frame, text="Adicionar", command=self._add_bank_account).pack(side='left', padx=(0, 6))
        ttk.Button(acc_btn_frame, text="Remover", command=self._remove_bank_account).pack(side='left', padx=6)
        ttk.Button(acc_btn_frame, text="Definir como Predefinida", command=self._set_default_account).pack(side='left', padx=6)

        # Botão Guardar
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Button(frame, text="Guardar Configurações", command=self._save_config).pack(anchor='e')

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
        iban_entry = ttk.Entry(f, textvariable=iban_var, width=35)
        iban_entry.grid(row=1, column=1, padx=5, pady=5)
        iban_status = ttk.Label(f, text='', foreground='gray', font=('Helvetica', 8))
        iban_status.grid(row=2, column=1, sticky='w', padx=5)

        def _on_iban_change(*_):
            raw = iban_var.get().strip()
            if not raw:
                iban_status.config(text='', foreground='gray')
                return
            ok, msg = validate_iban(raw)
            if ok:
                iban_status.config(text=f'✓ {format_iban(raw)}', foreground='green')
            else:
                iban_status.config(text=f'✗ {msg}', foreground='red')

        iban_var.trace_add('write', _on_iban_change)

        def confirm():
            bank = bank_var.get().strip()
            iban = iban_var.get().strip()
            if not bank or not iban:
                messagebox.showwarning("Aviso", "Preencha o nome do banco e o IBAN.", parent=popup)
                return
            ok, msg = validate_iban(iban)
            if not ok:
                if not messagebox.askyesno(
                    "IBAN inválido",
                    f"{msg}\n\nGuardar mesmo assim?",
                    parent=popup,
                ):
                    return
            # Guardar sempre formatado
            self.accounts_tree.insert('', 'end', values=(bank, format_iban(iban), ''))
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
