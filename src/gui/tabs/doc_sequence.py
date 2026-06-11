# -*- coding: utf-8 -*-
"""Sub-tab Nº Documentos — gestão de sequências de numeração de documentos."""

import tkinter as tk
from tkinter import ttk, messagebox

from src.doc_sequence import list_series, upsert_serie, reset_serie, delete_serie


class DocSequenceTabMixin:
    """Métodos da sub-tab Nº Documentos."""

    def _setup_doc_sequence_tab(self):
        """Tab de gestão de sequências de números de documento."""
        frame = ttk.Frame(self.tab_doc_seq, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text='Sequências de Nº de Documento',
                  style='Header.TLabel').pack(anchor='w')
        ttk.Label(
            frame,
            text='Gerir séries de numeração automática para faturas, recibos, etc.',
            foreground='gray',
        ).pack(anchor='w', pady=(0, 10))

        # --- Treeview ---
        cols = ('serie', 'ano', 'ultimo', 'proximo', 'reset')
        self._doc_seq_tree = ttk.Treeview(
            frame, columns=cols, show='headings', height=8, selectmode='browse'
        )
        self._doc_seq_tree.heading('serie',   text='Série')
        self._doc_seq_tree.heading('ano',     text='Ano')
        self._doc_seq_tree.heading('ultimo',  text='Último Nº')
        self._doc_seq_tree.heading('proximo', text='Próximo')
        self._doc_seq_tree.heading('reset',   text='Reset Anual')
        self._doc_seq_tree.column('serie',   width=80,  anchor='center')
        self._doc_seq_tree.column('ano',     width=60,  anchor='center')
        self._doc_seq_tree.column('ultimo',  width=90,  anchor='center')
        self._doc_seq_tree.column('proximo', width=130, anchor='center')
        self._doc_seq_tree.column('reset',   width=90,  anchor='center')
        self._doc_seq_tree.pack(fill='x', pady=(0, 6))

        self._reload_doc_seq_tree()

        # --- Botões ---
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(anchor='w', pady=(0, 12))
        ttk.Button(btn_frame, text='Nova Série',    command=self._add_doc_serie).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text='Editar',        command=self._edit_doc_serie).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text='Reiniciar',     command=self._reset_doc_serie).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text='Remover',       command=self._remove_doc_serie).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text='Actualizar',    command=self._reload_doc_seq_tree).pack(side='left')

        # --- Séries predefinidas ---
        sep = ttk.LabelFrame(frame, text='Séries Predefinidas', padding=8)
        sep.pack(fill='x')
        ttk.Label(sep, text='Criação rápida de séries comuns:').pack(anchor='w', pady=(0, 6))
        presets_frame = ttk.Frame(sep)
        presets_frame.pack(anchor='w')
        for code, label in [('FT', 'Fatura'), ('FR', 'Fatura-Recibo'),
                             ('REC', 'Recibo'), ('ND', 'Nota de Entrega'),
                             ('NC', 'Nota de Crédito')]:
            ttk.Button(
                presets_frame, text=f'{code} — {label}',
                command=lambda c=code: self._create_preset_serie(c),
            ).pack(side='left', padx=(0, 6))

    def _reload_doc_seq_tree(self):
        """Recarrega a treeview das séries."""
        for item in self._doc_seq_tree.get_children():
            self._doc_seq_tree.delete(item)
        for s in list_series():
            self._doc_seq_tree.insert('', 'end', values=(
                s['serie'],
                s['ano'],
                s['ultimo_numero'],
                s['proximo'],
                'Sim' if s['reset_anual'] else 'Não',
            ))

    def _add_doc_serie(self):
        """Diálogo para criar uma nova série."""
        self._doc_serie_dialog()

    def _edit_doc_serie(self):
        """Diálogo para editar a série seleccionada."""
        sel = self._doc_seq_tree.selection()
        if not sel:
            messagebox.showwarning('Aviso', 'Seleccione uma série para editar.')
            return
        vals = self._doc_seq_tree.item(sel[0], 'values')
        self._doc_serie_dialog(
            serie=vals[0],
            ano=int(vals[1]),
            ultimo=int(vals[2]),
            reset_anual=(vals[4] == 'Sim'),
        )

    def _doc_serie_dialog(self, serie='', ano=None, ultimo=0, reset_anual=True):
        """Diálogo genérico para criar/editar série."""
        from datetime import datetime
        if ano is None:
            ano = datetime.now().year

        dlg = tk.Toplevel(self.root)
        dlg.title('Série de Documento')
        dlg.resizable(False, False)
        dlg.grab_set()

        pad = {'padx': 8, 'pady': 4}
        ttk.Label(dlg, text='Série (ex: FT, FR, REC):').grid(row=0, column=0, sticky='w', **pad)
        v_serie = tk.StringVar(value=serie)
        e_serie = ttk.Entry(dlg, textvariable=v_serie, width=10)
        e_serie.grid(row=0, column=1, sticky='w', **pad)
        if serie:
            e_serie.config(state='disabled')

        ttk.Label(dlg, text='Ano:').grid(row=1, column=0, sticky='w', **pad)
        v_ano = tk.IntVar(value=ano)
        ttk.Spinbox(dlg, textvariable=v_ano, from_=2000, to=2100, width=8).grid(row=1, column=1, sticky='w', **pad)

        ttk.Label(dlg, text='Último Nº emitido:').grid(row=2, column=0, sticky='w', **pad)
        v_ultimo = tk.IntVar(value=ultimo)
        ttk.Spinbox(dlg, textvariable=v_ultimo, from_=0, to=99999, width=8).grid(row=2, column=1, sticky='w', **pad)

        v_reset = tk.BooleanVar(value=reset_anual)
        ttk.Checkbutton(dlg, text='Reset anual (reinicia em Jan)', variable=v_reset).grid(
            row=3, column=0, columnspan=2, sticky='w', **pad)

        def _confirm():
            s = v_serie.get().strip().upper()
            if not s:
                messagebox.showwarning('Aviso', 'A série não pode estar vazia.', parent=dlg)
                return
            upsert_serie(s, v_ultimo.get(), v_ano.get(), v_reset.get())
            self._reload_doc_seq_tree()
            dlg.destroy()

        btn_row = ttk.Frame(dlg)
        btn_row.grid(row=4, column=0, columnspan=2, pady=8)
        ttk.Button(btn_row, text='Guardar', command=_confirm).pack(side='left', padx=6)
        ttk.Button(btn_row, text='Cancelar', command=dlg.destroy).pack(side='left')

    def _reset_doc_serie(self):
        """Reinicia o contador da série seleccionada."""
        sel = self._doc_seq_tree.selection()
        if not sel:
            messagebox.showwarning('Aviso', 'Seleccione uma série para reiniciar.')
            return
        serie = self._doc_seq_tree.item(sel[0], 'values')[0]
        if messagebox.askyesno('Confirmar',
                               f"Reiniciar o contador da série '{serie}' para zero?"):
            reset_serie(serie)
            self._reload_doc_seq_tree()

    def _remove_doc_serie(self):
        """Remove a série seleccionada."""
        sel = self._doc_seq_tree.selection()
        if not sel:
            messagebox.showwarning('Aviso', 'Seleccione uma série para remover.')
            return
        serie = self._doc_seq_tree.item(sel[0], 'values')[0]
        if messagebox.askyesno('Confirmar',
                               f"Remover permanentemente a série '{serie}'?"):
            delete_serie(serie)
            self._reload_doc_seq_tree()

    def _create_preset_serie(self, code: str):
        """Cria uma série predefinida se ainda não existir."""
        existing = {s['serie'] for s in list_series()}
        if code in existing:
            messagebox.showinfo('Info', f"A série '{code}' já existe.")
            return
        upsert_serie(code)
        self._reload_doc_seq_tree()
