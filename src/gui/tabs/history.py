# -*- coding: utf-8 -*-
"""Tab Histórico — histórico de conversões, filtros, exportação e relatório anual."""

import os
import sys
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src import history


class HistoryTabMixin:
    """Métodos da tab Histórico."""

    def _setup_history_tab(self):
        """Tab de histórico de conversões."""
        frame = ttk.Frame(self.tab_history, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Histórico de Conversões", style='Header.TLabel').pack(anchor='w', pady=(0, 6))

        # --- Barra de filtros ---
        filter_frame = ttk.LabelFrame(frame, text="Filtros", padding=(self._PAD_INNER, 4))
        filter_frame.pack(fill='x', pady=(0, 8))

        # Linha 1: pesquisa e resultado
        row1 = ttk.Frame(filter_frame)
        row1.pack(fill='x', pady=(0, 4))

        ttk.Label(row1, text="Pesquisa:").pack(side='left', padx=(0, 4))
        self.history_search_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.history_search_var, width=24).pack(side='left', padx=(0, 12))

        ttk.Label(row1, text="Resultado:").pack(side='left', padx=(0, 4))
        self.history_result_var = tk.StringVar(value='Todos')
        ttk.Combobox(row1, textvariable=self.history_result_var,
                     values=['Todos', 'Sucesso', 'Erro'], state='readonly', width=10).pack(side='left', padx=(0, 12))

        ttk.Button(row1, text="Filtrar", command=self._refresh_history).pack(side='left', padx=(0, 6))
        ttk.Button(row1, text="Limpar Filtros", command=self._clear_history_filters).pack(side='left')

        # Linha 2: datas
        row2 = ttk.Frame(filter_frame)
        row2.pack(fill='x')

        ttk.Label(row2, text="De (YYYY-MM-DD):").pack(side='left', padx=(0, 4))
        self.history_date_from_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.history_date_from_var, width=12).pack(side='left', padx=(0, 12))

        ttk.Label(row2, text="Até (YYYY-MM-DD):").pack(side='left', padx=(0, 4))
        self.history_date_to_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.history_date_to_var, width=12).pack(side='left')

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

        self.history_tree.tag_configure('success', foreground='#107C10')
        self.history_tree.tag_configure('error', foreground='#D13438')

        self.history_tree.pack(fill='both', expand=True)

        # Botões
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=(10, 0))

        ttk.Button(btn_frame, text="Atualizar", command=self._refresh_history).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text="Limpar Histórico", command=self._clear_history).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Relatório Anual", command=self._generate_annual_report).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Exportar CSV", command=self._export_history_csv).pack(side='right', padx=(6, 0))
        ttk.Button(btn_frame, text="Exportar Excel", command=self._export_history_excel).pack(side='right', padx=6)

        self._refresh_history()

    def _clear_history_filters(self):
        """Limpa todos os filtros de histórico e recarrega."""
        self.history_search_var.set('')
        self.history_result_var.set('Todos')
        self.history_date_from_var.set('')
        self.history_date_to_var.set('')
        self._refresh_history()

    def _refresh_history(self):
        """Atualiza a lista de histórico aplicando os filtros ativos."""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        search = self.history_search_var.get().strip() if hasattr(self, 'history_search_var') else ''
        result_filter = self.history_result_var.get() if hasattr(self, 'history_result_var') else 'Todos'
        date_from = self.history_date_from_var.get().strip() if hasattr(self, 'history_date_from_var') else ''
        date_to = self.history_date_to_var.get().strip() if hasattr(self, 'history_date_to_var') else ''

        success_only = None
        if result_filter == 'Sucesso':
            success_only = True
        elif result_filter == 'Erro':
            success_only = False

        entries = history.get_history_filtered(
            limit=200,
            date_from=date_from or None,
            date_to=date_to or None,
            success_only=success_only,
            search_term=search or None,
        )

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

    def _export_history_csv(self):
        """Exporta o histórico para CSV."""
        output = filedialog.asksaveasfilename(
            title="Exportar histórico como CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="historico_conversoes.csv",
        )
        if not output:
            return
        try:
            history.export_to_csv(output)
            messagebox.showinfo("Sucesso", f"Histórico exportado:\n{output}")
            if sys.platform == 'linux':
                subprocess.Popen(['xdg-open', os.path.dirname(output)])
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', os.path.dirname(output)])
            else:
                os.startfile(os.path.dirname(output))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar:\n{e}")

    def _export_history_excel(self):
        """Exporta o histórico para Excel."""
        output = filedialog.asksaveasfilename(
            title="Exportar histórico como Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="historico_conversoes.xlsx",
        )
        if not output:
            return
        try:
            history.export_to_excel(output)
            messagebox.showinfo("Sucesso", f"Histórico exportado:\n{output}")
            if sys.platform == 'linux':
                subprocess.Popen(['xdg-open', output])
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', output])
            else:
                os.startfile(output)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar:\n{e}")

    def _generate_annual_report(self):
        """Diálogo para gerar o relatório anual de actividade."""
        from src.annual_report import get_available_years, generate_annual_report_pdf, generate_annual_report_excel

        anos = get_available_years()
        ano_atual = datetime.now().year if not anos else anos[0]
        opcoes = [str(a) for a in anos] if anos else [str(ano_atual)]

        dlg = tk.Toplevel(self.root)
        dlg.title("Relatório Anual")
        dlg.resizable(False, False)
        dlg.grab_set()

        pad = {'padx': 10, 'pady': 6}
        ttk.Label(dlg, text="Seleccione o ano:").grid(row=0, column=0, sticky='w', **pad)
        ano_var = tk.StringVar(value=opcoes[0])
        ttk.Combobox(dlg, textvariable=ano_var, values=opcoes,
                     state='readonly', width=10).grid(row=0, column=1, sticky='w', **pad)

        ttk.Label(dlg, text="Formato:").grid(row=1, column=0, sticky='w', **pad)
        fmt_var = tk.StringVar(value='PDF')
        ttk.Radiobutton(dlg, text='PDF', variable=fmt_var, value='PDF').grid(
            row=1, column=1, sticky='w', padx=10)
        ttk.Radiobutton(dlg, text='Excel (.xlsx)', variable=fmt_var, value='Excel').grid(
            row=2, column=1, sticky='w', padx=10)

        def _gerar():
            ano = int(ano_var.get())
            fmt = fmt_var.get()
            ext = '.pdf' if fmt == 'PDF' else '.xlsx'
            output = filedialog.asksaveasfilename(
                title="Guardar relatório como",
                defaultextension=ext,
                initialfile=f"relatorio_anual_{ano}{ext}",
                filetypes=[("PDF", "*.pdf")] if fmt == 'PDF' else [("Excel", "*.xlsx")],
                parent=dlg,
            )
            if not output:
                return
            dlg.destroy()
            try:
                if fmt == 'PDF':
                    path = generate_annual_report_pdf(ano, output, self.config)
                else:
                    path = generate_annual_report_excel(ano, output)
                messagebox.showinfo("Sucesso", f"Relatório gerado:\n{path}")
                if sys.platform == 'linux':
                    subprocess.Popen(['xdg-open', path])
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', path])
                else:
                    os.startfile(path)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{e}")

        btn_row = ttk.Frame(dlg)
        btn_row.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btn_row, text="Gerar", command=_gerar).pack(side='left', padx=6)
        ttk.Button(btn_row, text="Cancelar", command=dlg.destroy).pack(side='left')
