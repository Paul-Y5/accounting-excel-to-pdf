# -*- coding: utf-8 -*-
"""Tab Multificheiros — processamento em lote de ficheiros Excel."""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src import history
from src import notifier
from src.batch_processor import find_excel_files, process_batch


class BatchTabMixin:
    """Métodos da tab Multificheiros."""

    def _setup_batch_tab(self):
        """Tab de processamento em lote."""
        frame = ttk.Frame(self.tab_batch, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Processamento Multificheiros", style='Header.TLabel').pack(anchor='w', pady=(0, 10))

        # Seleção de pasta
        folder_frame = ttk.LabelFrame(frame, text="Pasta com ficheiros Excel", padding=self._PAD_INNER)
        folder_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.batch_folder_var = tk.StringVar()
        ttk.Entry(folder_frame, textvariable=self.batch_folder_var).pack(side='left', fill='x', expand=True)
        ttk.Button(folder_frame, text="Procurar...", command=self._browse_batch_folder).pack(side='right', padx=(8, 0))

        # Modo de geração
        mode_frame = ttk.LabelFrame(frame, text="Modo de Geração", padding=self._PAD_INNER)
        mode_frame.pack(fill='x', pady=self._PAD_SECTION)

        self.batch_mode_var = tk.StringVar(value='individual')
        ttk.Radiobutton(mode_frame, text="Por Linha (um PDF por cliente)",
                        variable=self.batch_mode_var, value='individual').pack(anchor='w', pady=1)
        ttk.Radiobutton(mode_frame, text="Agregado (um PDF por ficheiro Excel)",
                        variable=self.batch_mode_var, value='aggregate').pack(anchor='w', pady=1)

        # Lista de ficheiros encontrados
        files_frame = ttk.LabelFrame(frame, text="Ficheiros encontrados", padding=self._PAD_INNER)
        files_frame.pack(fill='both', expand=True, pady=self._PAD_SECTION)

        self.batch_files_var = tk.StringVar(value="Selecione uma pasta para ver os ficheiros.")
        ttk.Label(files_frame, textvariable=self.batch_files_var, foreground='#666666',
                  justify='left', style='Status.TLabel').pack(anchor='w')

        # Barra de progresso e status
        self.batch_progress_var = tk.DoubleVar(value=0)
        self.batch_progress_bar = ttk.Progressbar(frame, variable=self.batch_progress_var,
                                                   maximum=100, mode='determinate')
        self.batch_progress_bar.pack(fill='x', pady=(10, 2))

        self.batch_status_var = tk.StringVar(value="Pronto")
        ttk.Label(frame, textvariable=self.batch_status_var, foreground='#666666',
                  style='Status.TLabel').pack(pady=(0, 4))

        # Botão
        self.batch_run_btn = ttk.Button(frame, text="Processar Todos",
                                        command=self._run_batch, style='Accent.TButton')
        self.batch_run_btn.pack(anchor='e')

    def _browse_batch_folder(self):
        """Seleciona pasta para processamento em lote."""
        folder = filedialog.askdirectory(title="Selecionar pasta com ficheiros Excel")
        if not folder:
            return
        self.batch_folder_var.set(folder)
        try:
            files = find_excel_files(folder)
            if files:
                names = [os.path.basename(f) for f in files]
                self.batch_files_var.set(f"{len(files)} ficheiro(s):\n" + "\n".join(names))
            else:
                self.batch_files_var.set("Nenhum ficheiro Excel encontrado.")
        except Exception as e:
            self.batch_files_var.set(f"Erro: {e}")

    def _run_batch(self):
        """Executa o processamento em lote numa thread."""
        folder = self.batch_folder_var.get()
        if not folder:
            messagebox.showerror("Erro", "Selecione uma pasta.")
            return

        config = self._get_config_from_ui()
        mode = self.batch_mode_var.get()

        self.batch_run_btn.configure(state='disabled')
        self.batch_progress_var.set(0)

        def on_progress(current, total, filename):
            pct = (current / total) * 100 if total else 0
            self.root.after(0, lambda: self.batch_progress_var.set(pct))
            self.root.after(0, lambda: self.batch_status_var.set(
                f"[{current}/{total}] {filename}"))

        def task():
            try:
                results = process_batch(folder, config, mode=mode,
                                        progress_callback=on_progress)

                ok = sum(1 for r in results if r['success'])
                fail = len(results) - ok

                # Registar no histórico
                for r in results:
                    history.add_entry(r['file'], r['output_path'], f'batch_{mode}',
                                      r['clients_count'], r['success'], r['error'])

                self.root.after(0, lambda: self.batch_progress_var.set(100))
                self.root.after(0, lambda: self.batch_status_var.set(
                    f"Concluído: {ok} com sucesso, {fail} com erro(s)"))
                self.root.after(0, lambda o=ok, f=fail: notifier.notify(
                    "Batch concluído",
                    f"{o} ficheiro(s) com sucesso, {f} com erro(s)",
                    self.config,
                ))
                self.root.after(0, lambda: messagebox.showinfo(
                    "Processamento concluído",
                    f"Processados {len(results)} ficheiro(s).\n"
                    f"Com sucesso: {ok}   Com erros: {fail}"))

                if fail > 0:
                    erros = "\n".join(
                        f"{r['filename']}: {r['error']}"
                        for r in results if not r['success']
                    )
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Ficheiros com erro", erros))

            except Exception as e:
                self.root.after(0, lambda: self.batch_status_var.set(f"Erro: {e}"))
                self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
            finally:
                self.root.after(0, lambda: self.batch_run_btn.configure(state='normal'))
                self.root.after(1500, lambda: self.batch_progress_var.set(0))

        threading.Thread(target=task, daemon=True).start()
