# -*- coding: utf-8 -*-
"""Sub-tab Automação — Watch Folder, Agendamentos e Hooks."""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class AutomationTabMixin:
    """Métodos da sub-tab Automação (Watch Folder, Agendamentos, Hooks)."""

    def _setup_automation_tab(self):
        """Tab de automação: Watch Folder, Agendamentos e Hooks."""
        nb = ttk.Notebook(self.tab_automation)
        nb.pack(fill='both', expand=True, padx=4, pady=4)

        # Sub-sub-tab: Watch Folder
        self.tab_watch = ttk.Frame(nb)
        nb.add(self.tab_watch, text='Watch Folder')
        self._setup_watch_tab()

        # Sub-sub-tab: Agendamentos
        self.tab_schedules = ttk.Frame(nb)
        nb.add(self.tab_schedules, text='Agendamentos')
        self._setup_schedules_tab()

        # Sub-sub-tab: Hooks
        self.tab_hooks = ttk.Frame(nb)
        nb.add(self.tab_hooks, text='Hooks')
        self._setup_hooks_tab()

    def _setup_watch_tab(self):
        """Configuração do Watch Folder."""
        frame = ttk.Frame(self.tab_watch, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        # Activar
        self.watch_enabled_var = tk.BooleanVar(
            value=self.config.get('automation', {}).get('watch_enabled', False))
        ttk.Checkbutton(frame, text="Activar monitorização automática de pasta",
                        variable=self.watch_enabled_var).pack(anchor='w', pady=(0, 6))

        # Pasta
        folder_frame = ttk.LabelFrame(frame, text="Pasta a monitorizar", padding=self._PAD_INNER)
        folder_frame.pack(fill='x', pady=self._PAD_SECTION)
        folder_frame.columnconfigure(0, weight=1)

        self.watch_folder_var = tk.StringVar(
            value=self.config.get('automation', {}).get('watch_folder', ''))
        ttk.Entry(folder_frame, textvariable=self.watch_folder_var).grid(
            row=0, column=0, sticky='ew', pady=4)
        ttk.Button(folder_frame, text="Procurar...",
                   command=self._browse_watch_folder).grid(row=0, column=1, padx=(6, 0), pady=4)

        # Opções
        opts_frame = ttk.LabelFrame(frame, text="Opções", padding=self._PAD_INNER)
        opts_frame.pack(fill='x', pady=self._PAD_SECTION)

        row = ttk.Frame(opts_frame)
        row.pack(fill='x', pady=2)
        ttk.Label(row, text="Modo:").pack(side='left')
        self.watch_mode_var = tk.StringVar(
            value=self.config.get('automation', {}).get('watch_mode', 'individual'))
        ttk.Combobox(row, textvariable=self.watch_mode_var,
                     values=['individual', 'aggregate'], width=14,
                     state='readonly').pack(side='left', padx=(8, 0))

        row2 = ttk.Frame(opts_frame)
        row2.pack(fill='x', pady=2)
        ttk.Label(row2, text="Intervalo (segundos):").pack(side='left')
        self.watch_interval_var = tk.IntVar(
            value=self.config.get('automation', {}).get('watch_interval', 5))
        ttk.Spinbox(row2, textvariable=self.watch_interval_var,
                    from_=1, to=300, width=6).pack(side='left', padx=(8, 0))

        # Botões de controlo
        ctrl_frame = ttk.Frame(frame)
        ctrl_frame.pack(fill='x', pady=(12, 0))

        self.watch_start_btn = ttk.Button(ctrl_frame, text="Iniciar",
                                          command=self._start_watch)
        self.watch_start_btn.pack(side='left', padx=(0, 6))
        self.watch_stop_btn = ttk.Button(ctrl_frame, text="Parar",
                                         command=self._stop_watch, state='disabled')
        self.watch_stop_btn.pack(side='left')

        self.watch_status_var = tk.StringVar(value="Inactivo")
        ttk.Label(frame, textvariable=self.watch_status_var,
                  foreground='#888888', style='Status.TLabel').pack(anchor='w', pady=(6, 0))

        self._watcher = None

    def _browse_watch_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta a monitorizar")
        if folder:
            self.watch_folder_var.set(folder)

    def _start_watch(self):
        from src.watch_folder import WatchFolder
        folder = self.watch_folder_var.get()
        if not folder:
            messagebox.showerror("Erro", "Selecione uma pasta.")
            return
        config = self._get_config_from_ui()
        try:
            self._watcher = WatchFolder(
                folder, config,
                on_new_file=lambda p: self.root.after(
                    0, lambda: self.watch_status_var.set(f"Detectado: {os.path.basename(p)}")),
                on_converted=lambda p, outs: self.root.after(
                    0, lambda: self.watch_status_var.set(
                        f"Convertido: {os.path.basename(p)} ({len(outs)} PDF(s))")),
                on_error=lambda p, e: self.root.after(
                    0, lambda: self.watch_status_var.set(f"Erro: {os.path.basename(p)}: {e}")),
                interval=self.watch_interval_var.get(),
            )
            self._watcher.start()
            self.watch_start_btn.configure(state='disabled')
            self.watch_stop_btn.configure(state='normal')
            self.watch_status_var.set(f"A monitorizar: {folder}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _stop_watch(self):
        if self._watcher:
            self._watcher.stop()
            self._watcher = None
        self.watch_start_btn.configure(state='normal')
        self.watch_stop_btn.configure(state='disabled')
        self.watch_status_var.set("Inactivo")

    def _setup_schedules_tab(self):
        """Configuração de agendamentos."""
        frame = ttk.Frame(self.tab_schedules, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Agendamentos de conversão automática:").pack(anchor='w', pady=(0, 4))

        # Treeview de agendamentos
        cols = ('hora', 'dias', 'origem', 'modo', 'ativo')
        self.schedules_tree = ttk.Treeview(frame, columns=cols, show='headings', height=6)
        for col, heading, width in [
            ('hora', 'Hora', 60), ('dias', 'Dias', 160), ('origem', 'Origem', 220),
            ('modo', 'Modo', 80), ('ativo', 'Ativo', 50),
        ]:
            self.schedules_tree.heading(col, text=heading)
            self.schedules_tree.column(col, width=width, minwidth=40)
        self.schedules_tree.pack(fill='both', expand=True, pady=(0, 6))

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill='x')
        ttk.Button(btn_row, text="Adicionar...", command=self._add_schedule).pack(side='left', padx=(0, 4))
        ttk.Button(btn_row, text="Remover", command=self._remove_schedule).pack(side='left')

        # Carregar agendamentos da config
        self._reload_schedules_tree()

    def _reload_schedules_tree(self):
        """Preenche a treeview de agendamentos a partir da config."""
        from src.scheduler import DIAS_SEMANA
        self.schedules_tree.delete(*self.schedules_tree.get_children())
        for entry in self.config.get('automation', {}).get('schedules', []):
            dias_idx = entry.get('dias', list(range(7)))
            dias_str = ', '.join(DIAS_SEMANA[d] for d in dias_idx if 0 <= d <= 6)
            ativo = 'Sim' if entry.get('enabled', True) else 'Não'
            self.schedules_tree.insert('', 'end', values=(
                entry.get('hora', ''),
                dias_str,
                entry.get('source', ''),
                entry.get('mode', 'individual'),
                ativo,
            ))

    def _add_schedule(self):
        """Abre diálogo para adicionar agendamento."""
        from src.scheduler import DIAS_SEMANA, validate_schedule_entry
        dlg = tk.Toplevel(self.root)
        dlg.title("Novo Agendamento")
        dlg.resizable(False, False)
        dlg.grab_set()

        f = ttk.Frame(dlg, padding=14)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text="Hora (HH:MM):").grid(row=0, column=0, sticky='e', pady=4, padx=(0, 8))
        hora_var = tk.StringVar(value='08:00')
        ttk.Entry(f, textvariable=hora_var, width=10).grid(row=0, column=1, sticky='w')

        ttk.Label(f, text="Dias:").grid(row=1, column=0, sticky='ne', pady=4, padx=(0, 8))
        dias_vars = []
        dias_frame = ttk.Frame(f)
        dias_frame.grid(row=1, column=1, sticky='w')
        for i, nome in enumerate(DIAS_SEMANA):
            v = tk.BooleanVar(value=i < 5)
            dias_vars.append(v)
            ttk.Checkbutton(dias_frame, text=nome, variable=v).grid(
                row=i // 4, column=i % 4, sticky='w', padx=4)

        ttk.Label(f, text="Origem:").grid(row=2, column=0, sticky='e', pady=4, padx=(0, 8))
        source_var = tk.StringVar()
        src_row = ttk.Frame(f)
        src_row.grid(row=2, column=1, sticky='ew')
        ttk.Entry(src_row, textvariable=source_var, width=30).pack(side='left')
        def _browse():
            p = filedialog.askdirectory(title="Pasta de origem") or \
                filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
            if p:
                source_var.set(p)
        ttk.Button(src_row, text="...", command=_browse, width=3).pack(side='left', padx=(4, 0))

        ttk.Label(f, text="Modo:").grid(row=3, column=0, sticky='e', pady=4, padx=(0, 8))
        mode_var = tk.StringVar(value='individual')
        ttk.Combobox(f, textvariable=mode_var, values=['individual', 'aggregate'],
                     width=14, state='readonly').grid(row=3, column=1, sticky='w')

        enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="Activo", variable=enabled_var).grid(
            row=4, column=0, columnspan=2, pady=(8, 0))

        def _confirm():
            dias = [i for i, v in enumerate(dias_vars) if v.get()]
            entry = {'hora': hora_var.get(), 'dias': dias,
                     'source': source_var.get(), 'mode': mode_var.get(),
                     'enabled': enabled_var.get()}
            erros = validate_schedule_entry(entry)
            if erros:
                messagebox.showerror("Erro", '\n'.join(erros), parent=dlg)
                return
            schedules = self.config.setdefault('automation', {}).setdefault('schedules', [])
            schedules.append(entry)
            self._reload_schedules_tree()
            dlg.destroy()

        ttk.Button(f, text="Adicionar", command=_confirm).grid(
            row=5, column=0, columnspan=2, pady=(12, 0))

    def _remove_schedule(self):
        sel = self.schedules_tree.selection()
        if not sel:
            return
        idx = self.schedules_tree.index(sel[0])
        schedules = self.config.get('automation', {}).get('schedules', [])
        if 0 <= idx < len(schedules):
            del schedules[idx]
        self._reload_schedules_tree()

    def _setup_hooks_tab(self):
        """Configuração de post-conversion hooks."""
        frame = ttk.Frame(self.tab_hooks, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame,
                  text="Comandos executados após cada conversão.\n"
                       "Variáveis: {source}, {output}, {outputs}, {folder}").pack(
            anchor='w', pady=(0, 6))

        cols = ('name', 'command', 'timeout', 'enabled')
        self.hooks_tree = ttk.Treeview(frame, columns=cols, show='headings', height=5)
        for col, heading, width in [
            ('name', 'Nome', 100), ('command', 'Comando', 280),
            ('timeout', 'Timeout', 60), ('enabled', 'Ativo', 50),
        ]:
            self.hooks_tree.heading(col, text=heading)
            self.hooks_tree.column(col, width=width, minwidth=40)
        self.hooks_tree.pack(fill='both', expand=True, pady=(0, 6))

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill='x')
        ttk.Button(btn_row, text="Adicionar...", command=self._add_hook).pack(side='left', padx=(0, 4))
        ttk.Button(btn_row, text="Remover", command=self._remove_hook).pack(side='left')

        self._reload_hooks_tree()

    def _reload_hooks_tree(self):
        self.hooks_tree.delete(*self.hooks_tree.get_children())
        for h in self.config.get('automation', {}).get('hooks', []):
            self.hooks_tree.insert('', 'end', values=(
                h.get('name', ''),
                h.get('command', ''),
                h.get('timeout', 30),
                'Sim' if h.get('enabled', True) else 'Não',
            ))

    def _add_hook(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Novo Hook")
        dlg.resizable(False, False)
        dlg.grab_set()

        f = ttk.Frame(dlg, padding=14)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text="Nome:").grid(row=0, column=0, sticky='e', pady=4, padx=(0, 8))
        name_var = tk.StringVar()
        ttk.Entry(f, textvariable=name_var, width=30).grid(row=0, column=1, sticky='w')

        ttk.Label(f, text="Comando:").grid(row=1, column=0, sticky='e', pady=4, padx=(0, 8))
        cmd_var = tk.StringVar()
        ttk.Entry(f, textvariable=cmd_var, width=40).grid(row=1, column=1, sticky='ew')

        ttk.Label(f, text="Timeout (s):").grid(row=2, column=0, sticky='e', pady=4, padx=(0, 8))
        timeout_var = tk.IntVar(value=30)
        ttk.Spinbox(f, textvariable=timeout_var, from_=1, to=300, width=6).grid(
            row=2, column=1, sticky='w')

        enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="Activo", variable=enabled_var).grid(
            row=3, column=0, columnspan=2, pady=(8, 0))

        def _confirm():
            cmd = cmd_var.get().strip()
            if not cmd:
                messagebox.showerror("Erro", "O comando não pode estar vazio.", parent=dlg)
                return
            hook = {'name': name_var.get().strip(), 'command': cmd,
                    'timeout': timeout_var.get(), 'enabled': enabled_var.get()}
            self.config.setdefault('automation', {}).setdefault('hooks', []).append(hook)
            self._reload_hooks_tree()
            dlg.destroy()

        ttk.Button(f, text="Adicionar", command=_confirm).grid(
            row=4, column=0, columnspan=2, pady=(12, 0))

    def _remove_hook(self):
        sel = self.hooks_tree.selection()
        if not sel:
            return
        idx = self.hooks_tree.index(sel[0])
        hooks = self.config.get('automation', {}).get('hooks', [])
        if 0 <= idx < len(hooks):
            del hooks[idx]
        self._reload_hooks_tree()

    def _get_automation_from_ui(self) -> dict:
        """Lê a secção de automação da UI."""
        return {
            'watch_folder': self.watch_folder_var.get() if hasattr(self, 'watch_folder_var') else '',
            'watch_enabled': self.watch_enabled_var.get() if hasattr(self, 'watch_enabled_var') else False,
            'watch_mode': self.watch_mode_var.get() if hasattr(self, 'watch_mode_var') else 'individual',
            'watch_interval': self.watch_interval_var.get() if hasattr(self, 'watch_interval_var') else 5,
            'schedules': self.config.get('automation', {}).get('schedules', []),
            'hooks': self.config.get('automation', {}).get('hooks', []),
        }
