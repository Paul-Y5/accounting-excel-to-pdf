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

from src.config import load_config, save_config, export_config, import_config, DEFAULT_CONFIG, list_profiles, save_profile, load_profile, delete_profile
from src.converter import ExcelToPDFConverter
from src.nif_validator import validate_nif
from src.excel_exporter import export_to_excel
from src import history
from src.database import init_db, migrate_from_json, update_client_cache, get_cached_clients
from src.email_sender import open_email_client
from src.batch_processor import find_excel_files, process_batch
from src import notifier
class ConverterApp:
    """Aplicação principal com interface gráfica para conversão de Excel para PDF."""

    # Constantes de UI
    _PAD_OUTER = 12          # padding exterior das tabs
    _PAD_SECTION = (0, 8)    # espaço vertical entre secções
    _PAD_INNER = 10          # padding interior dos LabelFrames
    _FONT_FAMILY = 'Helvetica'
    _FONT_SIZE = 10
    _FONT_HEADER = 14

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor Excel")
        self.root.geometry("780x640")
        self.root.minsize(700, 560)
        self.root.resizable(True, True)

        # Tema moderno (Sun Valley) — aplicado antes de criar widgets
        self._sv_ttk_available = False
        try:
            import sv_ttk
            self._sv_ttk_available = True
        except ImportError:
            pass

        # Inicializar base de dados SQLite
        init_db()
        migrate_from_json()

        # Carregar configurações
        self.config = load_config()

        # Aplicar tema guardado
        if self._sv_ttk_available:
            import sv_ttk
            sv_ttk.set_theme(self.config.get('ui', {}).get('theme', 'light'))

        # Últimos PDFs gerados (para envio por email)
        self._last_generated_files = []

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
        # Fonte global
        default_font = (self._FONT_FAMILY, self._FONT_SIZE)
        self.root.option_add('*Font', default_font)

        # Estilos ttk
        style = ttk.Style()
        style.configure('TButton', padding=(12, 6))
        style.configure('TLabel', padding=2)
        style.configure('Header.TLabel', font=(self._FONT_FAMILY, self._FONT_HEADER, 'bold'))
        style.configure('Status.TLabel', font=(self._FONT_FAMILY, 9))

        # Estilo de destaque para o botão principal
        style.configure('Accent.TButton', padding=(16, 8))
        style.map('Accent.TButton',
                  background=[('active', '#005a9e'), ('!active', '#0078D4')],
                  foreground=[('active', 'white'), ('!active', 'white')])

        # Barra inferior (tema) — criada antes do notebook para ficar na base
        self._setup_bottom_bar()

        # Notebook (tabs) — 5 tabs principais
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=self._PAD_OUTER,
                           pady=(self._PAD_OUTER, 0))

        # Tab 1: Conversão
        self.tab_convert = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_convert, text='Converter')
        self._setup_convert_tab()

        # Tab 2: Perfis
        self.tab_profiles = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_profiles, text='Perfis')
        self._setup_profiles_tab()

        # Tab 3: Multificheiros
        self.tab_batch = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_batch, text='Multificheiros')
        self._setup_batch_tab()

        # Tab 4: Histórico
        self.tab_history = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_history, text='Histórico')
        self._setup_history_tab()

        # Tab 5: Definições (no final)
        self.tab_settings = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_settings, text='Definições')
        self._setup_settings_tab()
    
    def _setup_bottom_bar(self):
        """Barra inferior com controlos globais da interface."""
        bar = ttk.Frame(self.root, padding=(self._PAD_OUTER, 4))
        bar.pack(side='bottom', fill='x')

        ttk.Separator(self.root, orient='horizontal').pack(side='bottom', fill='x')

        current_theme = self.config.get('ui', {}).get('theme', 'light')
        self._theme_btn_text = tk.StringVar(
            value='Tema: Escuro' if current_theme == 'light' else 'Tema: Claro'
        )
        ttk.Button(bar, textvariable=self._theme_btn_text,
                   command=self._toggle_theme).pack(side='right')

    def _toggle_theme(self):
        """Alterna entre tema claro e escuro."""
        if not self._sv_ttk_available:
            messagebox.showinfo("Tema", "O pacote sv-ttk não está instalado.")
            return

        import sv_ttk
        current = self.config.get('ui', {}).get('theme', 'light')
        new_theme = 'dark' if current == 'light' else 'light'

        sv_ttk.set_theme(new_theme)
        self.config.setdefault('ui', {})['theme'] = new_theme
        self._theme_btn_text.set('Tema: Escuro' if new_theme == 'light' else 'Tema: Claro')
        save_config(self.config)

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

        # Sub-tab: Automação
        self.tab_automation = ttk.Frame(settings_nb)
        settings_nb.add(self.tab_automation, text='Automação')
        self._setup_automation_tab()

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
                filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
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

    def _setup_convert_tab(self):
        """Tab de conversão com scroll."""
        # Canvas com scrollbar para conteúdo que não cabe na janela
        canvas = tk.Canvas(self.tab_convert, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.tab_convert, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        frame = ttk.Frame(canvas, padding=self._PAD_OUTER)
        frame_id = canvas.create_window((0, 0), window=frame, anchor='nw')

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox('all'))

        def _on_canvas_configure(event):
            canvas.itemconfig(frame_id, width=event.width)

        frame.bind('<Configure>', _on_frame_configure)
        canvas.bind('<Configure>', _on_canvas_configure)

        # Scroll com roda do rato (Linux)
        canvas.bind_all('<Button-4>', lambda e: canvas.yview_scroll(-3, 'units'))
        canvas.bind_all('<Button-5>', lambda e: canvas.yview_scroll(3, 'units'))

        # --- Barra de progresso e status (topo, sempre visível) ---
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var,
                                            maximum=100, mode='determinate')
        self.progress_bar.pack(fill='x', pady=(0, 4))

        self.status_var = tk.StringVar(value="Pronto  |  Ctrl+O: Abrir  Ctrl+G: Gerar  Ctrl+S: Guardar")
        ttk.Label(frame, textvariable=self.status_var, style='Status.TLabel',
                  foreground='#666666').pack(pady=(0, 10))

        # --- Ficheiros ---
        files_frame = ttk.LabelFrame(frame, text="Ficheiros", padding=self._PAD_INNER)
        files_frame.pack(fill='x', pady=self._PAD_SECTION)
        files_frame.columnconfigure(1, weight=1)

        ttk.Label(files_frame, text="Excel:").grid(row=0, column=0, sticky='e', padx=(0, 8), pady=4)
        ttk.Entry(files_frame, textvariable=self.excel_path).grid(row=0, column=1, sticky='ew', pady=4)
        ttk.Button(files_frame, text="Procurar...", command=self._browse_excel).grid(
            row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(files_frame, text="PDF saída:").grid(row=1, column=0, sticky='e', padx=(0, 8), pady=4)
        ttk.Entry(files_frame, textvariable=self.output_path).grid(row=1, column=1, sticky='ew', pady=4)
        ttk.Button(files_frame, text="Procurar...", command=self._browse_output).grid(
            row=1, column=2, padx=(8, 0), pady=4)

        # --- Opções e Segurança lado a lado ---
        opts_sec_frame = ttk.Frame(frame)
        opts_sec_frame.pack(fill='x', pady=self._PAD_SECTION)

        options_frame = ttk.LabelFrame(opts_sec_frame, text="Opções", padding=self._PAD_INNER)
        options_frame.pack(side='left', fill='both', expand=True, padx=(0, 6))

        self.auto_open_var = tk.BooleanVar(value=self.config['output']['auto_open'])
        ttk.Checkbutton(options_frame, text="Abrir PDF após conversão",
                       variable=self.auto_open_var).pack(anchor='w', pady=2)

        self.add_timestamp_var = tk.BooleanVar(value=self.config['output']['add_timestamp'])
        ttk.Checkbutton(options_frame, text="Data/hora no nome do ficheiro",
                       variable=self.add_timestamp_var).pack(anchor='w', pady=2)

        security_frame = ttk.LabelFrame(opts_sec_frame, text="Segurança", padding=self._PAD_INNER)
        security_frame.pack(side='left', fill='both', expand=True, padx=(6, 0))

        pw_row = ttk.Frame(security_frame)
        pw_row.pack(fill='x', pady=2)
        ttk.Label(pw_row, text="Palavra-passe:").pack(side='left')
        self.pdf_password_var = tk.StringVar(value=self.config.get('security', {}).get('pdf_password', ''))
        ttk.Entry(pw_row, textvariable=self.pdf_password_var, width=14, show='*').pack(side='left', padx=(8, 0))

        wm_row = ttk.Frame(security_frame)
        wm_row.pack(fill='x', pady=2)
        self.watermark_enabled_var = tk.BooleanVar(value=self.config.get('watermark', {}).get('enabled', False))
        ttk.Checkbutton(wm_row, text="Marca d'água:", variable=self.watermark_enabled_var).pack(side='left')
        self.watermark_text_var = tk.StringVar(value=self.config.get('watermark', {}).get('text', 'RASCUNHO'))
        ttk.Combobox(wm_row, textvariable=self.watermark_text_var, width=13,
                     values=['RASCUNHO', 'CÓPIA', 'CONFIDENCIAL', 'ORIGINAL']).pack(side='left', padx=(8, 0))

        # --- Modo de geração ---
        mode_frame = ttk.LabelFrame(frame, text="Modo de Geração", padding=self._PAD_INNER)
        mode_frame.pack(fill='x', pady=self._PAD_SECTION)

        mode_left = ttk.Frame(mode_frame)
        mode_left.pack(side='left', fill='x', expand=True)
        self.generation_mode_var = tk.StringVar(value='individual')
        ttk.Radiobutton(mode_left, text="Por Linha (um PDF por cliente)",
                       variable=self.generation_mode_var, value='individual').pack(anchor='w', pady=1)
        ttk.Radiobutton(mode_left, text="Agregado (todos num único PDF)",
                       variable=self.generation_mode_var, value='aggregate').pack(anchor='w', pady=1)

        mode_right = ttk.Frame(mode_frame)
        mode_right.pack(side='right')
        ttk.Button(mode_right, text="Filtrar Clientes...", command=self._open_client_filter).pack(anchor='e')
        self.client_filter_label = ttk.Label(mode_right, text="Todos os clientes",
                                             foreground='#888888', style='Status.TLabel')
        self.client_filter_label.pack(anchor='e', pady=(4, 0))
        self._client_filter = None

        # --- Separador antes dos botões ---
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(12, 8))

        # --- Ações ---
        actions_frame = ttk.Frame(frame)
        actions_frame.pack(fill='x', pady=(0, 4))

        # Botão principal com destaque
        generate_btn = ttk.Button(actions_frame, text="Gerar PDF(s)",
                                 command=self._generate, style='Accent.TButton')
        generate_btn.pack(side='left', padx=(0, 6))

        ttk.Button(actions_frame, text="Exportar Excel",
                   command=self._export_excel).pack(side='left', padx=6)

        self.email_btn = ttk.Button(actions_frame, text="Enviar Email",
                                    command=self._send_email, state='disabled')
        self.email_btn.pack(side='left', padx=6)

        # Menu "Mais..." para ações secundárias
        self._more_menu = tk.Menu(self.root, tearoff=0)
        self._more_menu.add_command(label="Pré-visualizar dados", command=self._preview_excel)
        self._more_menu.add_command(label="Pré-visualizar PDF", command=self._preview_pdf)
        self._more_menu.add_command(label="Abrir pasta de destino", command=self._open_output_folder)
        self._more_menu.add_separator()
        self._more_menu.add_command(label="Resumo IRS", command=self._show_irs_summary)
        self._more_menu.add_separator()
        self._more_menu.add_command(label="Guardar configurações", command=self._save_config)

        def _show_more_menu():
            btn = more_btn
            self._more_menu.tk_popup(btn.winfo_rootx(), btn.winfo_rooty() + btn.winfo_height())

        more_btn = ttk.Button(actions_frame, text="Mais...", command=_show_more_menu)
        more_btn.pack(side='right', padx=(6, 0))
    
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
        ttk.Button(sel_frame, text="Desmarcar Todos", command=deselect_all).pack(side='left', padx=5)

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

    def _preview_pdf(self):
        """Abre janela de pré-visualização do último PDF gerado."""
        if not self._last_generated_files:
            messagebox.showinfo("Info", "Gere um PDF primeiro para pré-visualizar.")
            return

        pdf_path = self._last_generated_files[0]
        if not os.path.isfile(pdf_path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado:\n{pdf_path}")
            return

        try:
            from src.pdf_preview import render_page
        except ImportError:
            messagebox.showerror("Erro", "PyMuPDF não instalado.\npip install PyMuPDF")
            return

        preview = tk.Toplevel(self.root)
        preview.title(f"Pré-visualização — {os.path.basename(pdf_path)}")
        preview.geometry("700x900")
        preview.transient(self.root)

        current_page = [0]

        # Toolbar
        toolbar = ttk.Frame(preview, padding=5)
        toolbar.pack(fill='x')

        page_label = ttk.Label(toolbar, text="")
        page_label.pack(side='left', padx=8)

        def show_page(idx):
            try:
                img, total = render_page(pdf_path, idx)
            except Exception as e:
                messagebox.showerror("Erro", str(e), parent=preview)
                return
            current_page[0] = idx
            page_label.config(text=f"Página {idx + 1} / {total}")
            prev_btn.config(state='normal' if idx > 0 else 'disabled')
            next_btn.config(state='normal' if idx < total - 1 else 'disabled')

            # Redimensionar para caber na janela
            max_w = preview.winfo_width() - 30 or 670
            max_h = preview.winfo_height() - 80 or 820
            ratio = min(max_w / img.width, max_h / img.height, 1.0)
            if ratio < 1.0:
                new_size = (int(img.width * ratio), int(img.height * ratio))
                img = img.resize(new_size)

            from PIL import ImageTk
            photo = ImageTk.PhotoImage(img)
            canvas.delete('all')
            canvas.create_image(0, 0, anchor='nw', image=photo)
            canvas._photo = photo  # manter referência
            canvas.config(scrollregion=(0, 0, img.width, img.height))

        prev_btn = ttk.Button(toolbar, text="< Anterior", command=lambda: show_page(current_page[0] - 1))
        prev_btn.pack(side='left', padx=4)
        next_btn = ttk.Button(toolbar, text="Seguinte >", command=lambda: show_page(current_page[0] + 1))
        next_btn.pack(side='left', padx=4)

        canvas = tk.Canvas(preview, bg='#e0e0e0')
        canvas.pack(fill='both', expand=True)

        show_page(0)

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
        # Obter larguras de colunas configuradas (ignorar zeros)
        contab_col_widths = {}
        if hasattr(self, 'contab_col_widths_vars'):
            for col, var in self.contab_col_widths_vars.items():
                try:
                    v = float(var.get())
                    if v > 0:
                        contab_col_widths[col] = v
                except (ValueError, TypeError):
                    pass
        
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
                'company_website': self.company_website_var.get(),
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
                'filename_template': self.filename_template_var.get()
                    if hasattr(self, 'filename_template_var') else '',
            },
            'contabilidade': {
                'enabled': True,
                'colunas': contab_colunas,
                'destacar_total': self.contab_destacar_total_var.get() if hasattr(self, 'contab_destacar_total_var') else True,
                'destacar_valores': self.contab_destacar_valores_var.get() if hasattr(self, 'contab_destacar_valores_var') else True,
                'col_widths': contab_col_widths,
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
            'qrcode': {
                'enabled': self.qr_enabled_var.get() if hasattr(self, 'qr_enabled_var') else False,
                'content': self.qr_content_var.get() if hasattr(self, 'qr_content_var') else 'nif',
                'size_mm': self.qr_size_var.get() if hasattr(self, 'qr_size_var') else 25,
            },
            'fonts': self._get_fonts_from_ui(),
            'banking': self._get_banking_from_ui(),
            'automation': self._get_automation_from_ui(),
            'recent': self.config.get('recent', {'last_excel_dir': '', 'last_output_dir': ''}),
            'ui': {
                'theme': self.config.get('ui', {}).get('theme', 'light'),
                'notifications_enabled': self.notifications_enabled_var.get()
                    if hasattr(self, 'notifications_enabled_var') else True,
            },
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

                # Atualizar cache de clientes
                self._cache_clients_from_data(excel_path, data)

                self.root.after(0, lambda: self.status_var.set("A gerar PDF..."))
                self.root.after(0, lambda: self.progress_var.set(60))
                result_path = converter.generate_pdf(client_filter=self._client_filter)

                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(0, lambda: self.status_var.set(
                    f"PDF gerado: {os.path.basename(result_path)} ({clients_count} clientes)"))

                history.add_entry(excel_path, result_path, 'aggregate', clients_count, True)
                self.root.after(0, lambda n=clients_count: notifier.notify(
                    "Conversão concluída",
                    f"{n} cliente(s) — {os.path.basename(result_path)}",
                    self.config,
                ))

                self._last_generated_files = [result_path]
                self.root.after(0, lambda: self.email_btn.configure(state='normal'))

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

                self.root.after(0, lambda: self.progress_var.set(20))
                data = converter.read_excel_data()
                self._cache_clients_from_data(excel_path, data)

                self.root.after(0, lambda: self.progress_var.set(40))
                result_files = converter.generate_individual_pdfs(client_filter=self._client_filter)

                self.root.after(0, lambda: self.progress_var.set(100))

                if result_files:
                    folder = os.path.dirname(result_files[0])
                    self.root.after(0, lambda: self.status_var.set(
                        f"{len(result_files)} PDFs gerados!"))

                    history.add_entry(excel_path, folder, 'individual', len(result_files), True)
                    self.root.after(0, lambda n=len(result_files): notifier.notify(
                        "Conversão concluída",
                        f"{n} PDF(s) gerado(s)",
                        self.config,
                    ))

                    self._last_generated_files = list(result_files)
                    self.root.after(0, lambda: self.email_btn.configure(state='normal'))

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
    
    def _cache_clients_from_data(self, excel_path: str, data: dict):
        """Extrai clientes dos dados e atualiza a cache SQLite."""
        try:
            itens = data.get('itens', [])
            headers = data.get('headers', [])
            # Determinar índices de colunas relevantes
            h_lower = [h.lower().strip() if h else '' for h in headers]
            name_idx = None
            sigla_idx = None
            nif_idx = None
            for i, h in enumerate(h_lower):
                if h == 'cliente':
                    name_idx = i
                elif h == 'sigla':
                    sigla_idx = i
                elif h == 'nif':
                    nif_idx = i
            if name_idx is None:
                return
            clients = []
            for row in itens:
                name = str(row[name_idx]).strip() if name_idx < len(row) else ''
                if not name:
                    continue
                sigla = str(row[sigla_idx]).strip() if sigla_idx is not None and sigla_idx < len(row) else ''
                nif = str(row[nif_idx]).strip() if nif_idx is not None and nif_idx < len(row) else ''
                clients.append({'name': name, 'sigla': sigla, 'nif': nif})
            if clients:
                update_client_cache(os.path.basename(excel_path), clients)
        except Exception:
            pass  # Cache é best-effort

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
        frame = ttk.Frame(self.tab_profiles, padding=self._PAD_OUTER)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Perfis de Configuração", style='Header.TLabel').pack(anchor='w', pady=(0, 4))
        ttk.Label(frame, text="Guarde diferentes configurações como perfis reutilizáveis.",
                 foreground='#666666', style='Status.TLabel').pack(anchor='w', pady=(0, 10))

        # Lista de perfis
        list_frame = ttk.LabelFrame(frame, text="Perfis Guardados", padding=self._PAD_INNER)
        list_frame.pack(fill='both', expand=True, pady=self._PAD_SECTION)

        self.profiles_listbox = tk.Listbox(list_frame, height=8,
                                           font=(self._FONT_FAMILY, self._FONT_SIZE))
        self.profiles_listbox.pack(fill='both', expand=True)

        # Botões
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=(8, 0))

        ttk.Button(btn_frame, text="Guardar Perfil Atual", command=self._save_profile).pack(side='left', padx=(0, 6))
        ttk.Button(btn_frame, text="Carregar Perfil", command=self._load_profile).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Apagar Perfil", command=self._delete_profile).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Atualizar", command=self._refresh_profiles).pack(side='right')

        # Exportar / Importar configurações
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=(16, 8))
        ttk.Label(frame, text="Exportar / Importar Configurações",
                  style='Header.TLabel').pack(anchor='w', pady=(0, 4))
        ttk.Label(frame,
                  text="Partilhe ou faça cópia de segurança das configurações entre máquinas.",
                  foreground='#666666', style='Status.TLabel').pack(anchor='w', pady=(0, 8))

        io_frame = ttk.Frame(frame)
        io_frame.pack(anchor='w')
        ttk.Button(io_frame, text="Exportar Configurações...",
                   command=self._export_config).pack(side='left', padx=(0, 6))
        ttk.Button(io_frame, text="Importar Configurações...",
                   command=self._import_config).pack(side='left')

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

    def _export_config(self):
        """Exporta a configuração atual para um ficheiro JSON."""
        path = filedialog.asksaveasfilename(
            title="Exportar Configurações",
            defaultextension=".json",
            filetypes=[("Ficheiro JSON", "*.json"), ("Todos os ficheiros", "*.*")],
        )
        if not path:
            return
        config = self._get_config_from_ui()
        if export_config(config, path):
            messagebox.showinfo("Sucesso", f"Configurações exportadas para:\n{path}")
        else:
            messagebox.showerror("Erro", "Não foi possível exportar as configurações.")

    def _import_config(self):
        """Importa configurações a partir de um ficheiro JSON."""
        path = filedialog.askopenfilename(
            title="Importar Configurações",
            filetypes=[("Ficheiro JSON", "*.json"), ("Todos os ficheiros", "*.*")],
        )
        if not path:
            return
        try:
            imported = import_config(path)
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("Erro", f"Não foi possível importar:\n{e}")
            return

        self.config = imported
        save_config(self.config)
        self._reload_config_to_ui()
        messagebox.showinfo("Sucesso", "Configurações importadas e aplicadas.")

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
        self.company_website_var.set(cfg['header'].get('company_website', ''))
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
        if hasattr(self, 'filename_template_var'):
            self.filename_template_var.set(cfg.get('output', {}).get('filename_template', ''))
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
        if hasattr(self, 'contab_col_widths_vars'):
            col_widths_cfg = contab_cfg.get('col_widths', {})
            for col, var in self.contab_col_widths_vars.items():
                var.set(str(col_widths_cfg.get(col, 0)))
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
        # UI
        theme = cfg.get('ui', {}).get('theme', 'light')
        if self._sv_ttk_available:
            import sv_ttk
            sv_ttk.set_theme(theme)
        self._theme_btn_text.set('Tema: Escuro' if theme == 'light' else 'Tema: Claro')
        if hasattr(self, 'notifications_enabled_var'):
            self.notifications_enabled_var.set(
                cfg.get('ui', {}).get('notifications_enabled', True))
        # QR Code
        if hasattr(self, 'qr_enabled_var'):
            qr_cfg = cfg.get('qrcode', {})
            self.qr_enabled_var.set(qr_cfg.get('enabled', False))
            self.qr_content_var.set(qr_cfg.get('content', 'nif'))
            self.qr_size_var.set(qr_cfg.get('size_mm', 25))
        # Fonts
        if hasattr(self, 'body_font_var'):
            fonts_cfg = cfg.get('fonts', {})
            self.body_font_var.set(fonts_cfg.get('body_font', 'Helvetica'))
            self.header_font_var.set(fonts_cfg.get('header_font', 'Helvetica-Bold'))
        if hasattr(self, 'fonts_tree'):
            for item in self.fonts_tree.get_children():
                self.fonts_tree.delete(item)
            for entry in cfg.get('fonts', {}).get('registered', []):
                self.fonts_tree.insert('', 'end', values=(entry.get('name', ''), entry.get('path', '')))
        # Automação
        auto_cfg = cfg.get('automation', {})
        if hasattr(self, 'watch_folder_var'):
            self.watch_folder_var.set(auto_cfg.get('watch_folder', ''))
        if hasattr(self, 'watch_enabled_var'):
            self.watch_enabled_var.set(auto_cfg.get('watch_enabled', False))
        if hasattr(self, 'watch_mode_var'):
            self.watch_mode_var.set(auto_cfg.get('watch_mode', 'individual'))
        if hasattr(self, 'watch_interval_var'):
            self.watch_interval_var.set(auto_cfg.get('watch_interval', 5))
        if hasattr(self, 'schedules_tree'):
            self._reload_schedules_tree()
        if hasattr(self, 'hooks_tree'):
            self._reload_hooks_tree()

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

    def _send_email(self):
        """Abre o cliente de email com os últimos PDFs gerados em anexo."""
        if not self._last_generated_files:
            messagebox.showwarning("Aviso", "Nenhum PDF gerado nesta sessão.")
            return
        success, msg = open_email_client(self._last_generated_files)
        if not success:
            messagebox.showerror("Erro", msg)

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
