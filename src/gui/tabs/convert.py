# -*- coding: utf-8 -*-
"""Tab Converter — seleção de ficheiros, opções e ações de geração de PDF."""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src.config import save_config
from src.converter import ExcelToPDFConverter
from src.nif_validator import validate_nif
from src.excel_exporter import export_to_excel
from src import history
from src import notifier
from src.database import update_client_cache
from src.email_sender import open_email_client


class ConvertTabMixin:
    """Métodos da tab Converter e ações de conversão."""

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

    def _send_email(self):
        """Abre o cliente de email com os últimos PDFs gerados em anexo."""
        if not self._last_generated_files:
            messagebox.showwarning("Aviso", "Nenhum PDF gerado nesta sessão.")
            return
        success, msg = open_email_client(self._last_generated_files)
        if not success:
            messagebox.showerror("Erro", msg)

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
