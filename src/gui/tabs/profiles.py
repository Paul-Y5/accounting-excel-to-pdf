# -*- coding: utf-8 -*-
"""Tab Perfis — gestão de perfis e exportação/importação de configurações."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from src.config import (
    save_config, export_config, import_config,
    list_profiles, save_profile, load_profile, delete_profile,
)


class ProfilesTabMixin:
    """Métodos da tab Perfis."""

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
