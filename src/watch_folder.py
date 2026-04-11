#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de monitorização de pasta (Watch Folder).
Detecta novos ficheiros Excel e converte-os automaticamente para PDF.
"""

import os
import threading
import time


class WatchFolder:
    """Monitoriza uma pasta e converte automaticamente novos ficheiros Excel.

    Utiliza polling simples (sem dependência watchdog) para máxima compatibilidade.
    A cada intervalo verifica se há novos ficheiros .xlsx/.xls que não estejam
    já em processamento.

    Args:
        folder_path: Pasta a monitorizar.
        config: Configuração da aplicação.
        on_new_file: Callback chamado com (excel_path) quando um novo ficheiro é detectado.
        on_converted: Callback chamado com (excel_path, output_paths) após conversão.
        on_error: Callback chamado com (excel_path, error_msg) em caso de erro.
        interval: Intervalo entre verificações em segundos (default 5).
    """

    def __init__(self, folder_path: str, config: dict,
                 on_new_file=None, on_converted=None, on_error=None,
                 interval: int = 5):
        self.folder_path = folder_path
        self.config = config
        self.on_new_file = on_new_file
        self.on_converted = on_converted
        self.on_error = on_error
        self.interval = interval

        self._running = False
        self._thread = None
        self._seen: set = set()

    def start(self):
        """Inicia a monitorização em thread de fundo."""
        if self._running:
            return
        if not os.path.isdir(self.folder_path):
            raise ValueError(f"Pasta não encontrada: {self.folder_path}")

        # Registar os ficheiros já existentes para não os reprocessar
        self._seen = set(self._scan())
        self._running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        """Para a monitorização."""
        self._running = False
        if self._thread:
            self._thread.join(timeout=self.interval + 1)
            self._thread = None

    @property
    def is_running(self) -> bool:
        return self._running

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------

    def _scan(self) -> list:
        """Retorna lista de ficheiros Excel na pasta (sem temporários)."""
        if not os.path.isdir(self.folder_path):
            return []
        files = []
        for name in os.listdir(self.folder_path):
            if name.startswith('~$'):
                continue
            if name.lower().endswith(('.xlsx', '.xls')):
                files.append(os.path.join(self.folder_path, name))
        return files

    def _loop(self):
        """Loop principal de monitorização."""
        while self._running:
            try:
                current = set(self._scan())
                new_files = current - self._seen
                for path in sorted(new_files):
                    self._seen.add(path)
                    self._process(path)
            except Exception:
                pass
            time.sleep(self.interval)

    def _process(self, excel_path: str):
        """Converte um ficheiro Excel detectado."""
        if self.on_new_file:
            self.on_new_file(excel_path)
        try:
            from src.converter import ExcelToPDFConverter
            from src.hooks import run_hooks
            converter = ExcelToPDFConverter(excel_path, None, self.config)
            mode = self.config.get('automation', {}).get('watch_mode', 'individual')
            if mode == 'aggregate':
                output = converter.generate_pdf()
                outputs = [output]
            else:
                outputs = converter.generate_individual_pdfs()
            run_hooks(self.config, excel_path, outputs)
            if self.on_converted:
                self.on_converted(excel_path, outputs)
        except Exception as e:
            if self.on_error:
                self.on_error(excel_path, str(e))
