#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de agendamento de conversões.
Permite agendar conversões automáticas por hora e dia da semana.
"""

import os
import threading
import time
from datetime import datetime


# Nomes dos dias da semana em português (0=Segunda, 6=Domingo)
DIAS_SEMANA = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']


class Scheduler:
    """Executa conversões agendadas num horário configurável.

    O agendamento define:
    - Hora de execução (HH:MM)
    - Dias da semana em que executa (lista de 0..6)
    - Pasta de origem ou ficheiro Excel único
    - Modo de conversão ('individual' ou 'aggregate')

    Args:
        config: Configuração da aplicação.
        on_run: Callback chamado antes de cada execução com (schedule_entry).
        on_done: Callback chamado após execução com (schedule_entry, results).
        on_error: Callback chamado em caso de erro com (schedule_entry, error_msg).
    """

    def __init__(self, config: dict, on_run=None, on_done=None, on_error=None):
        self.config = config
        self.on_run = on_run
        self.on_done = on_done
        self.on_error = on_error

        self._running = False
        self._thread = None
        # Registo de última execução por entrada: {entry_id: datetime}
        self._last_run: dict = {}

    def start(self):
        """Inicia o loop de agendamento em thread de fundo."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        """Para o loop de agendamento."""
        self._running = False
        if self._thread:
            self._thread.join(timeout=70)
            self._thread = None

    @property
    def is_running(self) -> bool:
        return self._running

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------

    def _loop(self):
        """Verifica minuto a minuto se há algum agendamento para executar."""
        while self._running:
            now = datetime.now()
            schedules = self.config.get('automation', {}).get('schedules', [])
            for entry in schedules:
                if not entry.get('enabled', True):
                    continue
                if self._should_run(entry, now):
                    entry_id = id(entry)
                    self._last_run[entry_id] = now
                    threading.Thread(
                        target=self._execute,
                        args=(entry,),
                        daemon=True,
                    ).start()
            # Esperar até ao próximo minuto
            time.sleep(60 - datetime.now().second)

    def _should_run(self, entry: dict, now: datetime) -> bool:
        """Verifica se um agendamento deve ser executado agora."""
        hora = entry.get('hora', '')
        if not hora:
            return False
        try:
            h, m = hora.split(':')
            if int(h) != now.hour or int(m) != now.minute:
                return False
        except (ValueError, AttributeError):
            return False

        # Verificar dia da semana (0=Segunda, 6=Domingo como weekday())
        dias = entry.get('dias', list(range(7)))
        if now.weekday() not in dias:
            return False

        # Evitar executar duas vezes no mesmo minuto
        entry_id = id(entry)
        last = self._last_run.get(entry_id)
        if last and last.hour == now.hour and last.minute == now.minute and last.date() == now.date():
            return False

        return True

    def _execute(self, entry: dict):
        """Executa a conversão para um agendamento."""
        if self.on_run:
            self.on_run(entry)
        try:
            source = entry.get('source', '')
            mode = entry.get('mode', 'individual')
            if not source or not os.path.exists(source):
                raise FileNotFoundError(f"Origem não encontrada: {source}")

            from src.hooks import run_hooks

            if os.path.isdir(source):
                from src.batch_processor import process_batch
                results = process_batch(source, self.config, mode=mode)
                output_paths = [r['output_path'] for r in results if r['success']]
                run_hooks(self.config, source, output_paths)
                if self.on_done:
                    self.on_done(entry, results)
            else:
                from src.converter import ExcelToPDFConverter
                converter = ExcelToPDFConverter(source, None, self.config)
                if mode == 'aggregate':
                    output = converter.generate_pdf()
                    outputs = [output]
                else:
                    outputs = converter.generate_individual_pdfs()
                run_hooks(self.config, source, outputs)
                if self.on_done:
                    self.on_done(entry, outputs)

        except Exception as e:
            if self.on_error:
                self.on_error(entry, str(e))


def validate_schedule_entry(entry: dict) -> list:
    """Valida uma entrada de agendamento.

    Returns:
        Lista de erros encontrados (vazia se válida).
    """
    errors = []
    hora = entry.get('hora', '')
    if not hora:
        errors.append("Hora é obrigatória.")
    else:
        try:
            h, m = hora.split(':')
            if not (0 <= int(h) <= 23 and 0 <= int(m) <= 59):
                errors.append("Hora inválida (HH:MM, 00:00–23:59).")
        except (ValueError, AttributeError):
            errors.append("Formato de hora inválido (use HH:MM).")

    dias = entry.get('dias', [])
    if not dias:
        errors.append("Selecione pelo menos um dia da semana.")

    source = entry.get('source', '')
    if not source:
        errors.append("Origem (pasta ou ficheiro) é obrigatória.")

    return errors
