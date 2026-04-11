#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Testes para o módulo de agendamento.
"""

import pytest
from datetime import datetime
from unittest.mock import patch, MagicMock

from src.scheduler import Scheduler, validate_schedule_entry, DIAS_SEMANA


# ---------------------------------------------------------------------------
# TestDiasSemana
# ---------------------------------------------------------------------------

class TestDiasSemana:
    def test_seven_days(self):
        assert len(DIAS_SEMANA) == 7

    def test_starts_with_segunda(self):
        assert DIAS_SEMANA[0] == 'Segunda'

    def test_ends_with_domingo(self):
        assert DIAS_SEMANA[6] == 'Domingo'


# ---------------------------------------------------------------------------
# TestValidateScheduleEntry
# ---------------------------------------------------------------------------

class TestValidateScheduleEntry:
    def _entry(self, **kwargs):
        base = {'hora': '08:00', 'dias': [0, 1, 2, 3, 4], 'source': '/some/path'}
        base.update(kwargs)
        return base

    def test_valid_entry_no_errors(self, tmp_path):
        e = self._entry(source=str(tmp_path))
        assert validate_schedule_entry(e) == []

    def test_missing_hora(self):
        e = self._entry(hora='', source='/p')
        errors = validate_schedule_entry(e)
        assert any('Hora' in err for err in errors)

    def test_invalid_hora_format(self):
        e = self._entry(hora='25:00', source='/p')
        errors = validate_schedule_entry(e)
        assert any('inválida' in err or 'inválido' in err for err in errors)

    def test_invalid_hora_text(self):
        e = self._entry(hora='abc', source='/p')
        errors = validate_schedule_entry(e)
        assert errors

    def test_empty_dias(self):
        e = self._entry(dias=[], source='/p')
        errors = validate_schedule_entry(e)
        assert any('dia' in err.lower() for err in errors)

    def test_missing_source(self):
        e = self._entry(source='')
        errors = validate_schedule_entry(e)
        assert any('Origem' in err for err in errors)

    def test_multiple_errors(self):
        e = {'hora': '', 'dias': [], 'source': ''}
        errors = validate_schedule_entry(e)
        assert len(errors) >= 2


# ---------------------------------------------------------------------------
# TestSchedulerShouldRun
# ---------------------------------------------------------------------------

class TestSchedulerShouldRun:
    def _sched(self):
        return Scheduler({})

    def _now(self, hour, minute, weekday=0):
        """Cria um datetime mock com hora/minuto/dia da semana."""
        dt = MagicMock(spec=datetime)
        dt.hour = hour
        dt.minute = minute
        dt.weekday.return_value = weekday
        dt.date.return_value = object()
        return dt

    def test_matching_hour_and_day(self):
        s = self._sched()
        entry = {'hora': '08:30', 'dias': [0], 'source': '/p', 'enabled': True}
        now = self._now(8, 30, 0)
        assert s._should_run(entry, now)

    def test_wrong_hour(self):
        s = self._sched()
        entry = {'hora': '08:30', 'dias': [0], 'source': '/p', 'enabled': True}
        now = self._now(9, 30, 0)
        assert not s._should_run(entry, now)

    def test_wrong_day(self):
        s = self._sched()
        entry = {'hora': '08:30', 'dias': [1], 'source': '/p', 'enabled': True}
        now = self._now(8, 30, 0)  # weekday=Segunda (0), entry exige Terça (1)
        assert not s._should_run(entry, now)

    def test_empty_hora(self):
        s = self._sched()
        entry = {'hora': '', 'dias': [0], 'source': '/p'}
        now = self._now(8, 30, 0)
        assert not s._should_run(entry, now)

    def test_all_days_allowed(self):
        s = self._sched()
        entry = {'hora': '10:00', 'dias': list(range(7)), 'source': '/p', 'enabled': True}
        for day in range(7):
            now = self._now(10, 0, day)
            assert s._should_run(entry, now)


# ---------------------------------------------------------------------------
# TestSchedulerLifecycle
# ---------------------------------------------------------------------------

class TestSchedulerLifecycle:
    def test_start_stop(self):
        s = Scheduler({})
        assert not s.is_running
        s.start()
        assert s.is_running
        s.stop()
        assert not s.is_running

    def test_start_twice_noop(self):
        s = Scheduler({})
        s.start()
        s.start()
        s.stop()
        assert not s.is_running

    def test_stop_without_start(self):
        s = Scheduler({})
        s.stop()  # não deve lançar excepção
