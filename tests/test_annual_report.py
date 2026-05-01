#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Testes para src/annual_report.py."""

import os
import pytest

from src.annual_report import get_annual_data, get_available_years, generate_annual_report_excel
from src.database import add_history_entry, init_db


@pytest.fixture(autouse=True)
def _clean_db(tmp_path, monkeypatch):
    """BD temporária isolada para cada teste."""
    import src.database as db_mod
    db_file = str(tmp_path / 'test.db')
    monkeypatch.setattr(db_mod, '_get_db_path', lambda: db_file)
    init_db()
    yield


def _add(source='ficheiro.xlsx', mode='individual', clients=3,
         success=True, timestamp='2026-06-15T10:00:00'):
    """Helper para inserir entradas no histórico com timestamp fixo."""
    import src.database as db_mod
    from datetime import datetime as _dt
    conn = db_mod._get_connection()
    try:
        conn.execute(
            """INSERT INTO history
               (timestamp, source_file, source_path, output_path,
                mode, clients_count, success, error)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (timestamp, source, source, '', mode, clients,
             1 if success else 0, ''),
        )
        conn.commit()
    finally:
        conn.close()


class TestGetAnnualData:
    def test_empty_history(self):
        d = get_annual_data(2026)
        assert d['total'] == 0
        assert d['success'] == 0
        assert d['errors'] == 0
        assert len(d['by_month']) == 12

    def test_counts_conversions(self):
        _add(timestamp='2026-01-10T09:00:00', clients=5)
        _add(timestamp='2026-03-20T09:00:00', clients=2)
        d = get_annual_data(2026)
        assert d['total'] == 2
        assert d['clients_total'] == 7

    def test_counts_errors(self):
        _add(success=True,  timestamp='2026-02-01T09:00:00')
        _add(success=False, timestamp='2026-02-05T09:00:00')
        d = get_annual_data(2026)
        assert d['success'] == 1
        assert d['errors'] == 1

    def test_by_month_aggregation(self):
        _add(timestamp='2026-06-10T09:00:00', clients=3)
        _add(timestamp='2026-06-15T09:00:00', clients=4)
        _add(timestamp='2026-12-01T09:00:00', clients=1)
        d = get_annual_data(2026)
        june = d['by_month'][5]  # índice 0-based
        assert june['conversions'] == 2
        assert june['clients'] == 7
        dec = d['by_month'][11]
        assert dec['conversions'] == 1

    def test_excludes_other_years(self):
        _add(timestamp='2025-12-31T23:59:59')
        _add(timestamp='2026-01-01T00:00:00')
        _add(timestamp='2027-01-01T00:00:00')
        d = get_annual_data(2026)
        assert d['total'] == 1

    def test_top_files(self):
        for _ in range(5):
            _add(source='a.xlsx', timestamp='2026-01-01T09:00:00')
        for _ in range(2):
            _add(source='b.xlsx', timestamp='2026-01-02T09:00:00')
        d = get_annual_data(2026)
        assert d['top_files'][0]['file'] == 'a.xlsx'
        assert d['top_files'][0]['count'] == 5

    def test_by_mode(self):
        _add(mode='individual', timestamp='2026-01-01T09:00:00')
        _add(mode='individual', timestamp='2026-01-02T09:00:00')
        _add(mode='aggregate', timestamp='2026-01-03T09:00:00')
        d = get_annual_data(2026)
        assert d['by_mode']['individual'] == 2
        assert d['by_mode']['aggregate'] == 1


class TestGetAvailableYears:
    def test_empty(self):
        assert get_available_years() == []

    def test_returns_years(self):
        _add(timestamp='2025-05-01T09:00:00')
        _add(timestamp='2026-03-01T09:00:00')
        years = get_available_years()
        assert 2025 in years
        assert 2026 in years

    def test_sorted_desc(self):
        _add(timestamp='2024-01-01T09:00:00')
        _add(timestamp='2026-01-01T09:00:00')
        years = get_available_years()
        assert years[0] > years[-1]


class TestGenerateAnnualReportExcel:
    def test_creates_file(self, tmp_path):
        _add(timestamp='2026-04-01T09:00:00', clients=3)
        output = str(tmp_path / 'relatorio.xlsx')
        path = generate_annual_report_excel(2026, output)
        assert os.path.exists(path)
        assert path.endswith('.xlsx')

    def test_empty_year_still_creates_file(self, tmp_path):
        output = str(tmp_path / 'relatorio_vazio.xlsx')
        path = generate_annual_report_excel(2099, output)
        assert os.path.exists(path)
