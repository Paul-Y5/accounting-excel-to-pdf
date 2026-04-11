#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Testes para o módulo de watch folder.
"""

import os
import time
import tempfile
import pytest

from src.watch_folder import WatchFolder


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def tmp_folder():
    with tempfile.TemporaryDirectory() as d:
        yield d


@pytest.fixture
def basic_config():
    return {'automation': {'watch_mode': 'individual', 'watch_interval': 1}, 'hooks': []}


# ---------------------------------------------------------------------------
# TestWatchFolderScan
# ---------------------------------------------------------------------------

class TestWatchFolderScan:
    def test_scan_empty_folder(self, tmp_folder, basic_config):
        wf = WatchFolder(tmp_folder, basic_config)
        assert wf._scan() == []

    def test_scan_finds_xlsx(self, tmp_folder, basic_config):
        path = os.path.join(tmp_folder, 'teste.xlsx')
        open(path, 'w').close()
        wf = WatchFolder(tmp_folder, basic_config)
        assert path in wf._scan()

    def test_scan_ignores_temp_files(self, tmp_folder, basic_config):
        path = os.path.join(tmp_folder, '~$temp.xlsx')
        open(path, 'w').close()
        wf = WatchFolder(tmp_folder, basic_config)
        assert wf._scan() == []

    def test_scan_ignores_non_excel(self, tmp_folder, basic_config):
        open(os.path.join(tmp_folder, 'ficheiro.pdf'), 'w').close()
        open(os.path.join(tmp_folder, 'ficheiro.txt'), 'w').close()
        wf = WatchFolder(tmp_folder, basic_config)
        assert wf._scan() == []

    def test_scan_finds_xls(self, tmp_folder, basic_config):
        path = os.path.join(tmp_folder, 'teste.xls')
        open(path, 'w').close()
        wf = WatchFolder(tmp_folder, basic_config)
        assert path in wf._scan()


# ---------------------------------------------------------------------------
# TestWatchFolderLifecycle
# ---------------------------------------------------------------------------

class TestWatchFolderLifecycle:
    def test_start_stop(self, tmp_folder, basic_config):
        wf = WatchFolder(tmp_folder, basic_config, interval=1)
        assert not wf.is_running
        wf.start()
        assert wf.is_running
        wf.stop()
        assert not wf.is_running

    def test_start_invalid_folder(self, basic_config):
        wf = WatchFolder('/pasta/nao/existe', basic_config)
        with pytest.raises(ValueError):
            wf.start()

    def test_start_twice_noop(self, tmp_folder, basic_config):
        wf = WatchFolder(tmp_folder, basic_config, interval=1)
        wf.start()
        wf.start()  # segunda chamada não deve lançar excepção
        wf.stop()

    def test_stop_without_start(self, tmp_folder, basic_config):
        wf = WatchFolder(tmp_folder, basic_config)
        wf.stop()  # não deve lançar excepção

    def test_existing_files_not_reprocessed(self, tmp_folder, basic_config):
        """Ficheiros existentes antes do start não devem disparar callbacks."""
        path = os.path.join(tmp_folder, 'existente.xlsx')
        open(path, 'w').close()

        seen = []
        wf = WatchFolder(tmp_folder, basic_config,
                         on_new_file=lambda p: seen.append(p), interval=0.5)
        wf.start()
        time.sleep(1.5)
        wf.stop()
        assert path not in seen

    def test_new_file_triggers_callback(self, tmp_folder, basic_config):
        """Novo ficheiro criado após start deve disparar on_new_file."""
        seen = []
        wf = WatchFolder(tmp_folder, basic_config,
                         on_new_file=lambda p: seen.append(p), interval=0.5)
        wf.start()
        time.sleep(0.3)
        new_path = os.path.join(tmp_folder, 'novo.xlsx')
        open(new_path, 'w').close()
        time.sleep(3.0)
        wf.stop()
        assert new_path in seen
