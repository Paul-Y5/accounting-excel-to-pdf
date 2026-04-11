#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Testes para o módulo de post-conversion hooks.
"""

import sys
import pytest

from src.hooks import run_hooks


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _config(hooks):
    return {'automation': {'hooks': hooks}}


# ---------------------------------------------------------------------------
# TestRunHooks
# ---------------------------------------------------------------------------

class TestRunHooks:
    def test_no_hooks_returns_empty(self):
        assert run_hooks({}, '/src.xlsx', ['/out.pdf']) == []

    def test_empty_hooks_list(self):
        assert run_hooks(_config([]), '/src.xlsx', ['/out.pdf']) == []

    def test_disabled_hook_skipped(self):
        hooks = [{'name': 'skip', 'command': 'echo ok', 'enabled': False, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', ['/out.pdf'])
        assert results == []

    def test_successful_hook(self):
        hooks = [{'name': 'echo', 'command': 'echo hello', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', ['/out.pdf'])
        assert len(results) == 1
        assert results[0]['returncode'] == 0
        assert 'hello' in results[0]['stdout']

    def test_failed_command(self):
        hooks = [{'name': 'fail', 'command': 'exit 1', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', ['/out.pdf'])
        assert results[0]['returncode'] != 0

    def test_variable_substitution_source(self, tmp_path):
        src = str(tmp_path / 'test.xlsx')
        hooks = [{'name': 'src', 'command': f'echo {{source}}', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), src, [])
        assert src in results[0]['stdout']

    def test_variable_substitution_output(self, tmp_path):
        out = str(tmp_path / 'out.pdf')
        hooks = [{'name': 'out', 'command': 'echo {output}', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', [out])
        assert out in results[0]['stdout']

    def test_variable_substitution_outputs(self, tmp_path):
        outs = [str(tmp_path / 'a.pdf'), str(tmp_path / 'b.pdf')]
        hooks = [{'name': 'outs', 'command': 'echo {outputs}', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', outs)
        assert ','.join(outs) in results[0]['stdout']

    def test_variable_substitution_folder(self, tmp_path):
        out = str(tmp_path / 'sub' / 'out.pdf')
        hooks = [{'name': 'folder', 'command': 'echo {folder}', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', [out])
        assert str(tmp_path / 'sub') in results[0]['stdout']

    def test_timeout_returns_error(self):
        hooks = [{'name': 'slow', 'command': 'sleep 10', 'enabled': True, 'timeout': 0.01}]
        results = run_hooks(_config(hooks), '/src.xlsx', [])
        assert results[0]['error'] != '' or results[0]['returncode'] is None

    def test_empty_command_skipped(self):
        hooks = [{'name': 'empty', 'command': '', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', [])
        assert results == []

    def test_multiple_hooks_all_run(self):
        hooks = [
            {'name': 'h1', 'command': 'echo 1', 'enabled': True, 'timeout': 5},
            {'name': 'h2', 'command': 'echo 2', 'enabled': True, 'timeout': 5},
        ]
        results = run_hooks(_config(hooks), '/src.xlsx', [])
        assert len(results) == 2

    def test_missing_outputs_empty(self):
        hooks = [{'name': 'h', 'command': 'echo {output}', 'enabled': True, 'timeout': 5}]
        results = run_hooks(_config(hooks), '/src.xlsx', [])
        assert results[0]['returncode'] == 0
