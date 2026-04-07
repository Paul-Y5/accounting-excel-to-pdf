"""
Testes unitários para o módulo de notificações desktop.
"""

from unittest.mock import patch, MagicMock
import pytest

from src import notifier


class TestNotify:
    """Testes para a função notify."""

    def test_returns_false_when_disabled_in_config(self):
        config = {'ui': {'notifications_enabled': False}}
        result = notifier.notify("Título", "Mensagem", config=config)
        assert result is False

    def test_proceeds_when_enabled_in_config(self):
        config = {'ui': {'notifications_enabled': True}}
        with patch('src.notifier.sys') as mock_sys, \
             patch('subprocess.run') as mock_run:
            mock_sys.platform = 'linux'
            mock_run.return_value = MagicMock(returncode=0)
            result = notifier.notify("Título", "Mensagem", config=config)
        assert result is True

    def test_proceeds_when_config_is_none(self):
        with patch('src.notifier.sys') as mock_sys, \
             patch('subprocess.run') as mock_run:
            mock_sys.platform = 'linux'
            mock_run.return_value = MagicMock(returncode=0)
            result = notifier.notify("Título", "Mensagem", config=None)
        assert result is True

    def test_returns_false_on_exception(self):
        with patch('src.notifier.sys') as mock_sys, \
             patch('subprocess.run', side_effect=OSError("falhou")):
            mock_sys.platform = 'linux'
            result = notifier.notify("Título", "Mensagem")
        assert result is False


class TestNotifyLinux:
    """Testes para _notify_linux."""

    def test_calls_notify_send(self):
        with patch('subprocess.run') as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            result = notifier._notify_linux("Título", "Mensagem", 5)
        mock_run.assert_called_once()
        args = mock_run.call_args[0][0]
        assert 'notify-send' in args
        assert 'Título' in args
        assert 'Mensagem' in args

    def test_returns_false_on_nonzero_returncode(self):
        with patch('subprocess.run') as mock_run:
            mock_run.return_value = MagicMock(returncode=1)
            result = notifier._notify_linux("Título", "Mensagem", 5)
        assert result is False

    def test_timeout_converted_to_ms(self):
        with patch('subprocess.run') as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            notifier._notify_linux("T", "M", 3)
        args = mock_run.call_args[0][0]
        assert '3000' in args


class TestNotifyMacOS:
    """Testes para _notify_macos."""

    def test_calls_osascript(self):
        with patch('subprocess.run') as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            result = notifier._notify_macos("Título", "Mensagem")
        mock_run.assert_called_once()
        args = mock_run.call_args[0][0]
        assert 'osascript' in args

    def test_returns_false_on_nonzero_returncode(self):
        with patch('subprocess.run') as mock_run:
            mock_run.return_value = MagicMock(returncode=1)
            result = notifier._notify_macos("Título", "Mensagem")
        assert result is False


class TestNotifyWindows:
    """Testes para _notify_windows."""

    def test_returns_false_when_win10toast_not_installed(self):
        with patch.dict('sys.modules', {'win10toast': None}):
            result = notifier._notify_windows("Título", "Mensagem", 5)
        assert result is False
