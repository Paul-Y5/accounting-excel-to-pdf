"""
Testes unitários para o módulo de envio de email.
"""

import sys
import pytest
from unittest.mock import patch

from src.email_sender import build_xdg_email_cmd, open_email_client


class TestBuildXdgEmailCmd:
    """Testes para build_xdg_email_cmd."""

    def test_basic_command(self, tmp_path):
        """Verifica comando mínimo com um anexo."""
        pdf = str(tmp_path / 'test.pdf')
        cmd = build_xdg_email_cmd([pdf])
        assert cmd[0] == 'xdg-email'
        assert '--attach' in cmd
        assert pdf in cmd

    def test_with_subject(self, tmp_path):
        """Verifica que o assunto é incluído."""
        pdf = str(tmp_path / 'test.pdf')
        cmd = build_xdg_email_cmd([pdf], subject='Mapa Contabilidade')
        assert '--subject' in cmd
        assert 'Mapa Contabilidade' in cmd

    def test_with_body(self, tmp_path):
        """Verifica que o corpo é incluído."""
        pdf = str(tmp_path / 'test.pdf')
        cmd = build_xdg_email_cmd([pdf], body='Segue em anexo.')
        assert '--body' in cmd
        assert 'Segue em anexo.' in cmd

    def test_multiple_attachments(self, tmp_path):
        """Verifica múltiplos anexos."""
        pdfs = [str(tmp_path / f'f{i}.pdf') for i in range(3)]
        cmd = build_xdg_email_cmd(pdfs)
        assert cmd.count('--attach') == 3
        for pdf in pdfs:
            assert pdf in cmd

    def test_no_subject_no_flag(self, tmp_path):
        """Verifica que --subject não aparece quando assunto está vazio."""
        pdf = str(tmp_path / 'test.pdf')
        cmd = build_xdg_email_cmd([pdf])
        assert '--subject' not in cmd
        assert '--body' not in cmd

    def test_returns_list(self, tmp_path):
        """Verifica que retorna lista."""
        pdf = str(tmp_path / 'test.pdf')
        cmd = build_xdg_email_cmd([pdf])
        assert isinstance(cmd, list)


class TestOpenEmailClient:
    """Testes para open_email_client."""

    def test_returns_tuple(self, tmp_path):
        """Verifica que sempre retorna (bool, str)."""
        pdf = str(tmp_path / 'test.pdf')
        with patch('subprocess.Popen'), patch('sys.platform', 'linux'), \
             patch('shutil.which', return_value='/usr/bin/xdg-email'):
            result = open_email_client(pdf)
        assert isinstance(result, tuple)
        assert len(result) == 2
        assert isinstance(result[0], bool)
        assert isinstance(result[1], str)

    def test_string_path_converted_to_list(self, tmp_path):
        """Verifica que string é aceite além de lista."""
        pdf = str(tmp_path / 'test.pdf')
        with patch('subprocess.Popen') as mock_popen, \
             patch('sys.platform', 'linux'), \
             patch('shutil.which', return_value='/usr/bin/xdg-email'):
            open_email_client(pdf)
            args = mock_popen.call_args[0][0]
            assert pdf in args

    def test_linux_success_with_xdg_email(self, tmp_path):
        """Linux com xdg-email disponível retorna sucesso."""
        pdf = str(tmp_path / 'test.pdf')
        with patch('subprocess.Popen'), \
             patch('sys.platform', 'linux'), \
             patch('shutil.which', return_value='/usr/bin/xdg-email'):
            success, msg = open_email_client(pdf)
        assert success is True
        assert msg == "Cliente de email aberto."

    def test_linux_failure_without_xdg_email(self, tmp_path):
        """Linux sem xdg-email retorna erro."""
        pdf = str(tmp_path / 'test.pdf')
        with patch('sys.platform', 'linux'), \
             patch('shutil.which', return_value=None):
            success, msg = open_email_client(pdf)
        assert success is False
        assert 'xdg-email' in msg

    def test_linux_calls_popen_with_correct_cmd(self, tmp_path):
        """Verifica que Popen é chamado com o comando correto no Linux."""
        pdf = str(tmp_path / 'test.pdf')
        with patch('subprocess.Popen') as mock_popen, \
             patch('sys.platform', 'linux'), \
             patch('shutil.which', return_value='/usr/bin/xdg-email'):
            open_email_client([pdf], subject='Teste', body='Corpo')
            cmd = mock_popen.call_args[0][0]
            assert 'xdg-email' in cmd[0]
            assert '--subject' in cmd
            assert 'Teste' in cmd
            assert '--attach' in cmd
            assert pdf in cmd

    def test_windows_opens_mailto(self, tmp_path):
        """Windows abre mailto: via webbrowser."""
        pdf = str(tmp_path / 'test.pdf')
        with patch('sys.platform', 'win32'), \
             patch('webbrowser.open') as mock_open:
            success, msg = open_email_client(pdf, subject='Teste')
        assert success is True
        mock_open.assert_called_once()
        url = mock_open.call_args[0][0]
        assert url.startswith('mailto:')
        assert 'Teste' in url

    def test_unsupported_platform(self, tmp_path):
        """Plataforma desconhecida retorna erro."""
        pdf = str(tmp_path / 'test.pdf')
        with patch('sys.platform', 'freebsd7'):
            success, msg = open_email_client(pdf)
        assert success is False
        assert 'freebsd7' in msg

    def test_multiple_pdfs_linux(self, tmp_path):
        """Múltiplos PDFs são todos passados como --attach no Linux."""
        pdfs = [str(tmp_path / f'{i}.pdf') for i in range(3)]
        with patch('subprocess.Popen') as mock_popen, \
             patch('sys.platform', 'linux'), \
             patch('shutil.which', return_value='/usr/bin/xdg-email'):
            open_email_client(pdfs)
            cmd = mock_popen.call_args[0][0]
            assert cmd.count('--attach') == 3
