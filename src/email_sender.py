#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo para envio de PDFs por email.
Abre o cliente de email do sistema com os ficheiros em anexo.
"""

import shutil
import subprocess
import sys
import urllib.parse
import webbrowser


def build_xdg_email_cmd(pdf_paths: list, subject: str = '', body: str = '') -> list:
    """Constrói o comando xdg-email com anexos (Linux).

    Args:
        pdf_paths: Lista de caminhos de PDFs a anexar.
        subject: Assunto do email.
        body: Corpo do email.

    Returns:
        Lista de argumentos para subprocess.
    """
    cmd = ['xdg-email']
    if subject:
        cmd += ['--subject', subject]
    if body:
        cmd += ['--body', body]
    for path in pdf_paths:
        cmd += ['--attach', path]
    return cmd


def open_email_client(pdf_paths, subject: str = '', body: str = '') -> tuple:
    """Abre o cliente de email com os PDFs em anexo.

    Args:
        pdf_paths: Caminho (str) ou lista de caminhos de PDFs.
        subject: Assunto do email (opcional).
        body: Corpo do email (opcional).

    Returns:
        Tuplo (success: bool, message: str).
    """
    if isinstance(pdf_paths, str):
        pdf_paths = [pdf_paths]

    if sys.platform == 'linux':
        if shutil.which('xdg-email'):
            cmd = build_xdg_email_cmd(pdf_paths, subject, body)
            subprocess.Popen(cmd)
            return True, "Cliente de email aberto."
        return False, "xdg-email não encontrado. Instale xdg-utils."

    elif sys.platform == 'darwin':
        attachments_script = '\n'.join(
            f'make new attachment with properties '
            f'{{file name:POSIX file "{p}"}} at after last paragraph of content of newMessage'
            for p in pdf_paths
        )
        script = f'''
tell application "Mail"
    set newMessage to make new outgoing message with properties \
{{subject:"{subject}", content:"{body}", visible:true}}
    tell newMessage
        {attachments_script}
    end tell
    activate
end tell
'''
        subprocess.Popen(['osascript', '-e', script])
        return True, "Apple Mail aberto."

    elif sys.platform == 'win32':
        params = {}
        if subject:
            params['subject'] = subject
        if body:
            params['body'] = body
        query = urllib.parse.urlencode(params)
        mailto = f'mailto:?{query}' if query else 'mailto:'
        webbrowser.open(mailto)
        n = len(pdf_paths)
        return True, f"Cliente de email aberto. Anexe manualmente {n} PDF(s)."

    return False, f"Plataforma não suportada: {sys.platform}"
