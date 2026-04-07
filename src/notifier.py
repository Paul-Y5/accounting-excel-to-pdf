#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de notificações desktop.

Envia notificações nativas do sistema operativo após operações longas
(conversão em batch, geração de PDFs). Requer que a config tenha
ui.notifications_enabled = True para enviar notificações.

Plataformas suportadas:
    Linux  — notify-send (libnotify); fallback silencioso se ausente.
    macOS  — osascript (AppleScript).
    Windows — win10toast; fallback silencioso se não instalado.
"""

import subprocess
import sys


def notify(title: str, message: str, config: dict = None, timeout: int = 5) -> bool:
    """Envia uma notificação desktop.

    Args:
        title:   Título da notificação.
        message: Corpo da mensagem.
        config:  Configuração da aplicação. Se None, considera notificações ativas.
        timeout: Duração em segundos (nem todas as plataformas respeitam este valor).

    Returns:
        True se a notificação foi enviada, False caso contrário.
    """
    if config is not None:
        if not config.get('ui', {}).get('notifications_enabled', True):
            return False

    try:
        if sys.platform == 'linux':
            return _notify_linux(title, message, timeout)
        elif sys.platform == 'darwin':
            return _notify_macos(title, message)
        elif sys.platform == 'win32':
            return _notify_windows(title, message, timeout)
    except Exception:
        pass

    return False


def _notify_linux(title: str, message: str, timeout: int) -> bool:
    """Envia notificação via notify-send (libnotify)."""
    result = subprocess.run(
        ['notify-send', '--expire-time', str(timeout * 1000), title, message],
        capture_output=True,
    )
    return result.returncode == 0


def _notify_macos(title: str, message: str) -> bool:
    """Envia notificação via osascript."""
    script = f'display notification "{message}" with title "{title}"'
    result = subprocess.run(
        ['osascript', '-e', script],
        capture_output=True,
    )
    return result.returncode == 0


def _notify_windows(title: str, message: str, timeout: int) -> bool:
    """Envia notificação via win10toast."""
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(title, message, duration=timeout, threaded=True)
        return True
    except ImportError:
        return False
