#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de geração de QR Codes para inclusão no PDF.
Gera imagens QR a partir de NIF ou IBAN da empresa.
"""

import os
import tempfile


def get_qr_data(config: dict) -> str:
    """Obtém o conteúdo para o QR Code a partir da configuração.

    Args:
        config: Dicionário de configuração da aplicação.

    Returns:
        String com o NIF ou IBAN para codificar, ou vazio se não disponível.
    """
    qr_cfg = config.get('qrcode', {})
    content_type = qr_cfg.get('content', 'nif')

    if content_type == 'iban':
        from src.converter import _get_active_bank
        bank = _get_active_bank(config)
        return bank.get('iban', '').replace(' ', '')
    else:
        # NIF (default)
        nif = config.get('header', {}).get('company_nif', '')
        return nif.replace(' ', '')


def build_qr_image(data: str, size_mm: int = 25) -> str:
    """Gera uma imagem PNG de QR Code num ficheiro temporário.

    Args:
        data: Conteúdo a codificar no QR.
        size_mm: Tamanho aproximado do QR em milímetros (usado como box_size).

    Returns:
        Caminho absoluto do ficheiro PNG temporário.

    Raises:
        ImportError: Se o módulo ``qrcode`` não estiver instalado.
        ValueError: Se ``data`` for vazio.
    """
    if not data:
        raise ValueError("Dados para QR Code não podem ser vazios.")

    try:
        import qrcode
    except ImportError:
        raise ImportError(
            "Para gerar QR Codes, instale o pacote: pip install qrcode[pil]"
        )

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=max(4, size_mm // 5),
        border=2,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")

    fd, path = tempfile.mkstemp(suffix='.png', prefix='qr_')
    os.close(fd)
    img.save(path)

    return path
