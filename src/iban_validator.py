#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validação de IBAN (International Bank Account Number) — ISO 13616.

Suporta IBANs de qualquer país. Inclui comprimentos esperados
para os países da UE e países lusófonos.
"""

# Comprimentos de IBAN por código de país (ISO 13616-1)
_IBAN_LENGTHS: dict[str, int] = {
    'AD': 24, 'AE': 23, 'AL': 28, 'AT': 20, 'AZ': 28,
    'BA': 20, 'BE': 16, 'BG': 22, 'BH': 22, 'BR': 29,
    'BY': 28, 'CH': 21, 'CR': 22, 'CY': 28, 'CZ': 24,
    'DE': 22, 'DK': 18, 'DO': 28, 'EE': 20, 'EG': 29,
    'ES': 24, 'FI': 18, 'FK': 18, 'FR': 27, 'GB': 22,
    'GE': 22, 'GI': 23, 'GL': 18, 'GR': 27, 'GT': 28,
    'HR': 21, 'HU': 28, 'IE': 22, 'IL': 23, 'IQ': 23,
    'IS': 26, 'IT': 27, 'JO': 30, 'KW': 30, 'KZ': 20,
    'LB': 28, 'LC': 32, 'LI': 21, 'LT': 20, 'LU': 20,
    'LV': 21, 'LY': 25, 'MC': 27, 'MD': 24, 'ME': 22,
    'MK': 19, 'MR': 27, 'MT': 31, 'MU': 30, 'MZ': 25,
    'NL': 18, 'NO': 15, 'PK': 24, 'PL': 28, 'PS': 29,
    'PT': 25, 'QA': 29, 'RO': 24, 'RS': 22, 'SA': 24,
    'SC': 31, 'SD': 18, 'SE': 24, 'SI': 19, 'SK': 24,
    'SM': 27, 'ST': 25, 'SV': 28, 'TL': 23, 'TN': 24,
    'TR': 26, 'UA': 29, 'VA': 22, 'VG': 24, 'XK': 20,
}


def validate_iban(iban: str) -> tuple[bool, str]:
    """Valida um IBAN internacional (ISO 13616 / mod-97).

    Args:
        iban: IBAN a validar. Pode conter espaços.

    Returns:
        Tuplo ``(is_valid, mensagem)``.
    """
    if not iban:
        return False, "IBAN vazio"

    # Normalizar: maiúsculas sem espaços
    cleaned = iban.strip().upper().replace(' ', '').replace('-', '')

    if len(cleaned) < 5:
        return False, "IBAN demasiado curto"

    country = cleaned[:2]
    if not country.isalpha():
        return False, f"Código de país inválido: '{country}'"

    # Verificar comprimento esperado
    expected = _IBAN_LENGTHS.get(country)
    if expected and len(cleaned) != expected:
        return (
            False,
            f"Comprimento inválido para {country}: esperado {expected}, tem {len(cleaned)}",
        )

    # Verificar dígitos de controlo (posições 2-3)
    check_digits = cleaned[2:4]
    if not check_digits.isdigit():
        return False, f"Dígitos de controlo inválidos: '{check_digits}'"

    # Algoritmo mod-97: mover os 4 primeiros chars para o fim e converter letras
    rearranged = cleaned[4:] + cleaned[:4]
    numeric = ''.join(
        str(ord(ch) - ord('A') + 10) if ch.isalpha() else ch
        for ch in rearranged
    )

    if int(numeric) % 97 != 1:
        return False, f"IBAN com dígito de controlo inválido: {iban.strip()}"

    return True, "IBAN válido"


def format_iban(iban: str) -> str:
    """Formata um IBAN em grupos de 4 caracteres.

    Exemplo: ``'PT50017037820304005356729'`` → ``'PT50 0170 3782 0304 0053 5672 9'``

    Args:
        iban: IBAN limpo ou já formatado.

    Returns:
        IBAN formatado. Devolve a string original se estiver vazia.
    """
    cleaned = iban.strip().upper().replace(' ', '').replace('-', '')
    if not cleaned:
        return iban
    return ' '.join(cleaned[i:i + 4] for i in range(0, len(cleaned), 4))


def validate_iban_list(ibans: list[str]) -> list[dict]:
    """Valida uma lista de IBANs.

    Returns:
        Lista de dicts com ``iban``, ``valid``, ``message``, ``formatted``.
    """
    results = []
    for iban in ibans:
        is_valid, message = validate_iban(iban)
        results.append({
            'iban': iban,
            'valid': is_valid,
            'message': message,
            'formatted': format_iban(iban) if is_valid else iban.strip(),
        })
    return results
