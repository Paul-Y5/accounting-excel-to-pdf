#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validação de NIF (Número de Identificação Fiscal) português.
Usa o algoritmo do módulo 11 conforme especificação da AT.
"""


def validate_nif(nif: str) -> tuple[bool, str]:
    """Valida um NIF português.

    Args:
        nif: NIF a validar (pode conter espaços e prefixo PT).

    Returns:
        Tuplo (is_valid, message).
    """
    if not nif:
        return False, "NIF vazio"

    # Limpar: remover espaços, pontos, hífens e prefixo PT
    cleaned = nif.strip().upper()
    cleaned = cleaned.replace(' ', '').replace('.', '').replace('-', '')
    if cleaned.startswith('PT'):
        cleaned = cleaned[2:]

    # Verificar se tem 9 dígitos
    if not cleaned.isdigit():
        return False, f"NIF contém caracteres inválidos: {nif}"

    if len(cleaned) != 9:
        return False, f"NIF deve ter 9 dígitos, tem {len(cleaned)}: {nif}"

    # Verificar primeiro dígito (tipo de contribuinte)
    first_digit = int(cleaned[0])
    valid_first_digits = [1, 2, 3, 5, 6, 7, 8, 9]
    if first_digit not in valid_first_digits:
        return False, f"NIF com primeiro dígito inválido ({first_digit}): {nif}"

    # Algoritmo módulo 11
    weights = [9, 8, 7, 6, 5, 4, 3, 2]
    total = sum(int(cleaned[i]) * weights[i] for i in range(8))

    remainder = total % 11
    check_digit = 0 if remainder < 2 else 11 - remainder

    if int(cleaned[8]) != check_digit:
        return False, f"NIF com dígito de controlo inválido: {nif}"

    return True, "NIF válido"


def validate_nif_list(nifs: list[str]) -> list[dict]:
    """Valida uma lista de NIFs.

    Returns:
        Lista de dicionários com resultado de cada validação.
    """
    results = []
    for nif in nifs:
        is_valid, message = validate_nif(nif)
        results.append({
            'nif': nif,
            'valid': is_valid,
            'message': message,
        })
    return results
