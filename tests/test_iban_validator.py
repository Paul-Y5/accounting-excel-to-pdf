#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Testes para src/iban_validator.py."""

import pytest

from src.iban_validator import format_iban, validate_iban, validate_iban_list


class TestValidateIban:
    # IBANs válidos de vários países
    @pytest.mark.parametrize("iban", [
        "PT50 0170 3782 0304 0053 5672 9",   # Portugal (25)
        "PT500170378203040053567 29",         # PT com espaço irregular
        "GB29 NWBK 6016 1331 9268 19",       # Reino Unido (22)
        "DE89 3704 0044 0532 0130 00",        # Alemanha (22)
        "FR76 3000 6000 0112 3456 7890 189",  # França (27)
        "ES91 2100 0418 4502 0005 1332",      # Espanha (24)
        "NL91 ABNA 0417 1643 00",             # Países Baixos (18)
    ])
    def test_valid_ibans(self, iban):
        ok, msg = validate_iban(iban)
        assert ok, f"Expected valid IBAN but got: {msg} for '{iban}'"

    def test_empty_iban(self):
        ok, msg = validate_iban("")
        assert not ok
        assert "vazio" in msg.lower()

    def test_wrong_length_pt(self):
        ok, msg = validate_iban("PT50 0170 3782")  # curto demais
        assert not ok
        assert "PT" in msg

    def test_invalid_check_digit(self):
        # Trocar os dígitos de controlo de 50 para 99
        ok, msg = validate_iban("PT99 0170 3782 0304 0053 5672 9")
        assert not ok

    def test_invalid_country_code(self):
        ok, msg = validate_iban("12 3456 7890")
        assert not ok

    def test_non_alpha_country(self):
        ok, msg = validate_iban("12AB0001234567890")
        assert not ok

    def test_spaces_are_stripped(self):
        ok, _ = validate_iban("PT50 0170 3782 0304 0053 5672 9")
        assert ok

    def test_lowercase_accepted(self):
        # lowercase deve ser normalizado para maiúsculas
        ok, _ = validate_iban("pt50 0170 3782 0304 0053 5672 9")
        assert ok


class TestFormatIban:
    def test_formats_in_groups_of_4(self):
        result = format_iban("PT500170378203040053567 29")
        assert result == "PT50 0170 3782 0304 0053 5672 9"

    def test_already_formatted_stays_same(self):
        formatted = "PT50 0170 3782 0304 0053 5672 9"
        assert format_iban(formatted) == formatted

    def test_empty_returns_original(self):
        assert format_iban("") == ""

    def test_short_iban_grouped(self):
        result = format_iban("NL91ABNA0417164300")
        assert result == "NL91 ABNA 0417 1643 00"


class TestValidateIbanList:
    def test_mixed_list(self):
        ibans = [
            "PT50 0170 3782 0304 0053 5672 9",
            "INVALIDO",
        ]
        results = validate_iban_list(ibans)
        assert len(results) == 2
        assert results[0]['valid'] is True
        assert results[1]['valid'] is False

    def test_each_result_has_formatted(self):
        results = validate_iban_list(["PT50 0170 3782 0304 0053 5672 9"])
        assert 'formatted' in results[0]
        assert results[0]['formatted'] == "PT50 0170 3782 0304 0053 5672 9"
