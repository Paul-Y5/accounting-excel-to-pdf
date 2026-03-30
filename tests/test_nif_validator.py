"""
Testes unitários para o módulo de validação de NIF.
"""

import pytest

from src.nif_validator import validate_nif, validate_nif_list


class TestValidateNif:
    """Testes para a função validate_nif."""

    def test_valid_nif(self):
        """Verifica que um NIF válido é aceite."""
        is_valid, msg = validate_nif('123456789')
        assert is_valid is True
        assert msg == "NIF válido"

    def test_empty_nif(self):
        """Verifica que NIF vazio é rejeitado."""
        is_valid, msg = validate_nif('')
        assert is_valid is False
        assert "vazio" in msg

    def test_nif_with_spaces(self):
        """Verifica que espaços são removidos."""
        is_valid, _ = validate_nif('123 456 789')
        assert is_valid is True

    def test_nif_with_pt_prefix(self):
        """Verifica que prefixo PT é removido."""
        is_valid, _ = validate_nif('PT123456789')
        assert is_valid is True

    def test_nif_with_pt_prefix_and_spaces(self):
        """Verifica que PT com espaços é tratado."""
        is_valid, _ = validate_nif('PT 123 456 789')
        assert is_valid is True

    def test_nif_with_dots_and_hyphens(self):
        """Verifica que pontos e hífens são removidos."""
        is_valid, _ = validate_nif('123.456.789')
        assert is_valid is True

    def test_nif_with_letters(self):
        """Verifica que NIF com letras (não PT) é rejeitado."""
        is_valid, msg = validate_nif('12345678A')
        assert is_valid is False
        assert "caracteres inválidos" in msg

    def test_nif_too_short(self):
        """Verifica que NIF com menos de 9 dígitos é rejeitado."""
        is_valid, msg = validate_nif('12345678')
        assert is_valid is False
        assert "9 dígitos" in msg

    def test_nif_too_long(self):
        """Verifica que NIF com mais de 9 dígitos é rejeitado."""
        is_valid, msg = validate_nif('1234567890')
        assert is_valid is False
        assert "9 dígitos" in msg

    def test_nif_invalid_first_digit(self):
        """Verifica que NIF com primeiro dígito 0 ou 4 é rejeitado."""
        is_valid, msg = validate_nif('012345678')
        assert is_valid is False
        assert "primeiro dígito inválido" in msg

        is_valid, msg = validate_nif('412345678')
        assert is_valid is False
        assert "primeiro dígito inválido" in msg

    def test_nif_invalid_check_digit(self):
        """Verifica que NIF com dígito de controlo errado é rejeitado."""
        is_valid, msg = validate_nif('123456780')
        assert is_valid is False
        assert "dígito de controlo" in msg

    @pytest.mark.parametrize("first_digit", [1, 2, 3, 5, 6, 7, 8, 9])
    def test_valid_first_digits_accepted(self, first_digit):
        """Verifica que todos os primeiros dígitos válidos são aceites (quando o NIF completo é válido)."""
        # Construir um NIF válido com o algoritmo módulo 11
        base = f"{first_digit}00000000"[:8]
        weights = [9, 8, 7, 6, 5, 4, 3, 2]
        total = sum(int(base[i]) * weights[i] for i in range(8))
        remainder = total % 11
        check = 0 if remainder < 2 else 11 - remainder
        nif = base + str(check)

        is_valid, _ = validate_nif(nif)
        assert is_valid is True


class TestValidateNifList:
    """Testes para a função validate_nif_list."""

    def test_empty_list(self):
        """Verifica que lista vazia retorna lista vazia."""
        results = validate_nif_list([])
        assert results == []

    def test_single_valid_nif(self):
        """Verifica resultado para um NIF válido."""
        results = validate_nif_list(['123456789'])
        assert len(results) == 1
        assert results[0]['nif'] == '123456789'
        assert results[0]['valid'] is True

    def test_single_invalid_nif(self):
        """Verifica resultado para um NIF inválido."""
        results = validate_nif_list(['000000000'])
        assert len(results) == 1
        assert results[0]['valid'] is False

    def test_mixed_list(self):
        """Verifica lista com NIFs válidos e inválidos."""
        results = validate_nif_list(['123456789', '', '000000000'])
        assert len(results) == 3
        assert results[0]['valid'] is True
        assert results[1]['valid'] is False
        assert results[2]['valid'] is False

    def test_result_structure(self):
        """Verifica que cada resultado tem as chaves esperadas."""
        results = validate_nif_list(['123456789'])
        result = results[0]
        assert 'nif' in result
        assert 'valid' in result
        assert 'message' in result
