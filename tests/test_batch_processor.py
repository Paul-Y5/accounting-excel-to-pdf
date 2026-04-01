"""
Testes unitários para o módulo de processamento em lote.
"""

import os
import pytest
from unittest.mock import patch, MagicMock

from src.batch_processor import find_excel_files, process_batch


class TestFindExcelFiles:
    """Testes para find_excel_files."""

    def test_finds_xlsx_files(self, tmp_path):
        """Encontra ficheiros .xlsx."""
        (tmp_path / 'a.xlsx').touch()
        (tmp_path / 'b.xlsx').touch()
        files = find_excel_files(str(tmp_path))
        assert len(files) == 2

    def test_finds_xls_files(self, tmp_path):
        """Encontra ficheiros .xls antigos."""
        (tmp_path / 'old.xls').touch()
        files = find_excel_files(str(tmp_path))
        assert len(files) == 1

    def test_ignores_non_excel(self, tmp_path):
        """Ignora PDF, TXT e outros formatos."""
        (tmp_path / 'a.xlsx').touch()
        (tmp_path / 'b.pdf').touch()
        (tmp_path / 'c.txt').touch()
        files = find_excel_files(str(tmp_path))
        assert len(files) == 1

    def test_ignores_temp_files(self, tmp_path):
        """Ignora ficheiros temporários do Excel (~$)."""
        (tmp_path / '~$temp.xlsx').touch()
        (tmp_path / 'real.xlsx').touch()
        files = find_excel_files(str(tmp_path))
        assert len(files) == 1
        assert not any('~$' in f for f in files)

    def test_returns_sorted_by_name(self, tmp_path):
        """Retorna ficheiros ordenados pelo nome."""
        (tmp_path / 'c.xlsx').touch()
        (tmp_path / 'a.xlsx').touch()
        (tmp_path / 'b.xlsx').touch()
        files = find_excel_files(str(tmp_path))
        names = [os.path.basename(f) for f in files]
        assert names == sorted(names)

    def test_returns_absolute_paths(self, tmp_path):
        """Retorna caminhos absolutos."""
        (tmp_path / 'test.xlsx').touch()
        files = find_excel_files(str(tmp_path))
        assert os.path.isabs(files[0])

    def test_empty_folder_returns_empty(self, tmp_path):
        """Pasta vazia retorna lista vazia."""
        files = find_excel_files(str(tmp_path))
        assert files == []

    def test_invalid_folder_raises(self):
        """Pasta inexistente lança ValueError."""
        with pytest.raises(ValueError, match="Pasta não encontrada"):
            find_excel_files('/caminho/que/nao/existe/mesmo')

    def test_case_insensitive_extension(self, tmp_path):
        """Extensão em maiúsculas também é detetada."""
        (tmp_path / 'DADOS.XLSX').touch()
        files = find_excel_files(str(tmp_path))
        assert len(files) == 1


class TestProcessBatch:
    """Testes para process_batch."""

    def _mock_converter(self, tmp_path, clients=2, mode='aggregate'):
        """Cria mock do ExcelToPDFConverter."""
        mock = MagicMock()
        mock.read_excel_data.return_value = {
            'itens': [{'Cliente': f'C{i}'} for i in range(clients)]
        }
        mock.generate_pdf.return_value = str(tmp_path / 'out.pdf')
        mock.generate_individual_pdfs.return_value = [
            str(tmp_path / f'c{i}.pdf') for i in range(clients)
        ]
        return mock

    def test_empty_folder_returns_empty(self, tmp_path):
        """Pasta sem Excel retorna lista vazia."""
        results = process_batch(str(tmp_path), {})
        assert results == []

    def test_result_structure(self, tmp_path):
        """Cada resultado tem as chaves esperadas."""
        (tmp_path / 'test.xlsx').touch()
        with patch('src.batch_processor.ExcelToPDFConverter') as MockConv:
            MockConv.return_value = self._mock_converter(tmp_path)
            results = process_batch(str(tmp_path), {}, mode='aggregate')

        assert len(results) == 1
        r = results[0]
        assert {'file', 'filename', 'success', 'output_path', 'clients_count', 'error'} == set(r.keys())

    def test_successful_aggregate(self, tmp_path):
        """Modo aggregate regista sucesso e caminho do PDF."""
        (tmp_path / 'jan.xlsx').touch()
        with patch('src.batch_processor.ExcelToPDFConverter') as MockConv:
            MockConv.return_value = self._mock_converter(tmp_path, clients=3)
            results = process_batch(str(tmp_path), {}, mode='aggregate')

        assert results[0]['success'] is True
        assert results[0]['clients_count'] == 3
        assert results[0]['error'] == ''

    def test_successful_individual(self, tmp_path):
        """Modo individual regista sucesso."""
        (tmp_path / 'jan.xlsx').touch()
        with patch('src.batch_processor.ExcelToPDFConverter') as MockConv:
            MockConv.return_value = self._mock_converter(tmp_path, clients=2, mode='individual')
            results = process_batch(str(tmp_path), {}, mode='individual')

        assert results[0]['success'] is True

    def test_failed_file_recorded(self, tmp_path):
        """Ficheiro que falha é registado com success=False e mensagem de erro."""
        (tmp_path / 'bad.xlsx').touch()
        with patch('src.batch_processor.ExcelToPDFConverter') as MockConv:
            MockConv.return_value.read_excel_data.side_effect = Exception("Ficheiro corrompido")
            results = process_batch(str(tmp_path), {}, mode='aggregate')

        assert results[0]['success'] is False
        assert 'Ficheiro corrompido' in results[0]['error']
        assert results[0]['clients_count'] == 0

    def test_one_failure_does_not_stop_others(self, tmp_path):
        """Falha num ficheiro não interrompe os restantes."""
        (tmp_path / 'a.xlsx').touch()
        (tmp_path / 'b.xlsx').touch()

        call_count = 0

        def side_effect(*args, **kwargs):
            nonlocal call_count
            call_count += 1
            mock = MagicMock()
            if call_count == 1:
                mock.read_excel_data.side_effect = Exception("Erro")
            else:
                mock.read_excel_data.return_value = {'itens': []}
                mock.generate_pdf.return_value = str(tmp_path / 'b.pdf')
            return mock

        with patch('src.batch_processor.ExcelToPDFConverter', side_effect=side_effect):
            results = process_batch(str(tmp_path), {}, mode='aggregate')

        assert len(results) == 2
        assert results[0]['success'] is False
        assert results[1]['success'] is True

    def test_progress_callback_called_per_file(self, tmp_path):
        """Progress callback é chamado antes e depois de cada ficheiro."""
        (tmp_path / 'a.xlsx').touch()
        (tmp_path / 'b.xlsx').touch()
        calls = []

        with patch('src.batch_processor.ExcelToPDFConverter') as MockConv:
            MockConv.return_value = self._mock_converter(tmp_path)
            process_batch(str(tmp_path), {}, mode='aggregate',
                          progress_callback=lambda cur, total, fname: calls.append((cur, total, fname)))

        # 2 ficheiros × 2 chamadas (antes + depois) = 4 chamadas
        assert len(calls) == 4

    def test_progress_callback_final_state(self, tmp_path):
        """Última chamada do callback tem current == total."""
        (tmp_path / 'a.xlsx').touch()
        calls = []

        with patch('src.batch_processor.ExcelToPDFConverter') as MockConv:
            MockConv.return_value = self._mock_converter(tmp_path)
            process_batch(str(tmp_path), {}, mode='aggregate',
                          progress_callback=lambda cur, total, fname: calls.append((cur, total)))

        assert calls[-1][0] == calls[-1][1]

    def test_multiple_files_all_processed(self, tmp_path):
        """Todos os ficheiros da pasta são processados."""
        for name in ['jan.xlsx', 'fev.xlsx', 'mar.xlsx']:
            (tmp_path / name).touch()

        with patch('src.batch_processor.ExcelToPDFConverter') as MockConv:
            MockConv.return_value = self._mock_converter(tmp_path)
            results = process_batch(str(tmp_path), {}, mode='aggregate')

        assert len(results) == 3
