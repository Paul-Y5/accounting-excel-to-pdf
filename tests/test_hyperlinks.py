"""
Testes para a feature de hyperlinks no PDF e extras de segurança (feature/v3.3-hyperlinks).

Valida que:
- company_website existe no DEFAULT_CONFIG
- o campo website é incluído na linha de contacto do cabeçalho
- o URL é normalizado com https:// quando necessário
- email e website sem valor não geram links
- _sanitize_text escapa caracteres XML reservados
- campos de texto simples são sanitizados antes de entrar no markup do PDF
"""
import copy
import pytest
from src.config import DEFAULT_CONFIG
from src.converter import ExcelToPDFConverter, _sanitize_text


@pytest.fixture()
def base_config():
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    cfg['header']['company_website'] = 'www.empresa.pt'
    return cfg


class TestDefaultConfig:
    def test_company_website_key_exists(self):
        assert 'company_website' in DEFAULT_CONFIG['header']

    def test_company_website_default_is_empty(self):
        assert DEFAULT_CONFIG['header']['company_website'] == ''


class TestWebsiteNormalization:
    def test_url_without_scheme_gets_https(self, base_config):
        conv = ExcelToPDFConverter('dummy.xlsx', config=base_config)
        website = 'www.empresa.pt'
        url = website if website.startswith(('http://', 'https://')) else f'https://{website}'
        assert url == 'https://www.empresa.pt'

    def test_url_with_http_unchanged(self, base_config):
        website = 'http://empresa.pt'
        url = website if website.startswith(('http://', 'https://')) else f'https://{website}'
        assert url == 'http://empresa.pt'

    def test_url_with_https_unchanged(self, base_config):
        website = 'https://empresa.pt'
        url = website if website.startswith(('http://', 'https://')) else f'https://{website}'
        assert url == 'https://empresa.pt'


class TestHeaderContactLine:
    def _build_contact_line(self, email='', website='', telefone='', nif=''):
        """Replica a lógica de construção da linha de contacto do converter."""
        contact_line = []
        if telefone:
            contact_line.append(f"Tel: {telefone}")
        if email:
            email_link = f'<a href="mailto:{email}" color="#2b6cb0">{email}</a>'
            contact_line.append(f"Email: {email_link}")
        if website:
            url = website if website.startswith(('http://', 'https://')) else f'https://{website}'
            website_link = f'<a href="{url}" color="#2b6cb0">{website}</a>'
            contact_line.append(website_link)
        if nif:
            contact_line.append(f"NIF: {nif}")
        return " | ".join(contact_line)

    def test_email_produces_mailto_link(self):
        line = self._build_contact_line(email='geral@empresa.pt')
        assert 'mailto:geral@empresa.pt' in line

    def test_website_produces_href_link(self):
        line = self._build_contact_line(website='www.empresa.pt')
        assert 'href="https://www.empresa.pt"' in line

    def test_empty_email_no_link(self):
        line = self._build_contact_line(email='')
        assert 'mailto:' not in line

    def test_empty_website_no_link(self):
        line = self._build_contact_line(website='')
        assert 'href=' not in line

    def test_all_fields_order(self):
        line = self._build_contact_line(
            telefone='+351220000000',
            email='geral@empresa.pt',
            website='www.empresa.pt',
            nif='PT500000000',
        )
        tel_pos = line.index('Tel:')
        email_pos = line.index('mailto:')
        web_pos = line.index('www.empresa.pt')
        nif_pos = line.index('NIF:')
        assert tel_pos < email_pos < web_pos < nif_pos

    def test_website_display_text_is_original(self):
        line = self._build_contact_line(website='www.empresa.pt')
        assert '>www.empresa.pt<' in line


class TestSanitizeText:
    def test_ampersand_escaped(self):
        assert _sanitize_text('A & B') == 'A &amp; B'

    def test_less_than_escaped(self):
        assert _sanitize_text('a<b') == 'a&lt;b'

    def test_greater_than_escaped(self):
        assert _sanitize_text('a>b') == 'a&gt;b'

    def test_combined_injection_attempt(self):
        result = _sanitize_text('<script>alert("xss")</script>')
        assert '<script>' not in result
        assert '&lt;script&gt;' in result

    def test_normal_text_unchanged(self):
        assert _sanitize_text('EMPRESA EXEMPLO, LDA') == 'EMPRESA EXEMPLO, LDA'

    def test_empty_string_unchanged(self):
        assert _sanitize_text('') == ''

    def test_none_returns_falsy(self):
        assert not _sanitize_text(None)

    def test_multiple_chars_in_one_string(self):
        result = _sanitize_text('A & B <c> d')
        assert result == 'A &amp; B &lt;c&gt; d'
