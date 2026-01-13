# 📄 Conversor Excel → PDF

Aplicação para converter ficheiros Excel em documentos PDF formatados.

## Como Usar ?

### Opção 1: Executável (Recomendado)
1. Abra a pasta `dist/`
2. Execute o ficheiro **ConversorExcelPDF**
3. Use a interface gráfica para:
   - Selecionar o ficheiro Excel
   - Personalizar as opções do PDF
   - Converter

### Opção 2: Python (Desenvolvimento)
```bash
# Ativar ambiente virtual
source venv/bin/activate

# Executar com interface gráfica
python converter_excel_pdf.py

# Ou converter via linha de comandos
python converter_excel_pdf.py ficheiro.xlsx
```

---

## Estrutura do Excel

O ficheiro Excel deve ter idealmente duas folhas:

### Folha "Configuracao" (opcional)
| Campo | Valor |
|-------|-------|
| nome_empresa | Nome da Empresa, Lda |
| morada_empresa | Rua X, 123 - 4000-000 Porto |
| telefone_empresa | +351 220 000 000 |
| email_empresa | geral@empresa.pt |
| nif_empresa | PT 500 000 000 |
| nome_cliente | Cliente XYZ |
| morada_cliente | Morada do Cliente |
| nif_cliente | 500 000 000 |
| tipo_documento | GUIA DE REMESSA |
| numero_documento | GR2024001 |
| data_documento | 18/12/2024 |
| observacoes | Texto de observações... |

### Folha "Itens" (obrigatória)
| Código | Designação | Quantidade | Preço Unit. | Total |
|--------|------------|------------|-------------|-------|
| REF-001 | Filtro de Óleo | 5 | 12.50 | 62.50 |
| REF-002 | Pastilhas Travão | 3 | 45.00 | 135.00 |

---

## Opções de Personalização

A aplicação permite personalizar:

### Página PDF
- Tamanho: A4, A3, Letter
- Orientação: Retrato, Paisagem
- Margens: Superior, Inferior, Esquerda, Direita

### Cabeçalho
- Dados da empresa (nome, morada, telefone, email, NIF)
- Logo personalizado

### Cores
- Cor do cabeçalho da tabela
- Cor do texto
- Cor das linhas alternadas
- Cor das bordas

### Tabela
- Tamanho da fonte
- Padding das células
- Mostrar/ocultar grelha
- Cores alternadas nas linhas

### Rodapé
- Área de assinaturas
- Data de geração
- Texto personalizado

### Dados Bancários
- Nome do banco
- IBAN
- Mostrar/ocultar no PDF

> Os dados bancários aparecem no rodapé do PDF, substituindo os campos "Verificado por" e "Data".

---

## 📁 Ficheiros do Projeto

```
Script_conversor/
├── dist/
│   └── ConversorExcelPDF     # ← Executável
├── converter_excel_pdf.py    # Script principal
├── criar_excel_exemplo.py    # Gerador de dados de teste
├── dados_exemplo.xlsx        # Excel de exemplo
├── requirements.txt          # Dependências Python
├── config.json              # Configurações guardadas
└── venv/                    # Ambiente virtual Python
```

---

## Instalação (Desenvolvimento)

```bash
# Criar ambiente virtual
python3 -m venv venv

# Ativar
source venv/bin/activate

# Instalar dependências
pip install -r requirements.txt

# Executar
python converter_excel_pdf.py
```

---

## Criar Novo Executável

```bash
source venv/bin/activate
pip install pyinstaller
pyinstaller --onefile --windowed --name "ConversorExcelPDF" converter_excel_pdf.py
```

O executável será criado em `dist/ConversorExcelPDF`.

---

## Notas

- As configurações são **persistentes** e guardadas em `config.json`
- O cabeçalho e dados bancários são mantidos entre sessões
- O PDF é gerado na mesma pasta do Excel por defeito
- O executável funciona sem necessidade de Python instalado
