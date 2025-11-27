# Extrator de Código VBA de Arquivos Excel

Script Python para extrair automaticamente código VBA (macros) de arquivos Excel e salvar em arquivos separados (.bas, .cls, .frm).

## 📂 Localização

**Pasta de trabalho:** `C:\Users\seuUsuario\Documents\excel\`

Coloque seus arquivos Excel (.xlsm, .xls, .xlam) nesta pasta antes de executar o script.

## 📋 Requisitos

- Python 3.7 ou superior
- Biblioteca `oletools`

## 🚀 Instalação

1. Abra o PowerShell nesta pasta:

```powershell
cd C:\Users\seuUsuario\Documents\excel
```

2. Instale as dependências:

```powershell
pip install -r requirements.txt
```

## 💻 Como Usar

### Uso Básico

1. **Coloque seu arquivo Excel nesta pasta**
2. Execute o script:

```powershell
python extrair_vba.py seu_arquivo.xlsm
```

### Exemplos

```powershell
# Extrair macros de planilha.xlsm
python extrair_vba.py planilha.xlsm

# Especificar pasta de saída
python extrair_vba.py planilha.xlsm ./vba_extracted
```

## 📁 Estrutura de Saída

- **.bas** - Módulos padrão
- **.cls** - Módulos de classe
- **.frm** - UserForms

## ⚠️ Formatos Suportados

- .xlsm - Excel com macros (2007+)
- .xls - Excel antigo (97-2003)
- .xlam - Excel Add-in com macros
