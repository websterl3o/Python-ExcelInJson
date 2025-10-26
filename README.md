# Python-ExcelInJson

### Tags: python, excel, json, data-transformation, etl, data-cleaning

Um script simples para converter planilhas Excel (`.xlsx/.xls`) em **JSON** limpo, com opções para **normalizar cabeçalhos**, tratar **células vazias**, e **formatar datas** como **ISO-8601** ou **Unix timestamp** (segundos ou milissegundos).

---

## EN — Explanation

`excelInJson.py` turns an Excel sheet into JSON ready for your app. It assumes the **first row** contains **column names** and each subsequent row becomes one **JSON object**.

### Features

* Read `.xlsx/.xls`; pick a **sheet** by index or name.
* **Normalize headers** to `snake_case` (or keep as-is).
* Convert **empty cells** to **`null`** (optional).
* Convert **dates** to **ISO-8601** or **Unix timestamp**:

  * `--date-iso` → e.g., `2025-10-25T14:00:00`
  * `--date-epoch --epoch-unit {s,ms}` → e.g., `1733431800` or `1733431800000`
  * Parsing helpers: `--dayfirst`, `--date-format`, `--date-cols`, `--no-auto-dates`
* Output format:

  * `records` → `[{...}, {...}]` (default, API-friendly)
  * `table` → includes a **schema** (good for validation/ETL)

### Quickstart

```bash
# Simple conversion (sheet 0), ISO dates, empties -> null
python excelInJson.py data.xlsx -s 0 --date-iso --na-null

# Use sheet by name, keep headers as-is, custom output file
python excelInJson.py data.xlsx -s "Sheet1" --keep-headers -o output.json
```

### Date handling examples

```bash
# Dates as ISO, dd/mm/yyyy HH:MM
python excelInJson.py data.xlsx -s 0 --date-iso --dayfirst --date-cols modified --date-format "%d/%m/%Y %H:%M"

# Dates as Unix timestamp in seconds
python excelInJson.py data.xlsx -s 0 --date-epoch --epoch-unit s --dayfirst --date-cols modified --date-format "%d/%m/%Y %H:%M"

# Dates as Unix timestamp in milliseconds
python excelInJson.py data.xlsx -s 0 --date-epoch --epoch-unit ms --dayfirst --date-cols modified --date-format "%d/%m/%Y %H:%M"

# Disable auto-detection and convert only specific columns
python excelInJson.py data.xlsx --date-iso --no-auto-dates --date-cols modified
```

> Notes
> • **Numeric columns are never treated as dates** (prevents `1970-01-01` issues).
> • If you have ZIP/codes with **leading zeros**, keep them as **text** in Excel.
> • `pyarrow` (optional) improves typing, but the script works without it.

### Install

```bash
# Update pip
python -m pip install --upgrade pip
# Dependencies
python -m pip install pandas openpyxl pyarrow
```

### Troubleshooting

* `ModuleNotFoundError: No module named 'pandas'` → install deps (see above).
* Date parser warning (“Could not infer format…”) → pass `--date-format` and/or `--no-auto-dates`.
* Multiple Python installs on Windows? Prefer the launcher:

  ```bash
  py -0p            # list interpreters
  py -3.12 -m pip install pandas openpyxl pyarrow
  py -3.12 excelInJson.py data.xlsx --date-iso
  ```

---

## PT-BR — Explicação

O `excelInJson.py` pega sua planilha do Excel e gera um **JSON** prontinho pra usar. Ele entende que a **primeira linha** tem os **nomes das colunas** e cada linha de baixo vira um **objeto** no JSON.

### O que ele faz

* Lê `.xlsx/.xls` e permite escolher a **aba** (índice ou nome).
* **Normaliza cabeçalhos** para `snake_case` (ou mantém como está).
* Converte **células vazias** em **`null`** (opcional).
* Converte **datas** para **ISO-8601** ou **Unix timestamp**:

  * `--date-iso` → ex.: `2025-10-25T14:00:00`
  * `--date-epoch --epoch-unit {s,ms}` → ex.: `1733431800` ou `1733431800000`
  * Ajuda no parse: `--dayfirst`, `--date-format`, `--date-cols`, `--no-auto-dates`
* Formato de saída:

  * `records` → `[{...}, {...}]` (padrão, ideal pra APIs)
  * `table` → inclui **schema** (bom pra ETL/validação)

### Comece rápido

```bash
# Conversão simples (aba 0), datas ISO e vazios -> null
python excelInJson.py dados.xlsx -s 0 --date-iso --na-null

# Escolher aba pelo nome, manter cabeçalhos e definir arquivo de saída
python excelInJson.py dados.xlsx -s "Plan1" --keep-headers -o saida.json
```

### Exemplos de datas

```bash
# Datas em ISO, formato dd/mm/aaaa HH:MM
python excelInJson.py dados.xlsx -s 0 --date-iso --dayfirst --date-cols modificado --date-format "%d/%m/%Y %H:%M"

# Datas como Unix timestamp em segundos
python excelInJson.py dados.xlsx -s 0 --date-epoch --epoch-unit s --dayfirst --date-cols modificado --date-format "%d/%m/%Y %H:%M"

# Datas como Unix timestamp em milissegundos
python excelInJson.py dados.xlsx -s 0 --date-epoch --epoch-unit ms --dayfirst --date-cols modificado --date-format "%d/%m/%Y %H:%M"

# Desligar detecção automática e converter só colunas específicas
python excelInJson.py dados.xlsx --date-iso --no-auto-dates --date-cols modificado
```

> Observações
> • **Colunas numéricas nunca são tratadas como datas** (evita `1970-01-01`).
> • Se tiver CEP/código com **zero à esquerda**, mantenha como **texto** no Excel.
> • `pyarrow` é opcional: melhora tipos, mas não é obrigatório.

### Instalação

```bash
# 1) Atualize o pip
python -m pip install --upgrade pip
# 2) Dependências
python -m pip install pandas openpyxl pyarrow
```

### Solução de problemas

* `ModuleNotFoundError: No module named 'pandas'` → instale as dependências.
* Aviso de parser de data (“Could not infer format…”) → use `--date-format` e/ou `--no-auto-dates`.
* Vários Pythons no Windows? Use o launcher:

  ```bash
  py -0p
  py -3.12 -m pip install pandas openpyxl pyarrow
  py -3.12 excelInJson.py dados.xlsx --date-iso
  ```

---

## CLI Overview

**Common flags**

* `-s`, `--sheet` — sheet index or name
* `-o`, `--output` — output file name (defaults to input with `.json`)
* `--keep-headers` — keep headers exactly as in Excel
* `--na-null` — convert empty cells to `null`
* `--orient {records,table}` — output shape (default `records`)

**Date flags**

* `--date-iso` — format detected date columns as ISO-8601
* `--date-epoch` — format detected date columns as Unix timestamp
* `--epoch-unit {s,ms}` — timestamp unit (seconds/milliseconds)
* `--dayfirst` — parse as `DD/MM/YYYY` (pt-BR)
* `--date-format "<fmt>"` — explicit strptime format (e.g., `"%d/%m/%Y %H:%M"`)
* `--date-cols "col1,col2"` — force only these columns as dates
* `--no-auto-dates` — disable automatic detection (use with `--date-cols`)
