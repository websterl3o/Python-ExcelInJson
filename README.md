# Python-ExcelInJson

### Tags: python, excel, json, data-transformation, etl, data-cleaning

# EN — Explanation

`excelInJson.py` is your handy little script that takes an Excel sheet and spits out clean JSON ready for your app. It assumes the **first row** is your **column names** and every next row becomes a **JSON object**.

What it does:

* Reads `.xlsx/.xls` and lets you pick the **sheet**.
* **Normalizes headers** to `snake_case` so your backend stays happy.
* Converts **dates** to **ISO-8601** (e.g., `2025-10-25T14:00:00`) if you want.
* Turns **empty cells** into **`null`** (super helpful for validation).
* You pick the **output format**:

  * `records` → `[{...}, {...}]` (API-friendly)
  * `table` → includes a **schema** (great for validation/ETL)

How to run:

```bash
# Simple conversion (sheet 0), ISO dates, and empties -> null
python excelInJson.py data.xlsx -s 0 --date-iso --na-null

# Use sheet by name, keep headers as-is
python excelInJson.py data.xlsx -s "Sheet1" --keep-headers -o output.json
```

When it’s handy:

* Seeding your database or migrations.
* Posting data to an API without wrestling with Excel.
* Standardizing a “messy” client dataset.

Pro tip:

* If you’ve got codes/ZIPs with **leading zeros**, keep them as **text** in Excel so you don’t lose formatting.

If you hit errors like `ModuleNotFoundError: No module named 'pandas'`, install the dependencies with:

```bash
Em caso de erros como `ModuleNotFoundError: No module named 'pandas'`, instale as dependências com:

```bash
# 1) Update pip
python -m pip install --upgrade pip

# 2) Install script dependencies
python -m pip install pandas openpyxl pyarrow
```

---

# PT-BR — Explicação

O `excelInJson.py` é aquele script amigo que pega sua planilha do Excel e transforma em um JSON prontinho pra usar na sua aplicação. Ele entende que a **primeira linha** da planilha são os **nomes das colunas** (os “campos”) e cada linha de baixo vira um **objeto** no JSON.

O que ele faz por você:

* Lê `.xlsx/.xls` e escolhe a **aba** que você quiser.
* **Normaliza os cabeçalhos**: vira `snake_case` pra não dar ruim no backend.
* **Datas** viram **ISO-8601** direitinho (`2025-10-25T14:00:00`) se você pedir.
* **Células vazias** podem virar **`null`** (ótimo pra validação).
* Você escolhe o **formato**:

  * `records` → `[{...}, {...}]` (o comum das APIs)
  * `table` → inclui **schema** (bom pra validação/ETL)

Como usar:

```bash
# Conversão simples (aba 0), datas ISO e vazios -> null
python excelInJson.py dados.xlsx -s 0 --date-iso --na-null

# Escolhendo aba pelo nome, mantendo cabeçalhos como estão
python excelInJson.py dados.xlsx -s "Plan1" --keep-headers -o saida.json
```

Quando usar:

* Quer popular o banco via seed/migração.
* Precisa subir dados via API sem sofrer com Excel.
* Vai padronizar um dataset “meio torto” vindo do cliente.

Dica rápida:

* Se tiver código/CEP com **zero à esquerda**, trate como **texto** no Excel pra não perder o formato.

Em caso de erros como `ModuleNotFoundError: No module named 'pandas'`, instale as dependências com:

```bash
# 1) Atualize o pip
python -m pip install --upgrade pip

# 2) Instale dependências do script
python -m pip install pandas openpyxl pyarrow
```