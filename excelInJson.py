# excelInJson.py
import json
import argparse
from pathlib import Path
import re
import pandas as pd

def snake(s: str) -> str:
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"[^\w]+", "_", s, flags=re.A)
    s = re.sub(r"_+", "_", s).strip("_")
    return s.lower()

ap = argparse.ArgumentParser()
ap.add_argument("input", help="Arquivo Excel (.xlsx/.xls)")
ap.add_argument("-o", "--output", help="Arquivo de saída .json (default = mesmo nome)")
ap.add_argument("-s", "--sheet", help="Nome da planilha ou índice (0-based)")
ap.add_argument(
    "--orient",
    default="records",
    choices=["records", "table"],
    help="Formato JSON (records=lista de objetos, table=JSON Table Schema)",
)
ap.add_argument(
    "--keep-headers",
    dest="keep_headers",
    action="store_true",
    help="Não normalizar cabeçalhos (mantém como está no Excel)",
)
ap.add_argument(
    "--na-null",
    dest="na_null",
    action="store_true",
    help="Converter células vazias para null",
)
ap.add_argument(
    "--date-iso",
    dest="date_iso",
    action="store_true",
    help="Forçar datas para ISO8601 (UTC-naive)",
)
args = ap.parse_args()

# sheet pode ser índice (int) ou nome (str)
sheet = None
if args.sheet is not None:
    try:
        sheet = int(args.sheet)
    except ValueError:
        sheet = args.sheet

# tenta usar dtype_backend='pyarrow' se disponível
read_excel_kwargs = {"sheet_name": sheet}
try:
    import pyarrow  # noqa: F401
    read_excel_kwargs["dtype_backend"] = "pyarrow"
except Exception:
    pass  # segue sem pyarrow

df = pd.read_excel(args.input, **read_excel_kwargs)

# normaliza cabeçalhos (se não pediu para manter)
if not args.keep_headers:
    df.columns = [snake(str(c)) for c in df.columns]

# trata NaN -> None
if args.na_null:
    df = df.where(pd.notna(df), None)

# datas em ISO 8601
if args.date_iso:
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%dT%H:%M:%S")

out_path = Path(args.output) if args.output else Path(args.input).with_suffix(".json")

# gera JSON
if args.orient == "records":
    payload = df.to_dict(orient="records")
else:  # "table"
    payload = json.loads(df.to_json(orient="table"))

with open(out_path, "w", encoding="utf-8") as f:
    json.dump(payload, f, ensure_ascii=False, indent=2)

print(f"Gerado: {out_path}")
