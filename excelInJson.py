import json, argparse, pandas as pd
from pathlib import Path
import re

def snake(s):  # normaliza cabeçalhos p/ snake_case
    s = re.sub(r'\s+', '_', s.strip())
    s = re.sub(r'[^\w]+', '_', s, flags=re.A)
    s = re.sub(r'_+', '_', s).strip('_')
    return s.lower()

ap = argparse.ArgumentParser()
ap.add_argument("input", help="Arquivo Excel (.xlsx/.xls)")
ap.add_argument("-o","--output", help="Arquivo de saída .json (default = mesmo nome)")
ap.add_argument("-s","--sheet", help="Nome da planilha ou índice (0-based)")
ap.add_argument("--orient", default="records", choices=["records","table"],
                help="Formato JSON (records=lista de objetos, table=JSON Table Schema)")
ap.add_argument("--keep-headers", action="store_true", help="Não normalizar cabeçalhos")
ap.add_argument("--na-null", action="store_true", help="Converter células vazias para null")
ap.add_argument("--date-iso", action="store_true", help="Forçar datas para ISO8601 (UTC-naive)")
args = ap.parse_args()

sheet = None
if args.sheet is not None:
    try: sheet = int(args.sheet)
    except: sheet = args.sheet  # aceita nome ou índice

df = pd.read_excel(args.input, sheet_name=sheet, dtype_backend="pyarrow")

# normaliza cabeçalhos
if not args.keep-headers:
    df.columns = [snake(str(c)) for c in df.columns]

# trata NaN -> None
if args.na_null:
    df = df.where(pd.notna(df), None)

# datas em ISO 8601
if args.date_iso:
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%dT%H:%M:%S")

out = args.output or (Path(args.input).with_suffix(".json"))
# orient="records" -> [{...},{...}]
# orient="table"   -> {"schema":...,"data":[...]} (bom p/ validação)
with open(out, "w", encoding="utf-8") as f:
    json.dump(
        df.to_dict(orient="records") if args.orient=="records" else json.loads(df.to_json(orient="table")),
        f, ensure_ascii=False, indent=2
    )
print(f"Gerado: {out}")
