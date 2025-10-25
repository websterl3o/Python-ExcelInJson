# excelInJson.py
import argparse
import json
import re
from pathlib import Path
import datetime as dt

import pandas as pd
try:
    import numpy as np  # opcional (apenas para serialização segura)
except Exception:
    np = None


def snake(s: str) -> str:
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"[^\w]+", "_", s, flags=re.A)
    s = re.sub(r"_+", "_", s).strip("_")
    return s.lower()


def json_default(o):
    # pandas / datetime
    if isinstance(o, (pd.Timestamp, dt.datetime)):
        # formata sempre como ISO sem timezone
        try:
            # se tiver timezone, converte para naive
            if getattr(o, "tzinfo", None) is not None:
                o = o.astimezone(dt.timezone.utc).replace(tzinfo=None)
        except Exception:
            pass
        return o.strftime("%Y-%m-%dT%H:%M:%S")
    if isinstance(o, dt.date):
        return o.isoformat()

    # numpy (se importado)
    if np is not None:
        if isinstance(o, np.integer):
            return int(o)
        if isinstance(o, np.floating):
            return float(o)
        if isinstance(o, np.bool_):
            return bool(o)

    # fallback genérico
    return str(o)


def coerce_datetimes_to_iso(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converte colunas que são (ou podem virar) datetime em strings ISO-8601.
    Cobre dtypes datetime64, séries de Timestamp e colunas 'object' com datas.
    Remove timezone (ou converte e remove) para JSON simples.
    """
    for col in df.columns:
        # tenta converter sem explodir se não for datetime
        converted = pd.to_datetime(df[col], errors="ignore", utc=False)

        # se ficou datetime, formata ISO
        if pd.api.types.is_datetime64_any_dtype(converted):
            # remove timezone se houver
            try:
                converted = converted.dt.tz_convert("UTC").dt.tz_localize(None)
            except Exception:
                try:
                    converted = converted.dt.tz_localize(None)
                except Exception:
                    pass
            df[col] = converted.dt.strftime("%Y-%m-%dT%H:%M:%S")
    return df


def main():
    ap = argparse.ArgumentParser(
        description="Converte Excel (.xlsx/.xls) para JSON."
    )
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
        help="Não normalizar cabeçalhos (mantém como no Excel)",
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
        help="Forçar datas para ISO8601 (remove timezone)",
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

    # lê Excel
    df = pd.read_excel(args.input, **read_excel_kwargs)

    # normaliza cabeçalhos
    if not args.keep_headers:
        df.columns = [snake(str(c)) for c in df.columns]

    # vazios -> None
    if args.na_null:
        df = df.where(pd.notna(df), None)

    # datas -> ISO 8601
    if args.date_iso:
        df = coerce_datetimes_to_iso(df)

    # prepara saída
    out_path = Path(args.output) if args.output else Path(args.input).with_suffix(".json")

    # payload
    if args.orient == "records":
        payload = df.to_dict(orient="records")
    else:  # "table"
        # df.to_json(orient="table") já lida com schema/data,
        # mas pode conter Timestamps -> usa loads/dumps com default
        payload = json.loads(df.to_json(orient="table"))

    # grava JSON (com default para qualquer tipo “estranho”)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=json_default)

    print(f"Gerado: {out_path}")


if __name__ == "__main__":
    main()
