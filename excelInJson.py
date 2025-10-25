# excelInJson.py
import argparse
import json
import re
from pathlib import Path
import datetime as dt

import pandas as pd
try:
    import numpy as np  # opcional (para serialização segura)
except Exception:
    np = None


def snake(s: str) -> str:
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"[^\w]+", "_", s, flags=re.A)
    s = re.sub(r"_+", "_", s).strip("_")
    return s.lower()


def json_default(o):
    if isinstance(o, (pd.Timestamp, dt.datetime)):
        try:
            if getattr(o, "tzinfo", None) is not None:
                o = o.astimezone(dt.timezone.utc).replace(tzinfo=None)
        except Exception:
            pass
        return o.strftime("%Y-%m-%dT%H:%M:%S")
    if isinstance(o, dt.date):
        return o.isoformat()

    if np is not None:
        if isinstance(o, np.integer):  return int(o)
        if isinstance(o, np.floating): return float(o)
        if isinstance(o, np.bool_):    return bool(o)

    return str(o)


def format_dt_series_iso(s: pd.Series) -> pd.Series:
    """Formata uma Series datetime (com ou sem timezone) para string ISO."""
    try:
        s = s.dt.tz_convert("UTC").dt.tz_localize(None)
    except Exception:
        try:
            s = s.dt.tz_localize(None)
        except Exception:
            pass
    return s.dt.strftime("%Y-%m-%dT%H:%M:%S")


def coerce_datetimes_to_iso(
    df: pd.DataFrame,
    *,
    dayfirst: bool,
    auto_dates: bool,
    forced_cols: set[str],
    threshold: float = 0.8,
) -> pd.DataFrame:
    """
    - Converte colunas já datetime64 -> ISO.
    - Para colunas texto/object, tenta parsear; só converte se >= threshold parecem datas.
    - Nunca converte colunas numéricas.
    - 'forced_cols' sempre converte, mesmo que não atinja o limiar.
    """
    for col in df.columns:
        s = df[col]

        # 1) Já é datetime? Formata.
        if pd.api.types.is_datetime64_any_dtype(s):
            df[col] = format_dt_series_iso(s)
            continue

        # 2) Forçar por CLI (apenas tenta se não for numérica)
        if col in forced_cols and not (pd.api.types.is_integer_dtype(s) or pd.api.types.is_float_dtype(s)):
            parsed = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst, utc=False)
            if parsed.notna().any():
                df[col] = format_dt_series_iso(parsed)
            continue

        # 3) Heurística automática (se habilitada) — apenas em texto/object
        if auto_dates and (pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s)):
            parsed = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst, utc=False)
            non_null = s.notna().sum()
            ratio = (parsed.notna().sum() / non_null) if non_null else 0.0
            if ratio >= threshold and parsed.notna().any():
                df[col] = format_dt_series_iso(parsed)

        # 4) Se for numérica → ignora (não é data)
        # (nada a fazer)

    return df


def main():
    ap = argparse.ArgumentParser(description="Converte Excel (.xlsx/.xls) para JSON.")
    ap.add_argument("input", help="Arquivo Excel (.xlsx/.xls)")
    ap.add_argument("-o", "--output", help="Arquivo de saída .json (default = mesmo nome)")
    ap.add_argument("-s", "--sheet", help="Nome da planilha ou índice (0-based)")
    ap.add_argument("--orient", default="records", choices=["records", "table"],
                    help="Formato JSON (records=lista de objetos, table=JSON Table Schema)")
    ap.add_argument("--keep-headers", dest="keep_headers", action="store_true",
                    help="Não normalizar cabeçalhos (mantém como no Excel)")
    ap.add_argument("--na-null", dest="na_null", action="store_true",
                    help="Converter células vazias para null")
    ap.add_argument("--date-iso", dest="date_iso", action="store_true",
                    help="Formatar colunas de data em ISO8601")
    ap.add_argument("--dayfirst", dest="dayfirst", action="store_true",
                    help="Interpretar datas como DD/MM/AAAA (pt-BR)")
    ap.add_argument("--date-cols", dest="date_cols", default="",
                    help="Lista de colunas (separadas por vírgula) para forçar como data")
    ap.add_argument("--no-auto-dates", dest="no_auto_dates", action="store_true",
                    help="Desabilita a detecção automática de colunas de data")
    args = ap.parse_args()

    # Interpretar sheet como índice ou nome
    sheet = None
    if args.sheet is not None:
        try:
            sheet = int(args.sheet)
        except ValueError:
            sheet = args.sheet

    # dtype_backend='pyarrow' se disponível
    read_excel_kwargs = {"sheet_name": sheet}
    try:
        import pyarrow  # noqa: F401
        read_excel_kwargs["dtype_backend"] = "pyarrow"
    except Exception:
        pass

    # Lê Excel
    df = pd.read_excel(args.input, **read_excel_kwargs)

    # Normaliza cabeçalhos
    if not args.keep_headers:
        df.columns = [snake(str(c)) for c in df.columns]

    # Vazios -> None
    if args.na_null:
        df = df.where(pd.notna(df), None)

    # Datas -> ISO (robusto)
    if args.date_iso:
        forced_cols = {c.strip() for c in args.date_cols.split(",") if c.strip()}
        df = coerce_datetimes_to_iso(
            df,
            dayfirst=args.dayfirst,
            auto_dates=not args.no_auto_dates,
            forced_cols=forced_cols,
        )

    # Saída
    out_path = Path(args.output) if args.output else Path(args.input).with_suffix(".json")

    # Payload
    if args.orient == "records":
        payload = df.to_dict(orient="records")
    else:
        payload = json.loads(df.to_json(orient="table"))

    # Grava JSON
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=json_default)

    print(f"Gerado: {out_path}")


if __name__ == "__main__":
    main()
