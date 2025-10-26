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


def _naive_utc_series(s: pd.Series) -> pd.Series:
    """Remove timezone (ou converte a UTC e remove) para ficar datetime64[ns] 'naive'."""
    try:
        s = s.dt.tz_convert("UTC").dt.tz_localize(None)
    except Exception:
        try:
            s = s.dt.tz_localize(None)
        except Exception:
            pass
    return s


def format_dt_series_iso(s: pd.Series) -> pd.Series:
    s = _naive_utc_series(s)
    return s.dt.strftime("%Y-%m-%dT%H:%M:%S")


def format_dt_series_epoch(s: pd.Series, unit: str) -> pd.Series:
    """
    Converte datetime64 para Unix timestamp (segundos ou milissegundos).
    Mantém None onde houver NaT.
    """
    s = _naive_utc_series(s)
    mask = s.notna()
    # view('int64') dá nanos desde epoch; cuidado com NaT (-9223372036854775808)
    ns = s.view("int64")
    out = pd.Series([None] * len(s), index=s.index, dtype="object")
    if unit == "ms":
        out.loc[mask] = (ns[mask] // 1_000_000).astype("int64")
    else:  # "s"
        out.loc[mask] = (ns[mask] // 1_000_000_000).astype("int64")
    return out


def coerce_datetimes(
    df: pd.DataFrame,
    *,
    dayfirst: bool,
    auto_dates: bool,
    forced_cols: set[str],
    mode: str,             # "iso" ou "epoch"
    epoch_unit: str = "s", # "s" ou "ms"
    threshold: float = 0.8,
    date_format: str | None = None,
) -> pd.DataFrame:
    """
    - Converte colunas já datetime64.
    - Para colunas texto/object, tenta parsear; só converte se >= threshold parecerem datas.
    - Nunca converte colunas numéricas.
    - 'forced_cols' sempre tenta converter.
    - 'mode' define saída: "iso" (string) ou "epoch" (int s/ms).
    """
    for col in df.columns:
        s = df[col]

        # 1) Já é datetime?
        if pd.api.types.is_datetime64_any_dtype(s):
            df[col] = (format_dt_series_epoch(s, epoch_unit)
                       if mode == "epoch" else
                       format_dt_series_iso(s))
            continue

        # 2) Forçadas (não numéricas)
        if col in forced_cols and not (pd.api.types.is_integer_dtype(s) or pd.api.types.is_float_dtype(s)):
            parsed = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst, utc=False,
                                    format=(date_format or None))
            if parsed.notna().any():
                df[col] = (format_dt_series_epoch(parsed, epoch_unit)
                           if mode == "epoch" else
                           format_dt_series_iso(parsed))
            continue

        # 3) Heurística automática em texto/object
        if auto_dates and (pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s)):
            parsed = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst, utc=False,
                                    format=(date_format or None))
            non_null = s.notna().sum()
            ratio = (parsed.notna().sum() / non_null) if non_null else 0.0
            if ratio >= threshold and parsed.notna().any():
                df[col] = (format_dt_series_epoch(parsed, epoch_unit)
                           if mode == "epoch" else
                           format_dt_series_iso(parsed))

        # 4) Numéricas → nunca tratar como data

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

    # Opções de data
    ap.add_argument("--date-iso", dest="date_iso", action="store_true",
                    help="Formatar colunas de data em ISO8601")
    ap.add_argument("--date-epoch", dest="date_epoch", action="store_true",
                    help="Formatar colunas de data como Unix timestamp (epoch)")
    ap.add_argument("--epoch-unit", dest="epoch_unit", default="s", choices=["s", "ms"],
                    help="Unidade do timestamp (s=segundos, ms=milissegundos)")
    ap.add_argument("--dayfirst", dest="dayfirst", action="store_true",
                    help="Interpretar datas como DD/MM/AAAA (pt-BR)")
    ap.add_argument("--date-cols", dest="date_cols", default="",
                    help="Lista de colunas (separadas por vírgula) para forçar como data")
    ap.add_argument("--no-auto-dates", dest="no_auto_dates", action="store_true",
                    help="Desabilita a detecção automática de colunas de data")
    ap.add_argument("--date-format", dest="date_format", default="",
                    help="Formato strptime para datas (ex.: '%d/%m/%Y %H:%M') para parsing consistente")

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

    # Datas -> ISO ou EPOCH
    if args.date_iso or args.date_epoch:
        mode = "epoch" if args.date_epoch else "iso"
        forced_cols = {c.strip() for c in args.date_cols.split(",") if c.strip()}
        df = coerce_datetimes(
            df,
            dayfirst=args.dayfirst,
            auto_dates=not args.no_auto_dates,
            forced_cols=forced_cols,
            mode=mode,
            epoch_unit=args.epoch_unit,
            date_format=(args.date_format or None),
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
