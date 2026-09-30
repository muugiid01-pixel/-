"""
data_processor.py — Банкны CSV файлыг унших, цэвэрлэх, боловсруулах.

Workflow:  CSV унших → багана таних → огноо/дүн цэвэрлэх → ангилах
"""

import io
import re

import pandas as pd

import config
from categorizer import classify_dataframe

_DATE_FORMATS = [
    "%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d",
    "%d-%m-%Y", "%d.%m.%Y", "%d %b %Y", "%d %B %Y", "%Y-%m-%d %H:%M:%S",
]
_ENCODINGS = ["utf-8-sig", "cp1251", "latin-1"]


class CSVProcessingError(Exception):
    """CSV боловсруулах үед гарсан, хэрэглэгчид харуулах Монгол алдаа."""


# ---------------------------------------------------------------------------
# Утга цэвэрлэх туслах функцүүд
# ---------------------------------------------------------------------------
def parse_amount(value):
    """
    Дүнг float болгоно. Хоосон бол None, тоо биш бол ValueError.

    Дэмжих хэлбэрүүд: "1,500,000", "-85,000", "$1,234.50", "(500)",
    "500-", "250.00 CR", "250.00 DR", "1.234,56", "₮ 50 000".
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return None if pd.isna(value) else float(value)

    text = str(value).strip()
    if text == "" or text.lower() in {"nan", "none", "null", "-"}:
        return None

    # Валютын тэмдэг, зай зэргийг арилгана.
    text = re.sub(r"(?i)(aud|mnt|usd|eur|[$₮€£\s\u00a0'])", "", text)

    negative = False
    upper = text.upper()
    if upper.endswith("DR"):
        negative, text = True, text[:-2]
    elif upper.endswith("CR"):
        text = text[:-2]
    if text.startswith("(") and text.endswith(")"):
        negative, text = True, text[1:-1]
    if text.endswith("-"):
        negative, text = True, text[:-1]
    if text.startswith("+"):
        text = text[1:]
    if text.startswith("-"):
        negative, text = not negative, text[1:]

    # Тоо, цэг, таслалаас өөр тэмдэгт үлдвэл (жишээ нь "29/6/2025") тоо биш.
    if text == "" or re.search(r"[^\d.,]", text):
        raise ValueError(value)

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):   # 1.234,56
            text = text.replace(".", "").replace(",", ".")
        else:                                    # 1,234.56
            text = text.replace(",", "")
    elif "," in text:
        if re.fullmatch(r"\d{1,3}(,\d{3})+", text):
            text = text.replace(",", "")        # 1,500,000
        else:
            text = text.replace(",", ".")       # 12,5
    if text.count(".") > 1:
        raise ValueError(value)

    number = float(text)
    return -number if negative else number


def parse_date(value):
    """Огноог pandas Timestamp болгоно. Уншигдахгүй бол NaT."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    text = str(value).strip()
    if not text:
        return pd.NaT
    for fmt in _DATE_FORMATS:
        try:
            return pd.to_datetime(text, format=fmt)
        except (ValueError, TypeError):
            continue
    try:
        return pd.to_datetime(text, dayfirst=True)
    except (ValueError, TypeError):
        return pd.NaT


def _looks_like_date(value):
    """Утга огноо шиг харагдаж байгаа эсэх (header-гүй CSV таних)."""
    text = str(value).strip()
    return bool(re.search(r"\d", text)) and not pd.isna(parse_date(text)) and len(text) >= 6


def _looks_like_number(value):
    """Утга тоо шиг харагдаж байгаа эсэх."""
    try:
        return parse_amount(value) is not None
    except ValueError:
        return False


def make_dup_key(date, description, amount):
    """Duplicate шалгах түлхүүр: Огноо + Тайлбар + Дүн."""
    desc = re.sub(r"\s+", " ", str(description or "")).strip().lower()
    return f"{date}|{desc}|{float(amount):.2f}"


# ---------------------------------------------------------------------------
# CSV унших
# ---------------------------------------------------------------------------
def _decode(raw_bytes):
    """Файлын encoding-ийг таах (UTF-8, Cyrillic cp1251, Latin-1)."""
    for enc in _ENCODINGS:
        try:
            return raw_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    raise CSVProcessingError(config.ERRORS["bad_format"])


def read_csv(source):
    """
    CSV-г уншиж, бүх утгыг текст хэлбэрээр DataFrame болгоно.

    Args:
        source: файлын зам (str), bytes, эсвэл Streamlit UploadedFile.
    Returns:
        (DataFrame, has_header)
    """
    if isinstance(source, (bytes, bytearray)):
        raw = bytes(source)
    elif isinstance(source, str):
        with open(source, "rb") as fh:
            raw = fh.read()
    else:
        raw = source.getvalue() if hasattr(source, "getvalue") else source.read()

    if not raw or not raw.strip():
        raise CSVProcessingError(config.ERRORS["empty"])

    text = _decode(raw)
    try:
        df = pd.read_csv(io.StringIO(text), header=None, dtype=str, sep=None,
                         engine="python", skip_blank_lines=True)
    except Exception as exc:
        raise CSVProcessingError(f"{config.ERRORS['bad_format']} ({exc})")

    df = df.dropna(how="all")
    if df.empty:
        raise CSVProcessingError(config.ERRORS["empty"])

    first_row = df.iloc[0].fillna("").tolist()
    has_header = _is_header_row(first_row)
    if has_header:
        df.columns = [str(c).strip() if str(c).strip() else f"column_{i}"
                      for i, c in enumerate(first_row)]
        df = df.iloc[1:].reset_index(drop=True)
        if df.empty:
            raise CSVProcessingError(config.ERRORS["empty"])
    else:
        df = df.reset_index(drop=True)
    return df, has_header


def _is_header_row(cells):
    """Эхний мөр баганын нэр (header) эсэхийг тодорхойлно."""
    normalized = [str(c).strip().lower() for c in cells]
    all_aliases = {a for aliases in config.COLUMN_ALIASES.values() for a in aliases}
    if any(c in all_aliases for c in normalized):
        return True
    has_date = any(_looks_like_date(c) for c in cells)
    has_number = any(_looks_like_number(c) for c in cells)
    return not (has_date and has_number)


# ---------------------------------------------------------------------------
# Багана таних
# ---------------------------------------------------------------------------
def detect_columns(df, has_header):
    """
    Баганын нэрсийг стандарт нэр рүү (date, description, amount, income,
    expense, balance) тааруулна.

    Returns:
        dict: стандарт нэр → CSV дахь баганын нэр
    """
    if has_header:
        return _detect_by_header(df.columns)
    return _detect_by_content(df)


def _detect_by_header(columns):
    """Header-ийн нэрээр багана таних (эхлээд яг тааруулж, дараа нь агуулгаар)."""
    mapping = {}
    lowered = {col: str(col).strip().lower() for col in columns}

    for key, aliases in config.COLUMN_ALIASES.items():
        for col, name in lowered.items():
            if col not in mapping.values() and name in aliases:
                mapping[key] = col
                break

    for key, aliases in config.COLUMN_ALIASES.items():
        if key in mapping:
            continue
        for col, name in lowered.items():
            if col in mapping.values():
                continue
            if any(len(a) > 3 and a in name for a in aliases):
                mapping[key] = col
                break
    return mapping


def _detect_by_content(df):
    """
    Header-гүй CSV (жишээ нь CommBank: Огноо, Дүн, Тайлбар, Үлдэгдэл)-ийн
    багануудыг агуулгаар нь таних.
    """
    sample = df.head(50)
    mapping = {}
    numeric_cols, text_cols = [], []

    for col in df.columns:
        values = [v for v in sample[col].dropna().tolist() if str(v).strip()]
        if not values:
            continue
        date_share = sum(_looks_like_date(v) for v in values) / len(values)
        num_share = sum(_looks_like_number(v) for v in values) / len(values)
        if "date" not in mapping and date_share >= 0.8 and num_share < 0.8:
            mapping["date"] = col
        elif num_share >= 0.8:
            numeric_cols.append(col)
        else:
            text_cols.append(col)

    if numeric_cols:
        mapping["amount"] = numeric_cols[0]
    if len(numeric_cols) >= 2:
        mapping["balance"] = numeric_cols[-1]
    if text_cols:
        mapping["description"] = max(
            text_cols, key=lambda c: sample[c].fillna("").astype(str).str.len().mean()
        )
    return mapping


# ---------------------------------------------------------------------------
# Цэвэрлэх
# ---------------------------------------------------------------------------
def _parse_numeric_column(series, label, warnings):
    """Баганын утгуудыг тоо болгоно. Тоо биш утгуудыг NaN болгож анхааруулна."""
    values, bad_rows = [], []
    for idx, raw in series.items():
        try:
            parsed = parse_amount(raw)
            values.append(float("nan") if parsed is None else parsed)
        except ValueError:
            values.append(float("nan"))
            bad_rows.append(idx + 1)
    if bad_rows:
        non_empty = series.fillna("").astype(str).str.strip().ne("").sum()
        if len(bad_rows) >= non_empty:
            raise CSVProcessingError(f"{config.ERRORS['not_numeric']} (багана: {label})")
        warnings.append(
            f"{config.ERRORS['not_numeric']} '{label}' баганын {len(bad_rows)} мөрийг алгаслаа "
            f"(мөр: {', '.join(map(str, bad_rows[:10]))}{'…' if len(bad_rows) > 10 else ''})."
        )
    return pd.Series(values, index=series.index, dtype="float64")


def clean_dataframe(df, mapping):
    """
    Танигдсан багануудаас стандарт DataFrame үүсгэнэ.

    Returns:
        (DataFrame[date, description, amount, balance], warnings)
    """
    warnings = []

    if "date" not in mapping:
        raise CSVProcessingError(config.ERRORS["no_date"])
    has_amount = "amount" in mapping
    has_split = "income" in mapping or "expense" in mapping
    if not has_amount and not has_split:
        raise CSVProcessingError(config.ERRORS["no_amount"])

    out = pd.DataFrame(index=df.index)

    # Огноо
    dates = df[mapping["date"]].apply(parse_date)
    bad_dates = dates.isna()
    if bad_dates.all():
        raise CSVProcessingError(config.ERRORS["bad_date"])
    if bad_dates.any():
        rows = [str(i + 1) for i in df.index[bad_dates][:10]]
        warnings.append(f"{config.ERRORS['bad_date']} {int(bad_dates.sum())} мөрийг алгаслаа (мөр: {', '.join(rows)}).")
    out["date"] = dates

    # Тайлбар
    if "description" in mapping:
        out["description"] = df[mapping["description"]].fillna("").astype(str)
    else:
        used = set(mapping.values())
        others = [c for c in df.columns if c not in used]
        out["description"] = (
            df[others].fillna("").astype(str).agg(" ".join, axis=1) if others else ""
        )
    out["description"] = out["description"].str.replace(r"\s+", " ", regex=True).str.strip()

    # Дүн — нэг Amount багана эсвэл тусдаа Орлого/Зарлага багана
    if has_amount:
        out["amount"] = _parse_numeric_column(df[mapping["amount"]], mapping["amount"], warnings)
    else:
        income = (_parse_numeric_column(df[mapping["income"]], mapping["income"], warnings)
                  if "income" in mapping else pd.Series(float("nan"), index=df.index))
        expense = (_parse_numeric_column(df[mapping["expense"]], mapping["expense"], warnings)
                   if "expense" in mapping else pd.Series(float("nan"), index=df.index))
        both_empty = income.isna() & expense.isna()
        out["amount"] = income.fillna(0).abs() - expense.fillna(0).abs()
        out.loc[both_empty, "amount"] = float("nan")

    # Үлдэгдэл
    if "balance" in mapping:
        out["balance"] = _parse_numeric_column(df[mapping["balance"]], mapping["balance"], warnings)
    else:
        out["balance"] = float("nan")

    missing_amount = out["amount"].isna() & ~bad_dates
    if missing_amount.any():
        warnings.append(f"Дүнгүй {int(missing_amount.sum())} мөрийг алгаслаа.")

    out = out[~bad_dates & out["amount"].notna()].copy()
    if out.empty:
        raise CSVProcessingError(config.ERRORS["empty"])

    out["date"] = out["date"].dt.strftime("%Y-%m-%d")
    out["amount"] = out["amount"].round(2)
    out["balance"] = out["balance"].round(2)
    return out.reset_index(drop=True), warnings


def process_csv(source, source_name="uploaded.csv"):
    """
    CSV-г бүрэн боловсруулна: унших → багана таних → цэвэрлэх → ангилах.

    Returns:
        (DataFrame, mapping, warnings)
    """
    raw_df, has_header = read_csv(source)
    mapping = detect_columns(raw_df, has_header)
    clean_df, warnings = clean_dataframe(raw_df, mapping)
    processed = classify_dataframe(clean_df)
    processed["source_file"] = source_name
    processed["dup_key"] = [
        make_dup_key(d, desc, a)
        for d, desc, a in zip(processed["date"], processed["description"], processed["amount"])
    ]
    return processed, mapping, warnings
