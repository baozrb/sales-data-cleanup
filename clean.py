"""Load each export, map it onto the common schema, then validate and de-duplicate."""
from datetime import datetime

import pandas as pd

from config import COLUMNS, REPORT_TZ

# Why a row was left out of the report, in the order the checks run.
MISSING_ID = "Missing order ID"
BAD_DATE = "Date could not be read"
MISSING_SKU = "Missing SKU"
BAD_QTY = "Quantity not a positive whole number (e.g. a return)"
BAD_PRICE = "Missing or invalid price"
UNKNOWN_SKU = "SKU not in product list"
EXACT_DUP = "Exact duplicate row"
REPLACED = "Same order line found in a newer export"
CANCELLED = "Cancelled or refunded"
REASONS = [MISSING_ID, BAD_DATE, MISSING_SKU, BAD_QTY, BAD_PRICE, UNKNOWN_SKU, EXACT_DUP, REPLACED, CANCELLED]


def load_source(name, cfg, path):
    """Read one export file and rename its columns to the common schema."""
    read = cfg.get("read", {})
    if path.suffix.lower() == ".xlsx":
        df = pd.read_excel(path, dtype=object, **read)  # object keeps real date cells as dates
        first_data_row = read.get("skiprows", 0) + 2
    else:
        df = pd.read_csv(path, dtype=str, **read)
        first_data_row = 2
    df["source_row"] = df.index + first_data_row  # the row number you see in Excel

    df = df.rename(columns=cfg["columns"])
    for col in COLUMNS:
        if col not in df:
            df[col] = None
    df = df[COLUMNS + ["source_row"]].dropna(how="all", subset=COLUMNS)  # skip blank rows

    if cfg.get("fill_down"):
        cols = cfg["fill_down"]
        df[cols] = df.groupby("order_id")[cols].ffill()

    df.insert(0, "source", name)
    df.insert(1, "source_file", path.name)
    return df


def standardize(df, cfg):
    """Turn text into clean types: dates, numbers, upper-case SKUs, lower-case emails."""
    out = df.copy()
    out["order_id"] = clean_text(out["order_id"])
    out["order_date"] = parse_dates(out["order_date"], cfg["date_format"])
    out["sku"] = clean_text(out["sku"]).str.upper().str.replace(r"\s+", "", regex=True)
    out["qty"] = to_number(clean_text(out["qty"]))
    out["unit_price"] = parse_money(out["unit_price"])
    if cfg.get("price_is_line_total"):
        out["unit_price"] = out["unit_price"] / out["qty"]
    out["email"] = clean_text(out["email"]).str.lower()
    out["status"] = clean_text(out["status"]).str.lower().fillna("completed")
    out["excluded"] = out["status"].isin(cfg["exclude_status"])
    return out


def process(frames, products):
    """Combine the sources. Returns clean rows, and removed rows with the reason for each."""
    df = pd.concat(frames, ignore_index=True).sort_values(["source", "source_file", "source_row"])
    df = df.merge(products[["sku", "product", "category"]], on="sku", how="left")
    removed = []

    def drop(mask, reason):
        nonlocal df
        removed.append(df[mask].assign(reason=reason))
        df = df[~mask]

    drop(df["order_id"].isna(), MISSING_ID)
    drop(df["order_date"].isna(), BAD_DATE)
    drop(df["sku"].isna(), MISSING_SKU)
    drop(~((df["qty"] > 0) & (df["qty"] % 1 == 0)), BAD_QTY)
    drop(~(df["unit_price"] > 0), BAD_PRICE)
    drop(df["product"].isna(), UNKNOWN_SKU)
    drop(df.duplicated(["source"] + COLUMNS), EXACT_DUP)
    drop(df.duplicated(["source", "order_id", "sku"], keep="last"), REPLACED)
    drop(df["excluded"], CANCELLED)

    df = df.assign(qty=df["qty"].astype(int), unit_price=df["unit_price"].round(2))
    df = df.assign(revenue=(df["qty"] * df["unit_price"]).round(2))
    clean = df[["source", "order_id", "order_date", "sku", "product", "category", "qty",
                "unit_price", "revenue", "email", "source_file", "source_row"]]
    return clean.reset_index(drop=True), pd.concat(removed, ignore_index=True).drop(columns="excluded")


def clean_text(values):
    text = values.astype("str").str.strip()
    return text.where(text != "")


def to_number(text):
    return pd.to_numeric(text, errors="coerce").astype(float)


def parse_money(values):
    """'$1,234.50', '1234.5' and 1234.5 all become 1234.5."""
    return to_number(clean_text(values).str.replace(r"[$,]|USD", "", regex=True).str.strip())


def parse_dates(values, fmt):
    """Parse with the source's known format. pandas is never left to guess day vs month."""
    if fmt == "ISO8601" or "%z" in fmt:
        ts = pd.to_datetime(clean_text(values), format=fmt, utc=True, errors="coerce")
        return ts.dt.tz_convert(REPORT_TZ).dt.tz_localize(None).dt.normalize()
    return pd.to_datetime(values.map(lambda v: _parse_local_date(v, fmt)))


def _parse_local_date(value, fmt):
    if isinstance(value, datetime):  # a real date cell in Excel
        return pd.Timestamp(value).normalize()
    try:
        return pd.Timestamp(datetime.strptime(str(value).strip(), fmt))
    except ValueError:
        return pd.NaT
