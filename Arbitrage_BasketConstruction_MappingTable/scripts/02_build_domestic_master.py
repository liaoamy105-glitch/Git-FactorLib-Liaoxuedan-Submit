import re
from pathlib import Path

import pandas as pd


FUTURES_INPUT = Path("data/raw/tushare_futures_basic.csv")
OPTIONS_INPUT = Path("data/raw/tushare_options_basic.csv")
MASTER_OUTPUT = Path("data/processed/domestic_master.csv")
SUMMARY_OUTPUT = Path("data/processed/domestic_master_summary.csv")
LINK_CHECK_OUTPUT = Path("data/processed/domestic_option_future_link_check.csv")

EXCHANGE_NAMES = {
    "SHFE": "上海期货交易所",
    "INE": "上海国际能源交易中心",
    "DCE": "大连商品交易所",
    "CZCE": "郑州商品交易所",
    "GFEX": "广州期货交易所",
    "CFFEX": "中国金融期货交易所",
}

CFFEX_STOCK_INDEX_OPTIONS = {
    "IO": {
        "product_name": "沪深300股指期权",
        "underlying_product_code": "IF",
    },
    "HO": {
        "product_name": "上证50股指期权",
        "underlying_product_code": "IH",
    },
    "MO": {
        "product_name": "中证1000股指期权",
        "underlying_product_code": "IM",
    },
}

MASTER_COLUMNS = [
    "instrument_type",
    "product_code",
    "product_name",
    "exchange_code",
    "exchange_name",
    "country",
    "underlying_product_code",
    "underlying_instrument_type",
    "option_type",
    "exercise_type",
    "trade_unit",
    "per_unit",
    "quote_unit",
    "quote_unit_desc",
    "min_price_chg",
    "delivery_mode",
    "first_list_date",
    "last_delist_date",
    "sample_contract_count",
    "source",
    "source_file",
]


def read_raw_csv(path):
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path, dtype=str)
    except pd.errors.EmptyDataError:
        return pd.DataFrame()


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_code(value):
    return clean_text(value).upper()


def most_frequent(series):
    values = series.dropna().astype(str).str.strip()
    values = values[values != ""]
    if values.empty:
        return ""
    return values.value_counts().index[0]


def min_text(series):
    values = series.dropna().astype(str).str.strip()
    values = values[values != ""]
    if values.empty:
        return ""
    return values.min()


def max_text(series):
    values = series.dropna().astype(str).str.strip()
    values = values[values != ""]
    if values.empty:
        return ""
    return values.max()


def futures_product_name(name):
    cleaned = clean_text(name)
    cleaned = re.sub(r"\d+", "", cleaned)
    return cleaned.strip()


def extract_option_product_code(row):
    for column in ["opt_code", "ts_code"]:
        raw_code = clean_text(row.get(column, ""))
        if not raw_code:
            continue
        code = raw_code.split(".")[0].upper()
        code = re.sub(r"^OP", "", code, flags=re.IGNORECASE)
        match = re.match(r"^[A-Z]+", code)
        if match:
            return match.group(0)
    return ""


def option_product_name(name):
    cleaned = clean_text(name)
    cleaned = cleaned.replace("期权", "")
    cleaned = re.sub(r"\d+(\.\d+)?", "", cleaned)
    cleaned = re.sub(r"[CP]$", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"[CP]", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip()
    if not cleaned:
        return ""
    return f"{cleaned}期权"


def build_futures_master(futures_raw):
    if futures_raw.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    futures = futures_raw.copy()
    futures["product_code"] = futures["fut_code"].apply(normalize_code)
    futures["exchange_code"] = futures["exchange"].apply(clean_text)
    futures = futures[(futures["product_code"] != "") & (futures["exchange_code"] != "")]

    rows = []
    for (exchange_code, product_code), group in futures.groupby(
        ["exchange_code", "product_code"], dropna=False
    ):
        rows.append(
            {
                "instrument_type": "FUTURE",
                "product_code": product_code,
                "product_name": futures_product_name(most_frequent(group["name"])),
                "exchange_code": exchange_code,
                "exchange_name": EXCHANGE_NAMES.get(exchange_code, ""),
                "country": "中国",
                "underlying_product_code": "",
                "underlying_instrument_type": "",
                "option_type": "",
                "exercise_type": "",
                "trade_unit": most_frequent(group["trade_unit"]),
                "per_unit": most_frequent(group["per_unit"]),
                "quote_unit": most_frequent(group["quote_unit"]),
                "quote_unit_desc": most_frequent(group["quote_unit_desc"]),
                "min_price_chg": "",
                "delivery_mode": most_frequent(group["d_mode_desc"]),
                "first_list_date": min_text(group["list_date"]),
                "last_delist_date": max_text(group["delist_date"]),
                "sample_contract_count": len(group),
                "source": "tushare_futures_basic",
                "source_file": str(FUTURES_INPUT),
            }
        )

    return pd.DataFrame(rows, columns=MASTER_COLUMNS)


def build_options_master(options_raw):
    if options_raw.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    options = options_raw.copy()
    options["product_code"] = options.apply(extract_option_product_code, axis=1)
    options["exchange_code"] = options["exchange"].apply(clean_text)
    options = options[(options["product_code"] != "") & (options["exchange_code"] != "")]

    rows = []
    for (exchange_code, product_code), group in options.groupby(
        ["exchange_code", "product_code"], dropna=False
    ):
        if exchange_code == "CFFEX":
            stock_index_rule = CFFEX_STOCK_INDEX_OPTIONS.get(product_code, {})
            instrument_type = "STOCK_INDEX_OPTION"
            product_name = stock_index_rule.get(
                "product_name", option_product_name(most_frequent(group["name"]))
            )
            underlying_product_code = stock_index_rule.get(
                "underlying_product_code", ""
            )
        else:
            instrument_type = "OPTION"
            product_name = option_product_name(most_frequent(group["name"]))
            underlying_product_code = product_code

        rows.append(
            {
                "instrument_type": instrument_type,
                "product_code": product_code,
                "product_name": product_name,
                "exchange_code": exchange_code,
                "exchange_name": EXCHANGE_NAMES.get(exchange_code, ""),
                "country": "中国",
                "underlying_product_code": underlying_product_code,
                "underlying_instrument_type": "FUTURE",
                "option_type": most_frequent(group["opt_type"]),
                "exercise_type": most_frequent(group["exercise_type"]),
                "trade_unit": "",
                "per_unit": most_frequent(group["per_unit"]),
                "quote_unit": most_frequent(group["quote_unit"]),
                "quote_unit_desc": "",
                "min_price_chg": most_frequent(group["min_price_chg"]),
                "delivery_mode": "",
                "first_list_date": min_text(group["list_date"]),
                "last_delist_date": max_text(group["delist_date"]),
                "sample_contract_count": len(group),
                "source": "tushare_options_basic",
                "source_file": str(OPTIONS_INPUT),
            }
        )

    return pd.DataFrame(rows, columns=MASTER_COLUMNS)


def build_summary(master):
    if master.empty:
        return pd.DataFrame(columns=["instrument_type", "exchange_code", "record_count"])

    return (
        master.groupby(["instrument_type", "exchange_code"], dropna=False)
        .size()
        .reset_index(name="record_count")
    )


def build_option_future_link_check(master):
    columns = [
        "option_exchange_code",
        "option_product_code",
        "option_product_name",
        "underlying_product_code",
        "matched_future",
        "message",
    ]

    futures = master[master["instrument_type"] == "FUTURE"]
    options = master[
        master["instrument_type"].isin(["OPTION", "STOCK_INDEX_OPTION"])
    ]
    future_keys = set(zip(futures["exchange_code"], futures["product_code"]))

    rows = []
    for _, option in options.iterrows():
        key = (option["exchange_code"], option["underlying_product_code"])
        matched = key in future_keys
        if (
            option["instrument_type"] == "STOCK_INDEX_OPTION"
            and option["exchange_code"] == "CFFEX"
            and option["product_code"] not in CFFEX_STOCK_INDEX_OPTIONS
        ):
            message = "unknown CFFEX stock index option code"
        elif matched:
            message = ""
        else:
            message = "underlying future not found in domestic_master"

        rows.append(
            {
                "option_exchange_code": option["exchange_code"],
                "option_product_code": option["product_code"],
                "option_product_name": option["product_name"],
                "underlying_product_code": option["underlying_product_code"],
                "matched_future": matched,
                "message": message,
            }
        )

    return pd.DataFrame(rows, columns=columns)


def main():
    futures_raw = read_raw_csv(FUTURES_INPUT)
    options_raw = read_raw_csv(OPTIONS_INPUT)

    futures_master = build_futures_master(futures_raw)
    options_master = build_options_master(options_raw)
    domestic_master = pd.concat(
        [futures_master, options_master], ignore_index=True
    ).reindex(columns=MASTER_COLUMNS)

    MASTER_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    domestic_master.to_csv(MASTER_OUTPUT, index=False, encoding="utf-8-sig")

    summary = build_summary(domestic_master)
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    link_check = build_option_future_link_check(domestic_master)
    link_check.to_csv(LINK_CHECK_OUTPUT, index=False, encoding="utf-8-sig")

    print("domestic master built successfully")
    print(f"domestic_master rows: {len(domestic_master)}")
    print(f"summary saved to {SUMMARY_OUTPUT}")
    print(f"option future link check saved to {LINK_CHECK_OUTPUT}")


if __name__ == "__main__":
    main()
