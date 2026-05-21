import re
from pathlib import Path

import pandas as pd


ETF_OPTIONS_INPUT = Path("data/raw/tushare_etf_options_basic.csv")
FUND_BASIC_INPUT = Path("data/raw/tushare_fund_basic.csv")
INDEX_BASIC_INPUT = Path("data/raw/tushare_index_basic.csv")
DOMESTIC_MASTER_INPUT = Path("data/processed/domestic_master.csv")

ETF_INDEX_MASTER_OUTPUT = Path("data/processed/etf_index_master.csv")
DOMESTIC_MASTER_WITH_ETF_OUTPUT = Path("data/processed/domestic_master_with_etf.csv")
UNDERLYING_CHECK_OUTPUT = Path("data/processed/etf_option_underlying_check.csv")
SUMMARY_OUTPUT = Path("data/processed/etf_index_master_summary.csv")

EXCHANGE_NAMES = {
    "SSE": "上海证券交易所",
    "SZSE": "深圳证券交易所",
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
        return pd.read_csv(path, dtype=str).fillna("")
    except pd.errors.EmptyDataError:
        return pd.DataFrame()


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_exchange(value):
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


def exchange_from_ts_code(ts_code):
    code = clean_text(ts_code).upper()
    if code.endswith(".SH"):
        return "SSE"
    if code.endswith(".SZ"):
        return "SZSE"
    return ""


def extract_underlying_etf_code(row):
    opt_code = clean_text(row.get("opt_code", "")).upper()
    if opt_code.startswith("OP") and re.match(r"^OP\d{6}\.(SH|SZ)$", opt_code):
        return opt_code[2:]

    for column in ["ts_code", "name"]:
        value = clean_text(row.get(column, "")).upper()
        match = re.search(r"(?:OP)?(\d{6}\.(?:SH|SZ))", value)
        if match:
            return match.group(1)

    return ""


def etf_option_product_name(name, underlying_etf_code):
    text = clean_text(name)
    if "期权" in text:
        return text.split("期权", 1)[0].strip() + "期权"
    return f"{underlying_etf_code}ETF期权"


def index_product_name(name):
    text = clean_text(name)
    if not text:
        return ""
    if text.endswith("指数"):
        return text
    return f"{text}指数"


def build_etf_option_master(etf_options_raw):
    if etf_options_raw.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    options = etf_options_raw.copy()
    options["underlying_etf_code"] = options.apply(extract_underlying_etf_code, axis=1)
    options["exchange_code"] = options["exchange"].apply(normalize_exchange)
    options.loc[options["exchange_code"] == "", "exchange_code"] = options.loc[
        options["exchange_code"] == "", "fetch_exchange"
    ].apply(normalize_exchange)
    options = options[options["underlying_etf_code"] != ""]

    rows = []
    for underlying_etf_code, group in options.groupby("underlying_etf_code"):
        exchange_code = most_frequent(group["exchange_code"])
        rows.append(
            {
                "instrument_type": "ETF_OPTION",
                "product_code": f"OP{underlying_etf_code}",
                "product_name": etf_option_product_name(
                    most_frequent(group["name"]), underlying_etf_code
                ),
                "exchange_code": exchange_code,
                "exchange_name": EXCHANGE_NAMES.get(exchange_code, ""),
                "country": "中国",
                "underlying_product_code": underlying_etf_code,
                "underlying_instrument_type": "ETF_SPOT",
                "option_type": most_frequent(group["opt_type"]),
                "exercise_type": most_frequent(group["exercise_type"]),
                "trade_unit": "",
                "per_unit": most_frequent(group["per_unit"]),
                "quote_unit": most_frequent(group["quote_unit"]),
                "quote_unit_desc": "",
                "min_price_chg": most_frequent(group["min_price_chg"]),
                "delivery_mode": "实物交割/证券给付，待交易所规则核验",
                "first_list_date": min_text(group["list_date"]),
                "last_delist_date": max_text(group["delist_date"]),
                "sample_contract_count": len(group),
                "source": "tushare_etf_options_basic",
                "source_file": str(ETF_OPTIONS_INPUT),
            }
        )

    return pd.DataFrame(rows, columns=MASTER_COLUMNS)


def build_etf_spot_master(fund_basic_raw, etf_options_master):
    if fund_basic_raw.empty or etf_options_master.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    required_codes = set(etf_options_master["underlying_product_code"])
    funds = fund_basic_raw.copy()
    funds["ts_code"] = funds["ts_code"].apply(lambda value: clean_text(value).upper())
    funds = funds[funds["ts_code"].isin(required_codes)]

    rows = []
    for _, fund in funds.iterrows():
        exchange_code = exchange_from_ts_code(fund["ts_code"])
        rows.append(
            {
                "instrument_type": "ETF_SPOT",
                "product_code": fund["ts_code"],
                "product_name": fund.get("name", ""),
                "exchange_code": exchange_code,
                "exchange_name": EXCHANGE_NAMES.get(exchange_code, ""),
                "country": "中国",
                "underlying_product_code": fund.get("benchmark", ""),
                "underlying_instrument_type": "INDEX_SPOT",
                "option_type": "",
                "exercise_type": "",
                "trade_unit": "",
                "per_unit": "",
                "quote_unit": "",
                "quote_unit_desc": fund.get("benchmark", ""),
                "min_price_chg": "",
                "delivery_mode": "",
                "first_list_date": fund.get("list_date", ""),
                "last_delist_date": fund.get("delist_date", ""),
                "sample_contract_count": 1,
                "source": "tushare_fund_basic",
                "source_file": str(FUND_BASIC_INPUT),
            }
        )

    return pd.DataFrame(rows, columns=MASTER_COLUMNS)


def build_index_spot_master(index_basic_raw):
    if index_basic_raw.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    index_basic = index_basic_raw.copy()
    if "ts_code" in index_basic.columns:
        index_basic["ts_code"] = index_basic["ts_code"].apply(
            lambda value: clean_text(value).upper()
        )
        index_basic = index_basic.drop_duplicates(subset=["ts_code"], keep="first")

    rows = []
    for _, index in index_basic.iterrows():
        exchange_code = normalize_exchange(index.get("market", ""))
        if not exchange_code:
            exchange_code = exchange_from_ts_code(index.get("ts_code", ""))
        rows.append(
            {
                "instrument_type": "INDEX_SPOT",
                "product_code": index.get("ts_code", ""),
                "product_name": index_product_name(index.get("name", "")),
                "exchange_code": exchange_code,
                "exchange_name": EXCHANGE_NAMES.get(exchange_code, ""),
                "country": "中国",
                "underlying_product_code": "",
                "underlying_instrument_type": "",
                "option_type": "",
                "exercise_type": "",
                "trade_unit": "",
                "per_unit": "",
                "quote_unit": "",
                "quote_unit_desc": index.get("category", ""),
                "min_price_chg": "",
                "delivery_mode": "",
                "first_list_date": index.get("list_date", ""),
                "last_delist_date": "",
                "sample_contract_count": 1,
                "source": "tushare_index_basic",
                "source_file": str(INDEX_BASIC_INPUT),
            }
        )

    return pd.DataFrame(rows, columns=MASTER_COLUMNS)


def build_underlying_check(etf_options_master, etf_spot_master):
    columns = [
        "etf_option_product_code",
        "etf_option_product_name",
        "underlying_etf_code",
        "matched_etf_spot",
        "message",
    ]

    etf_spot_codes = set(etf_spot_master["product_code"])
    rows = []
    for _, option in etf_options_master.iterrows():
        matched = option["underlying_product_code"] in etf_spot_codes
        rows.append(
            {
                "etf_option_product_code": option["product_code"],
                "etf_option_product_name": option["product_name"],
                "underlying_etf_code": option["underlying_product_code"],
                "matched_etf_spot": matched,
                "message": "underlying ETF spot found"
                if matched
                else "underlying ETF spot not found in tushare_fund_basic",
            }
        )

    return pd.DataFrame(rows, columns=columns)


def build_summary(etf_index_master):
    if etf_index_master.empty:
        return pd.DataFrame(columns=["instrument_type", "exchange_code", "record_count"])

    return (
        etf_index_master.groupby(["instrument_type", "exchange_code"], dropna=False)
        .size()
        .reset_index(name="record_count")
    )


def build_domestic_master_with_etf(domestic_master, etf_index_master):
    combined = pd.concat(
        [
            domestic_master.reindex(columns=MASTER_COLUMNS).fillna(""),
            etf_index_master.reindex(columns=MASTER_COLUMNS).fillna(""),
        ],
        ignore_index=True,
    )
    combined["exchange_code"] = combined["exchange_code"].apply(normalize_exchange)
    combined = combined.drop_duplicates(
        subset=["instrument_type", "product_code", "exchange_code"], keep="first"
    )
    return combined.reindex(columns=MASTER_COLUMNS)


def main():
    etf_options_raw = read_raw_csv(ETF_OPTIONS_INPUT)
    fund_basic_raw = read_raw_csv(FUND_BASIC_INPUT)
    index_basic_raw = read_raw_csv(INDEX_BASIC_INPUT)
    domestic_master = read_raw_csv(DOMESTIC_MASTER_INPUT)

    etf_options_master = build_etf_option_master(etf_options_raw)
    etf_spot_master = build_etf_spot_master(fund_basic_raw, etf_options_master)
    index_spot_master = build_index_spot_master(index_basic_raw)
    etf_index_master = pd.concat(
        [etf_options_master, etf_spot_master, index_spot_master],
        ignore_index=True,
    ).reindex(columns=MASTER_COLUMNS)

    domestic_master_with_etf = build_domestic_master_with_etf(
        domestic_master, etf_index_master
    )
    underlying_check = build_underlying_check(etf_options_master, etf_spot_master)
    summary = build_summary(etf_index_master)

    ETF_INDEX_MASTER_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    etf_index_master.to_csv(
        ETF_INDEX_MASTER_OUTPUT, index=False, encoding="utf-8-sig"
    )
    domestic_master_with_etf.to_csv(
        DOMESTIC_MASTER_WITH_ETF_OUTPUT, index=False, encoding="utf-8-sig"
    )
    underlying_check.to_csv(
        UNDERLYING_CHECK_OUTPUT, index=False, encoding="utf-8-sig"
    )
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    unmatched_count = int((underlying_check["matched_etf_spot"] == False).sum())

    print("etf index master built successfully")
    print(f"etf_index_master rows: {len(etf_index_master)}")
    print(f"domestic_master_with_etf rows: {len(domestic_master_with_etf)}")
    print(f"etf option underlying unmatched: {unmatched_count}")
    print(f"summary saved to {SUMMARY_OUTPUT}")


if __name__ == "__main__":
    main()
