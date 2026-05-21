from pathlib import Path

import pandas as pd


MASTER_INPUT = Path("data/processed/domestic_master_with_etf.csv")
MAPPING_OUTPUT = Path("data/processed/domestic_cross_market_mapping.csv")
CHECK_OUTPUT = Path("data/processed/domestic_cross_market_mapping_check.csv")
SUMMARY_OUTPUT = Path("data/processed/domestic_cross_market_mapping_summary.csv")

SOURCE = "manual_domestic_cross_market_rule"
NOTE = "国内跨市场映射规则，人工指定"

MAPPING_ORDER = {
    "FUTURE": 1,
    "STOCK_INDEX_OPTION": 2,
    "INDEX_SPOT": 3,
    "ETF_SPOT": 4,
    "ETF_OPTION": 5,
}

CROSS_MARKET_RULES = [
    {
        "base_group": "上证50",
        "items": [
            {"instrument_type": "FUTURE", "product_code": "IH", "exchange_code": "CFFEX"},
            {
                "instrument_type": "STOCK_INDEX_OPTION",
                "product_code": "HO",
                "exchange_code": "CFFEX",
            },
            {
                "instrument_type": "INDEX_SPOT",
                "product_code": "000016.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "510050.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP510050.SH",
                "exchange_code": "SSE",
            },
        ],
    },
    {
        "base_group": "沪深300",
        "items": [
            {"instrument_type": "FUTURE", "product_code": "IF", "exchange_code": "CFFEX"},
            {
                "instrument_type": "STOCK_INDEX_OPTION",
                "product_code": "IO",
                "exchange_code": "CFFEX",
            },
            {
                "instrument_type": "INDEX_SPOT",
                "product_code": "000300.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "510300.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "159919.SZ",
                "exchange_code": "SZSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP510300.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP159919.SZ",
                "exchange_code": "SZSE",
            },
        ],
    },
    {
        "base_group": "中证500",
        "items": [
            {"instrument_type": "FUTURE", "product_code": "IC", "exchange_code": "CFFEX"},
            {
                "instrument_type": "INDEX_SPOT",
                "product_code": "000905.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "510500.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "159922.SZ",
                "exchange_code": "SZSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP510500.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP159922.SZ",
                "exchange_code": "SZSE",
            },
        ],
    },
    {
        "base_group": "中证1000",
        "items": [
            {"instrument_type": "FUTURE", "product_code": "IM", "exchange_code": "CFFEX"},
            {
                "instrument_type": "STOCK_INDEX_OPTION",
                "product_code": "MO",
                "exchange_code": "CFFEX",
            },
            {
                "instrument_type": "INDEX_SPOT",
                "product_code": "000852.SH",
                "exchange_code": "SSE",
            },
        ],
    },
    {
        "base_group": "深证100",
        "items": [
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "159901.SZ",
                "exchange_code": "SZSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP159901.SZ",
                "exchange_code": "SZSE",
            },
        ],
    },
    {
        "base_group": "创业板",
        "items": [
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "159915.SZ",
                "exchange_code": "SZSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP159915.SZ",
                "exchange_code": "SZSE",
            },
        ],
    },
    {
        "base_group": "科创50",
        "items": [
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "588000.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_SPOT",
                "product_code": "588080.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP588000.SH",
                "exchange_code": "SSE",
            },
            {
                "instrument_type": "ETF_OPTION",
                "product_code": "OP588080.SH",
                "exchange_code": "SSE",
            },
        ],
    },
]

MAPPING_COLUMNS = [
    "base_group",
    "instrument_type",
    "product_code",
    "product_name",
    "exchange_code",
    "exchange_name",
    "mapping_role",
    "mapping_order",
    "source",
    "note",
]

CHECK_COLUMNS = [
    "base_group",
    "instrument_type",
    "product_code",
    "exchange_code",
    "matched_master",
    "message",
]


def read_csv(path):
    if not path.exists():
        raise FileNotFoundError(f"Required input file not found: {path}")
    return pd.read_csv(path, dtype=str).fillna("")


def normalize_exchange(value):
    return str(value).strip().upper()


def master_key(row):
    return (
        str(row["instrument_type"]).strip(),
        str(row["product_code"]).strip(),
        normalize_exchange(row["exchange_code"]),
    )


def build_master_lookup(master):
    lookup = {}
    for _, row in master.iterrows():
        key = master_key(row)
        if key not in lookup:
            lookup[key] = row
    return lookup


def iter_rules():
    for group in CROSS_MARKET_RULES:
        base_group = group["base_group"]
        for item in group["items"]:
            yield {
                "base_group": base_group,
                "instrument_type": item["instrument_type"],
                "product_code": item["product_code"],
                "exchange_code": normalize_exchange(item["exchange_code"]),
            }


def build_mapping_and_check(master_lookup):
    mapping_rows = []
    check_rows = []

    for rule in iter_rules():
        key = (
            rule["instrument_type"],
            rule["product_code"],
            rule["exchange_code"],
        )
        matched_row = master_lookup.get(key)
        matched = matched_row is not None

        mapping_rows.append(
            {
                "base_group": rule["base_group"],
                "instrument_type": rule["instrument_type"],
                "product_code": rule["product_code"],
                "product_name": matched_row["product_name"] if matched else "",
                "exchange_code": rule["exchange_code"],
                "exchange_name": matched_row["exchange_name"] if matched else "",
                "mapping_role": rule["instrument_type"],
                "mapping_order": MAPPING_ORDER[rule["instrument_type"]],
                "source": SOURCE,
                "note": NOTE,
            }
        )

        check_rows.append(
            {
                "base_group": rule["base_group"],
                "instrument_type": rule["instrument_type"],
                "product_code": rule["product_code"],
                "exchange_code": rule["exchange_code"],
                "matched_master": matched,
                "message": "found in domestic_master_with_etf"
                if matched
                else "not found in domestic_master_with_etf, need confirm or add to master",
            }
        )

    return (
        pd.DataFrame(mapping_rows, columns=MAPPING_COLUMNS),
        pd.DataFrame(check_rows, columns=CHECK_COLUMNS),
    )


def build_summary(mapping):
    return (
        mapping.groupby(["base_group", "mapping_role"], dropna=False)
        .size()
        .reset_index(name="record_count")
    )


def main():
    master = read_csv(MASTER_INPUT)
    master["exchange_code"] = master["exchange_code"].apply(normalize_exchange)
    master_lookup = build_master_lookup(master)

    mapping, check = build_mapping_and_check(master_lookup)
    summary = build_summary(mapping)

    MAPPING_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    mapping.to_csv(MAPPING_OUTPUT, index=False, encoding="utf-8-sig")
    check.to_csv(CHECK_OUTPUT, index=False, encoding="utf-8-sig")
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    unmatched_count = int((check["matched_master"] == False).sum())

    print("domestic cross market mapping built successfully")
    print(f"domestic cross market mapping rows: {len(mapping)}")
    print(f"unmatched rules: {unmatched_count}")
    print(f"summary saved to {SUMMARY_OUTPUT}")


if __name__ == "__main__":
    main()
