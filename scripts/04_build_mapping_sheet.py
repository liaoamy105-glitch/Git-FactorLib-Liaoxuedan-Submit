from pathlib import Path

import pandas as pd


MASTER_INPUT = Path("data/processed/domestic_master_with_etf.csv")
FINAL_SCOPE_INPUT = Path("data/processed/final_product_scope.csv")
FOREIGN_MAPPING_INPUT = Path("data/processed/effective_foreign_mapping.csv")
CROSS_MARKET_INPUT = Path("data/processed/domestic_cross_market_mapping.csv")

MAPPING_SHEET_OUTPUT = Path("data/processed/mapping_sheet.csv")
CHECK_OUTPUT = Path("data/processed/mapping_sheet_check.csv")
SUMMARY_OUTPUT = Path("data/processed/mapping_sheet_summary.csv")

EMPTY = "空"
NOT_APPLICABLE = "不适用"
DATA_SOURCE = "Tushare + 人工映射规则"
VERIFY_STATUS = "初版待核验"

MAPPING_SHEET_COLUMNS = [
    "期货",
    "商品名称",
    "期权",
    "ETF期权",
    "ETF现货",
    "现货/指数",
    "国内跨市场映射标的",
    "国外交易所A期货1",
    "国外交易所B期货2",
    "国外交易所C期货3",
    "备注",
    "数据来源",
    "核验状态",
]


def read_csv(path):
    if not path.exists():
        raise FileNotFoundError(f"Required input file not found: {path}")
    return pd.read_csv(path, dtype=str).fillna("")


def normalize_code(value):
    return str(value).strip().upper()


def clean_text(value):
    return str(value).strip()


def parse_bool(value):
    return str(value).strip().lower() == "true"


def sort_mapping_order(mapping):
    ordered = mapping.copy()
    ordered["_mapping_order_num"] = pd.to_numeric(
        ordered["mapping_order"], errors="coerce"
    )
    return ordered.sort_values(
        ["_mapping_order_num", "mapping_order"], na_position="last"
    )


def display_value(value):
    return clean_text(value) if clean_text(value) else EMPTY


def target_label(name, code, exchange_name, country="中国"):
    if not any(clean_text(value) for value in [name, code, exchange_name, country]):
        return EMPTY
    return (
        f"{display_value(name)}({display_value(code)})_"
        f"{display_value(exchange_name)}_{display_value(country)}"
    )


def master_label(row):
    if row is None:
        return EMPTY
    return target_label(
        row.get("product_name", ""),
        row.get("product_code", ""),
        row.get("exchange_name", ""),
        row.get("country", "中国"),
    )


def cross_market_label(row):
    return target_label(
        row.get("product_name", ""),
        row.get("product_code", ""),
        row.get("exchange_name", ""),
        "中国",
    )


def foreign_label(row):
    return target_label(
        row.get("foreign_product_name", ""),
        row.get("foreign_product_code", ""),
        row.get("foreign_exchange_name", ""),
        row.get("foreign_country", ""),
    )


def join_labels(labels, empty_value=EMPTY):
    cleaned = [label for label in labels if clean_text(label) and label != EMPTY]
    return "；".join(cleaned) if cleaned else empty_value


def blank_row():
    return {column: EMPTY for column in MAPPING_SHEET_COLUMNS}


def build_master_lookups(master):
    normalized = master.copy()
    normalized["_instrument_type_key"] = normalized["instrument_type"].apply(
        normalize_code
    )
    normalized["_product_code_key"] = normalized["product_code"].apply(normalize_code)
    normalized["_exchange_code_key"] = normalized["exchange_code"].apply(
        normalize_code
    )
    return normalized


def master_match(master, instrument_type, product_code, exchange_code):
    matched = master[
        (master["_instrument_type_key"] == normalize_code(instrument_type))
        & (master["_product_code_key"] == normalize_code(product_code))
        & (master["_exchange_code_key"] == normalize_code(exchange_code))
    ]
    if matched.empty:
        return None
    return matched.iloc[0]


def find_commodity_options(master, future_code, future_exchange):
    return master[
        (master["_instrument_type_key"] == "OPTION")
        & (
            master["underlying_product_code"].apply(normalize_code)
            == normalize_code(future_code)
        )
        & (master["_exchange_code_key"] == normalize_code(future_exchange))
    ]


def cross_market_groups_for_future(cross_market, product_code, exchange_code):
    matched = cross_market[
        (cross_market["instrument_type"].apply(normalize_code) == "FUTURE")
        & (
            cross_market["product_code"].apply(normalize_code)
            == normalize_code(product_code)
        )
        & (
            cross_market["exchange_code"].apply(normalize_code)
            == normalize_code(exchange_code)
        )
    ]
    return matched["base_group"].drop_duplicates().tolist()


def rows_for_base_group(cross_market, base_group, instrument_type=None):
    rows = cross_market[cross_market["base_group"] == base_group]
    if instrument_type:
        rows = rows[rows["instrument_type"].apply(normalize_code) == instrument_type]
    return rows


def labels_for_type(group_rows, instrument_type):
    rows = group_rows[group_rows["instrument_type"].apply(normalize_code) == instrument_type]
    return [cross_market_label(row) for _, row in rows.iterrows()]


def base_groups_without_future(cross_market):
    groups = []
    for base_group, group in cross_market.groupby("base_group", sort=False):
        has_future = (group["instrument_type"].apply(normalize_code) == "FUTURE").any()
        if not has_future:
            groups.append(base_group)
    return groups


def fill_cross_market_fields(row, group_rows, future_code="", future_exchange=""):
    row["ETF期权"] = join_labels(labels_for_type(group_rows, "ETF_OPTION"))
    row["ETF现货"] = join_labels(labels_for_type(group_rows, "ETF_SPOT"))
    row["现货/指数"] = join_labels(labels_for_type(group_rows, "INDEX_SPOT"))

    other_rows = group_rows.copy()
    if future_code and future_exchange:
        other_rows = other_rows[
            ~(
                (other_rows["instrument_type"].apply(normalize_code) == "FUTURE")
                & (
                    other_rows["product_code"].apply(normalize_code)
                    == normalize_code(future_code)
                )
                & (
                    other_rows["exchange_code"].apply(normalize_code)
                    == normalize_code(future_exchange)
                )
            )
        ]

    row["国内跨市场映射标的"] = join_labels(
        [cross_market_label(item) for _, item in other_rows.iterrows()],
        empty_value=NOT_APPLICABLE,
    )


def fill_foreign_mapping_fields(row, foreign_rows):
    columns = ["国外交易所A期货1", "国外交易所B期货2", "国外交易所C期货3"]
    for column, (_, foreign) in zip(
        columns, sort_mapping_order(foreign_rows).head(3).iterrows()
    ):
        row[column] = foreign_label(foreign)


def build_future_based_rows(master, final_scope, foreign_mapping, cross_market):
    rows = []
    included = final_scope[final_scope["include_in_final"].apply(parse_bool)].copy()

    for _, scope in included.iterrows():
        future_code = scope["domestic_product_code"]
        future_exchange = scope["domestic_exchange_code"]
        future = master_match(master, "FUTURE", future_code, future_exchange)

        row = blank_row()
        if future is not None:
            future_name = future["product_name"]
            future_label = master_label(future)
        else:
            future_name = scope["domestic_product_name"]
            future_label = target_label(
                scope["domestic_product_name"],
                future_code,
                scope["domestic_exchange_name"],
                "中国",
            )

        row["期货"] = future_label
        row["商品名称"] = target_label(
            future_name,
            future_code,
            scope["domestic_exchange_name"],
            "中国",
        )

        commodity_options = find_commodity_options(master, future_code, future_exchange)
        row["期权"] = join_labels(
            [master_label(option) for _, option in commodity_options.iterrows()]
        )

        base_groups = cross_market_groups_for_future(
            cross_market, future_code, future_exchange
        )
        if base_groups:
            group_rows = pd.concat(
                [rows_for_base_group(cross_market, base_group) for base_group in base_groups],
                ignore_index=True,
            )
            stock_index_options = labels_for_type(group_rows, "STOCK_INDEX_OPTION")
            if stock_index_options:
                row["期权"] = join_labels(
                    [value for value in [row["期权"]] if value != EMPTY]
                    + stock_index_options
                )
            fill_cross_market_fields(row, group_rows, future_code, future_exchange)
        else:
            row["国内跨市场映射标的"] = NOT_APPLICABLE

        foreign_rows = foreign_mapping[
            (
                foreign_mapping["domestic_product_code"].apply(normalize_code)
                == normalize_code(future_code)
            )
            & (
                foreign_mapping["domestic_exchange_code"].apply(normalize_code)
                == normalize_code(future_exchange)
            )
        ]
        fill_foreign_mapping_fields(row, foreign_rows)

        policy = clean_text(scope.get("foreign_mapping_policy", ""))
        scope_note = clean_text(scope.get("scope_note", ""))
        row["备注"] = "；".join(value for value in [policy, scope_note] if value)
        row["数据来源"] = DATA_SOURCE
        row["核验状态"] = VERIFY_STATUS
        rows.append(row)

    return rows


def build_etf_index_based_rows(cross_market):
    rows = []
    for base_group in base_groups_without_future(cross_market):
        group_rows = rows_for_base_group(cross_market, base_group)
        row = blank_row()
        row["期货"] = EMPTY
        row["商品名称"] = EMPTY
        row["期权"] = EMPTY
        fill_cross_market_fields(row, group_rows)
        row["国外交易所A期货1"] = EMPTY
        row["国外交易所B期货2"] = EMPTY
        row["国外交易所C期货3"] = EMPTY
        row["备注"] = "ETF/指数独立映射组，无期货主标的。"
        row["数据来源"] = DATA_SOURCE
        row["核验状态"] = VERIFY_STATUS
        rows.append(row)
    return rows


def build_mapping_sheet(master, final_scope, foreign_mapping, cross_market):
    rows = build_future_based_rows(master, final_scope, foreign_mapping, cross_market)
    rows.extend(build_etf_index_based_rows(cross_market))
    return pd.DataFrame(rows, columns=MAPPING_SHEET_COLUMNS)


def add_check(checks, check_item, passed, message):
    checks.append(
        {
            "check_item": check_item,
            "status": "passed" if passed else "failed",
            "message": message,
        }
    )


def row_contains(row, values):
    text = " ".join(str(value) for value in row.fillna("").tolist())
    return all(value in text for value in values)


def find_mapping_row_by_future_code(mapping_sheet, future_code):
    matched = mapping_sheet[
        mapping_sheet["期货"].apply(lambda value: f"({normalize_code(future_code)})" in normalize_code(value))
    ]
    if matched.empty:
        return None
    return matched.iloc[0]


def find_pb_shfe_row(mapping_sheet):
    def is_pb_shfe_future(value):
        text = clean_text(value)
        return ("PB" in text and "上海期货交易所" in text) or (
            "沪铅" in text and "PB" in text
        )

    matched = mapping_sheet[mapping_sheet["期货"].apply(is_pb_shfe_future)]
    if matched.empty:
        return None
    return matched.iloc[0]


def pb_row_maps_to_lme(pb_row):
    if pb_row is None:
        return False

    for column in ["国外交易所A期货1", "国外交易所B期货2", "国外交易所C期货3"]:
        value = clean_text(pb_row.get(column, ""))
        has_lead = "铅" in value or "PB" in value
        has_lme = "LME" in value or "London Metal Exchange" in value
        if has_lead and has_lme:
            return True
    return False


def pb_foreign_mapping_debug_message(pb_row):
    if pb_row is None:
        return "PB/SHFE row not found in mapping_sheet"
    return (
        "PB foreign fields: "
        f"A={pb_row.get('国外交易所A期货1', '')}; "
        f"B={pb_row.get('国外交易所B期货2', '')}; "
        f"C={pb_row.get('国外交易所C期货3', '')}"
    )


def build_checks(mapping_sheet, foreign_mapping):
    checks = []

    pb_row = find_pb_shfe_row(mapping_sheet)
    pb_row_has_lme = pb_row_maps_to_lme(pb_row)
    add_check(
        checks,
        "PB_SHFE_maps_to_LME_PB",
        pb_row_has_lme,
        "PB/SHFE maps to LME PB"
        if pb_row_has_lme
        else pb_foreign_mapping_debug_message(pb_row),
    )

    chain_checks = [
        (
            "上证50_chain_contains_required_codes",
            "IH",
            ["IH", "HO", "000016.SH", "510050.SH", "OP510050.SH"],
        ),
        (
            "沪深300_chain_contains_required_codes",
            "IF",
            [
                "IF",
                "IO",
                "000300.SH",
                "510300.SH",
                "159919.SZ",
                "OP510300.SH",
                "OP159919.SZ",
            ],
        ),
        (
            "中证500_chain_contains_required_codes",
            "IC",
            [
                "IC",
                "000905.SH",
                "510500.SH",
                "159922.SZ",
                "OP510500.SH",
                "OP159922.SZ",
            ],
        ),
        (
            "中证1000_chain_contains_required_codes",
            "IM",
            ["IM", "MO", "000852.SH"],
        ),
    ]
    for check_item, future_code, values in chain_checks:
        row = find_mapping_row_by_future_code(mapping_sheet, future_code)
        passed = row is not None and row_contains(row, values)
        add_check(checks, check_item, passed, f"required codes present: {', '.join(values)}")

    futures = mapping_sheet.loc[mapping_sheet["期货"] != EMPTY, "期货"]
    duplicated_futures = sorted(
        futures[futures.duplicated(keep=False)].dropna().unique().tolist()
    )
    add_check(
        checks,
        "future_primary_target_no_duplicates",
        not duplicated_futures,
        "no duplicated futures primary target"
        if not duplicated_futures
        else "duplicated futures primary target: " + "；".join(duplicated_futures),
    )

    unresolved_exists = mapping_sheet["备注"].apply(
        lambda value: "UNRESOLVED" in normalize_code(value)
    ).any()
    add_check(
        checks,
        "foreign_mapping_policy_unresolved_absent",
        not unresolved_exists,
        "UNRESOLVED does not exist",
    )

    return pd.DataFrame(checks, columns=["check_item", "status", "message"])


def build_summary(mapping_sheet, check_report):
    check_failed_count = int((check_report["status"] == "failed").sum())
    future_based_rows = int((mapping_sheet["期货"] != EMPTY).sum())
    etf_index_based_rows = len(mapping_sheet) - future_based_rows
    foreign_mapped_rows = int(
        (
            (mapping_sheet["国外交易所A期货1"] != EMPTY)
            | (mapping_sheet["国外交易所B期货2"] != EMPTY)
            | (mapping_sheet["国外交易所C期货3"] != EMPTY)
        ).sum()
    )
    no_foreign_mapping_rows = len(mapping_sheet) - foreign_mapped_rows

    return pd.DataFrame(
        [
            {"metric": "mapping_sheet_rows", "value": len(mapping_sheet)},
            {"metric": "future_based_rows", "value": future_based_rows},
            {"metric": "etf_index_based_rows", "value": etf_index_based_rows},
            {"metric": "foreign_mapped_rows", "value": foreign_mapped_rows},
            {"metric": "no_foreign_mapping_rows", "value": no_foreign_mapping_rows},
            {"metric": "check_failed_count", "value": check_failed_count},
        ],
        columns=["metric", "value"],
    )


def main():
    master = build_master_lookups(read_csv(MASTER_INPUT))
    final_scope = read_csv(FINAL_SCOPE_INPUT)
    foreign_mapping = read_csv(FOREIGN_MAPPING_INPUT)
    cross_market = read_csv(CROSS_MARKET_INPUT)
    cross_market["exchange_code"] = cross_market["exchange_code"].apply(normalize_code)

    mapping_sheet = build_mapping_sheet(master, final_scope, foreign_mapping, cross_market)
    check_report = build_checks(mapping_sheet, foreign_mapping)
    summary = build_summary(mapping_sheet, check_report)

    MAPPING_SHEET_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    mapping_sheet.to_csv(MAPPING_SHEET_OUTPUT, index=False, encoding="utf-8-sig")
    check_report.to_csv(CHECK_OUTPUT, index=False, encoding="utf-8-sig")
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    check_failed_count = int((check_report["status"] == "failed").sum())

    print("mapping sheet built successfully")
    print(f"mapping_sheet rows: {len(mapping_sheet)}")
    print(f"check failed count: {check_failed_count}")
    print(f"mapping sheet saved to {MAPPING_SHEET_OUTPUT}")
    print(f"check report saved to {CHECK_OUTPUT}")
    print(f"summary saved to {SUMMARY_OUTPUT}")


if __name__ == "__main__":
    main()
