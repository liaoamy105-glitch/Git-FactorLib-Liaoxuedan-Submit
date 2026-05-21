from pathlib import Path

import pandas as pd


INPUT_PATH = Path("data/manual/exchange_rule_mapping_template_priority.csv")
OUTPUT_PATH = Path("data/manual/exchange_rule_mapping_template_priority_clean.csv")
SUMMARY_OUTPUT = Path("data/processed/rule_gap_priority_clean_summary.csv")
CHECK_OUTPUT = Path("data/processed/rule_gap_priority_clean_check.csv")

CONTINUOUS_KEYWORDS = ["主力", "连续", "当月", "次月", "当季", "下季", "上季", "主连"]
INVALID_CARRY_VALUES = {"", "待核验", "不适用", "nan", "None"}

TYPE_ORDER = {
    "FUTURE": 0,
    "OPTION": 1,
    "STOCK_INDEX_OPTION": 2,
    "ETF_OPTION": 3,
    "FOREIGN_FUTURE": 4,
}

OUTPUT_COLUMNS = [
    "品种类型",
    "交易所代码",
    "交易所名称",
    "品种代码",
    "品种名称",
    "合约行数",
    "优先级",
    "合约月份规则",
    "合约数量规则",
    "合约数量规则标准化代码",
    "到期日规则",
    "最后交易日规则",
    "是否节假日顺延",
    "行权方式",
    "交割方式",
    "规则来源",
    "核验状态",
    "备注",
]


def read_csv(path, required=False):
    if not path.exists():
        if required:
            raise FileNotFoundError(f"Required input file not found: {path}")
        return pd.DataFrame()
    try:
        return pd.read_csv(path, dtype=str).fillna("")
    except pd.errors.EmptyDataError:
        return pd.DataFrame()


def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_code(value):
    return clean_text(value).upper()


def first_non_empty(series):
    for value in series.tolist():
        text = clean_text(value)
        if text:
            return text
    return ""


def to_number(value):
    number = pd.to_numeric(value, errors="coerce")
    return 0 if pd.isna(number) else int(number)


def has_continuous_keyword(name):
    text = clean_text(name)
    return any(keyword in text for keyword in CONTINUOUS_KEYWORDS)


def representative_name(group):
    candidates = group[~group["品种名称"].apply(has_continuous_keyword)].copy()
    if candidates.empty:
        candidates = group.copy()

    candidates["_contract_count_num"] = candidates["合约行数"].apply(to_number)
    candidates["_name_len"] = candidates["品种名称"].apply(lambda value: len(clean_text(value)))
    selected = candidates.sort_values(
        ["_contract_count_num", "_name_len"], ascending=[False, True]
    ).iloc[0]
    return clean_text(selected["品种名称"])


def most_common_valid(series):
    values = [
        clean_text(value)
        for value in series.tolist()
        if clean_text(value) not in INVALID_CARRY_VALUES
    ]
    if not values:
        return ""
    return pd.Series(values).value_counts().index[0]


def sort_template(template):
    if template.empty:
        return template
    sorted_template = template.copy()
    sorted_template["_type_order"] = (
        sorted_template["品种类型"].apply(normalize_code).map(TYPE_ORDER).fillna(9)
    )
    sorted_template = sorted_template.sort_values(
        ["_type_order", "交易所代码", "品种代码"]
    )
    return sorted_template.drop(columns=["_type_order"])


def build_clean_template(priority_template):
    rows = []
    group_columns = ["品种类型", "交易所代码", "品种代码"]

    for keys, group in priority_template.groupby(group_columns, dropna=False):
        instrument_type, exchange_code, product_code = keys
        contract_count = int(group["合约行数"].apply(to_number).sum())
        rows.append(
            {
                "品种类型": instrument_type,
                "交易所代码": exchange_code,
                "交易所名称": first_non_empty(group["交易所名称"]),
                "品种代码": product_code,
                "品种名称": representative_name(group),
                "合约行数": contract_count,
                "优先级": "高",
                "合约月份规则": "",
                "合约数量规则": "",
                "合约数量规则标准化代码": "",
                "到期日规则": "",
                "最后交易日规则": "",
                "是否节假日顺延": "",
                "行权方式": most_common_valid(group["行权方式"]),
                "交割方式": most_common_valid(group["交割方式"]),
                "规则来源": "",
                "核验状态": "待人工核验",
                "备注": "已按品种类型+交易所代码+品种代码聚合，主力/连续行情口径已合并",
            }
        )

    return sort_template(pd.DataFrame(rows, columns=OUTPUT_COLUMNS))


def add_check(checks, check_item, passed, message):
    checks.append(
        {
            "check_item": check_item,
            "status": "passed" if passed else "failed",
            "message": message,
        }
    )


def build_checks(priority_template, clean_template):
    checks = []
    add_check(
        checks,
        "input_priority_template_exists",
        INPUT_PATH.exists(),
        str(INPUT_PATH),
    )
    add_check(
        checks,
        "output_clean_template_created",
        OUTPUT_PATH.exists(),
        str(OUTPUT_PATH),
    )
    duplicated = clean_template.duplicated(
        subset=["品种类型", "交易所代码", "品种代码"], keep=False
    )
    add_check(
        checks,
        "no_duplicate_by_type_exchange_code",
        not duplicated.any(),
        "no duplicate by 品种类型 + 交易所代码 + 品种代码"
        if not duplicated.any()
        else f"duplicated rows: {int(duplicated.sum())}",
    )
    continuous_names = clean_template[
        clean_template["品种名称"].apply(has_continuous_keyword)
    ]
    add_check(
        checks,
        "no_continuous_name_in_clean_template",
        continuous_names.empty,
        "no continuous/main contract names"
        if continuous_names.empty
        else "continuous names: " + "；".join(continuous_names["品种名称"].tolist()),
    )
    add_check(
        checks,
        "clean_rows_less_than_input_rows",
        len(clean_template) < len(priority_template),
        f"input rows: {len(priority_template)}, clean rows: {len(clean_template)}",
    )
    missing_columns = [column for column in OUTPUT_COLUMNS if column not in clean_template.columns]
    add_check(
        checks,
        "required_columns_exist",
        not missing_columns,
        "required columns exist"
        if not missing_columns
        else "missing columns: " + "；".join(missing_columns),
    )
    return pd.DataFrame(checks, columns=["check_item", "status", "message"])


def build_summary(priority_template, clean_template, check_report):
    type_counts = clean_template["品种类型"].value_counts()
    return pd.DataFrame(
        [
            {"metric": "input_priority_rows", "value": len(priority_template)},
            {"metric": "clean_priority_rows", "value": len(clean_template)},
            {
                "metric": "removed_duplicate_or_continuous_rows",
                "value": len(priority_template) - len(clean_template),
            },
            {"metric": "future_rows", "value": int(type_counts.get("FUTURE", 0))},
            {"metric": "option_rows", "value": int(type_counts.get("OPTION", 0))},
            {
                "metric": "stock_index_option_rows",
                "value": int(type_counts.get("STOCK_INDEX_OPTION", 0)),
            },
            {"metric": "etf_option_rows", "value": int(type_counts.get("ETF_OPTION", 0))},
            {
                "metric": "foreign_future_rows",
                "value": int(type_counts.get("FOREIGN_FUTURE", 0)),
            },
            {
                "metric": "check_failed_count",
                "value": int((check_report["status"] == "failed").sum()),
            },
        ],
        columns=["metric", "value"],
    )


def main():
    priority_template = read_csv(INPUT_PATH, required=True)
    clean_template = build_clean_template(priority_template)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    SUMMARY_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    clean_template.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

    check_report = build_checks(priority_template, clean_template)
    summary = build_summary(priority_template, clean_template, check_report)
    check_report.to_csv(CHECK_OUTPUT, index=False, encoding="utf-8-sig")
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    removed_rows = len(priority_template) - len(clean_template)

    print("priority rule template cleaned successfully")
    print(f"input rows: {len(priority_template)}")
    print(f"clean rows: {len(clean_template)}")
    print(f"removed rows: {removed_rows}")
    print(f"clean template saved to {OUTPUT_PATH}")
    print(f"check report saved to {CHECK_OUTPUT}")


if __name__ == "__main__":
    main()
