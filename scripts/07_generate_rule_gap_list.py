from pathlib import Path

import pandas as pd


DETAIL_SHEET_INPUT = Path("data/processed/detail_sheet.csv")
MAPPING_SHEET_INPUT = Path("data/processed/mapping_sheet.csv")

RULE_GAP_LIST_OUTPUT = Path("data/processed/rule_gap_list.csv")
TEMPLATE_OUTPUT = Path("data/manual/exchange_rule_mapping_template.csv")
PRIORITY_TEMPLATE_OUTPUT = Path(
    "data/manual/exchange_rule_mapping_template_priority.csv"
)
SUMMARY_OUTPUT = Path("data/processed/rule_gap_summary.csv")
CHECK_OUTPUT = Path("data/processed/rule_gap_check.csv")

RULE_FIELDS = [
    "合约月份规则",
    "合约数量规则",
    "合约数量规则标准化代码",
    "到期日规则",
    "最后交易日规则",
]

CURRENT_VALUE_FIELDS = {
    "合约月份规则": "合约月份规则当前值",
    "合约数量规则": "合约数量规则当前值",
    "合约数量规则标准化代码": "合约数量规则标准化代码当前值",
    "到期日规则": "到期日规则当前值",
    "最后交易日规则": "最后交易日规则当前值",
    "行权方式": "行权方式当前值",
    "交割方式": "交割方式当前值",
}

GAP_KEYWORDS = [
    "待核验",
    "待交易所核验",
    "具体规则待核验",
    "Tushare日期字段自动获取",
    "Tushare期货基础信息自动获取",
    "FUTURE_RULE_PENDING",
    "FUTURE_OPTION_RULE_PENDING",
    "FOREIGN_RULE_PENDING",
    "PENDING",
    "nan",
    "None",
    "空",
]

CORE_PRODUCT_CODES = {
    "AU",
    "AG",
    "CU",
    "AL",
    "ZN",
    "PB",
    "NI",
    "SN",
    "RU",
    "NR",
    "SC",
    "FU",
    "LU",
    "RB",
    "HC",
    "C",
    "M",
    "Y",
    "P",
    "A",
    "B",
    "CF",
    "SR",
    "TA",
    "MA",
    "OI",
    "RM",
    "IF",
    "IH",
    "IC",
    "IM",
}

PRIORITY_ORDER = {"高": 0, "中": 1, "低": 2}
TEMPLATE_TYPE_ORDER = {
    "FUTURE": 0,
    "OPTION": 1,
    "STOCK_INDEX_OPTION": 2,
    "ETF_OPTION": 3,
    "FOREIGN_FUTURE": 4,
}

RULE_GAP_COLUMNS = [
    "品种类型",
    "交易所代码",
    "交易所名称",
    "品种代码",
    "品种名称",
    "合约行数",
    "合约月份规则当前值",
    "合约数量规则当前值",
    "合约数量规则标准化代码当前值",
    "到期日规则当前值",
    "最后交易日规则当前值",
    "行权方式当前值",
    "交割方式当前值",
    "是否需要补规则",
    "缺失字段",
    "优先级",
    "是否出现在映射表",
    "备注",
]

TEMPLATE_COLUMNS = [
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


def normalize_detail_columns(detail_sheet):
    detail = detail_sheet.copy()
    if "交易所代码" not in detail.columns and "交易所缩写" in detail.columns:
        detail["交易所代码"] = detail["交易所缩写"]
    if "交易所名称" not in detail.columns and "交易所全称" in detail.columns:
        detail["交易所名称"] = detail["交易所全称"]

    required_columns = [
        "品种类型",
        "交易所代码",
        "交易所名称",
        "品种代码",
        "品种名称",
        "行权方式",
        "交割方式",
    ] + RULE_FIELDS
    for column in required_columns:
        if column not in detail.columns:
            detail[column] = ""
    return detail


def representative_value(series):
    values = [clean_text(value) for value in series.tolist() if clean_text(value)]
    if not values:
        return ""
    counts = pd.Series(values).value_counts()
    top_values = counts.index.tolist()[:3]
    return "；".join(top_values)


def needs_fill(value):
    text = clean_text(value)
    if not text:
        return True
    return any(keyword.lower() in text.lower() for keyword in GAP_KEYWORDS)


def mapping_related(mapping_text, product_code, product_name):
    code = clean_text(product_code)
    name = clean_text(product_name)
    return bool((code and code in mapping_text) or (name and name in mapping_text))


def priority_for(row):
    instrument_type = normalize_code(row["品种类型"])
    product_code = normalize_code(row["品种代码"])

    if not row["是否需要补规则"]:
        return "低"

    if instrument_type == "FOREIGN_FUTURE":
        return "低"

    if instrument_type == "STOCK_INDEX_OPTION":
        return "高"

    if instrument_type == "ETF_OPTION":
        return "高"

    if instrument_type in {"FUTURE", "OPTION"} and product_code in CORE_PRODUCT_CODES:
        return "高"

    if (
        instrument_type in {"FUTURE", "OPTION"}
        and row["是否出现在映射表"]
        and product_code in CORE_PRODUCT_CODES
    ):
        return "高"

    if instrument_type in {"FUTURE", "OPTION"}:
        return "中"

    return "低"


def build_rule_gap_list(detail_sheet, mapping_sheet):
    detail = normalize_detail_columns(detail_sheet)
    mapping_text = " ".join(
        mapping_sheet.astype(str).fillna("").agg(" ".join, axis=1).tolist()
    )

    rows = []
    group_columns = ["品种类型", "交易所代码", "品种代码", "品种名称"]
    for keys, group in detail.groupby(group_columns, dropna=False):
        instrument_type, exchange_code, product_code, product_name = keys
        row = {
            "品种类型": instrument_type,
            "交易所代码": exchange_code,
            "交易所名称": representative_value(group["交易所名称"]),
            "品种代码": product_code,
            "品种名称": product_name,
            "合约行数": len(group),
        }

        for source_field, output_field in CURRENT_VALUE_FIELDS.items():
            row[output_field] = representative_value(group[source_field])

        missing_fields = []
        instrument_type_key = normalize_code(instrument_type)
        for field in RULE_FIELDS:
            current_value = row[CURRENT_VALUE_FIELDS[field]]
            if instrument_type_key in {"ETF_SPOT", "INDEX_SPOT"} and current_value == "不适用":
                continue
            if needs_fill(current_value):
                missing_fields.append(field)

        row["是否需要补规则"] = bool(missing_fields)
        row["缺失字段"] = "；".join(missing_fields)
        row["是否出现在映射表"] = mapping_related(mapping_text, product_code, product_name)
        row["优先级"] = priority_for(row)
        row["备注"] = (
            "需要按交易所合约细则补充规则"
            if row["是否需要补规则"]
            else "规则字段已基本完整或不适用"
        )
        rows.append(row)

    rule_gap = pd.DataFrame(rows, columns=RULE_GAP_COLUMNS)
    return sort_rule_gap(rule_gap)


def safe_template_value(value):
    return "" if needs_fill(value) else clean_text(value)


def build_template(rule_gap):
    need_fill = rule_gap[rule_gap["是否需要补规则"] == True].copy()
    template = pd.DataFrame(
        {
            "品种类型": need_fill["品种类型"],
            "交易所代码": need_fill["交易所代码"],
            "交易所名称": need_fill["交易所名称"],
            "品种代码": need_fill["品种代码"],
            "品种名称": need_fill["品种名称"],
            "合约行数": need_fill["合约行数"],
            "优先级": need_fill["优先级"],
            "合约月份规则": "",
            "合约数量规则": "",
            "合约数量规则标准化代码": "",
            "到期日规则": "",
            "最后交易日规则": "",
            "是否节假日顺延": "",
            "行权方式": need_fill["行权方式当前值"].apply(safe_template_value),
            "交割方式": need_fill["交割方式当前值"].apply(safe_template_value),
            "规则来源": "",
            "核验状态": "待人工核验",
            "备注": "请按交易所官网合约细则补充，不要按单个合约逐行填写",
        },
        columns=TEMPLATE_COLUMNS,
    )
    return sort_rule_gap(template)


def sort_priority_template(template):
    if template.empty:
        return template
    sorted_template = template.copy()
    sorted_template["_type_order"] = (
        sorted_template["品种类型"].apply(normalize_code).map(TEMPLATE_TYPE_ORDER).fillna(9)
    )
    sorted_template = sorted_template.sort_values(
        ["_type_order", "交易所代码", "品种代码"]
    )
    return sorted_template.drop(columns=["_type_order"])


def build_priority_template(rule_gap):
    high_need_fill = rule_gap[
        (rule_gap["是否需要补规则"] == True) & (rule_gap["优先级"] == "高")
    ].copy()
    template = build_template(high_need_fill)
    return sort_priority_template(template)


def sort_rule_gap(data):
    if data.empty:
        return data
    sorted_data = data.copy()
    sorted_data["_priority_order"] = sorted_data["优先级"].map(PRIORITY_ORDER).fillna(9)
    if "是否出现在映射表" in sorted_data.columns:
        sorted_data["_mapping_order"] = sorted_data["是否出现在映射表"].map(
            {True: 0, False: 1, "True": 0, "False": 1}
        ).fillna(1)
    else:
        sorted_data["_mapping_order"] = 0
    sorted_data = sorted_data.sort_values(
        ["_priority_order", "_mapping_order", "品种类型", "交易所代码", "品种代码"]
    )
    return sorted_data.drop(columns=["_priority_order", "_mapping_order"])


def build_summary(detail_sheet, rule_gap, priority_template):
    need_fill = rule_gap[rule_gap["是否需要补规则"] == True]
    return pd.DataFrame(
        [
            {"metric": "detail_sheet_rows", "value": len(detail_sheet)},
            {"metric": "rule_unit_count", "value": len(rule_gap)},
            {"metric": "need_rule_fill_count", "value": len(need_fill)},
            {
                "metric": "need_fill_high_priority_count",
                "value": int((need_fill["优先级"] == "高").sum()),
            },
            {
                "metric": "need_fill_medium_priority_count",
                "value": int((need_fill["优先级"] == "中").sum()),
            },
            {
                "metric": "need_fill_low_priority_count",
                "value": int((need_fill["优先级"] == "低").sum()),
            },
            {
                "metric": "mapping_related_need_fill_count",
                "value": int(
                    (
                        (rule_gap["是否需要补规则"] == True)
                        & (rule_gap["是否出现在映射表"] == True)
                    ).sum()
                ),
            },
            {"metric": "priority_template_rows", "value": len(priority_template)},
        ],
        columns=["metric", "value"],
    )


def add_check(checks, check_item, passed, message):
    checks.append(
        {
            "check_item": check_item,
            "status": "passed" if passed else "failed",
            "message": message,
        }
    )


def build_checks(detail_sheet, mapping_sheet, rule_gap, template, priority_template):
    checks = []
    add_check(
        checks,
        "detail_sheet_exists",
        DETAIL_SHEET_INPUT.exists(),
        str(DETAIL_SHEET_INPUT),
    )
    add_check(
        checks,
        "mapping_sheet_exists",
        MAPPING_SHEET_INPUT.exists(),
        str(MAPPING_SHEET_INPUT),
    )
    add_check(
        checks,
        "rule_gap_list_created",
        not rule_gap.empty or detail_sheet.empty,
        f"rule gap rows: {len(rule_gap)}",
    )
    add_check(
        checks,
        "exchange_rule_mapping_template_created",
        TEMPLATE_OUTPUT.exists(),
        str(TEMPLATE_OUTPUT),
    )
    add_check(
        checks,
        "no_empty_product_code_in_rule_gap",
        not (rule_gap["品种代码"].apply(clean_text) == "").any(),
        "no empty product code",
    )
    add_check(
        checks,
        "no_empty_exchange_code_in_rule_gap",
        not (rule_gap["交易所代码"].apply(clean_text) == "").any(),
        "no empty exchange code",
    )
    add_check(
        checks,
        "high_priority_items_exist",
        (rule_gap["优先级"] == "高").any(),
        f"high priority rows: {int((rule_gap['优先级'] == '高').sum())}",
    )
    need_fill_count = int((rule_gap["是否需要补规则"] == True).sum())
    add_check(
        checks,
        "output_template_rows_equal_need_fill_rows",
        len(template) == need_fill_count,
        f"template rows: {len(template)}, need fill rows: {need_fill_count}",
    )
    priority_template_readable = False
    if PRIORITY_TEMPLATE_OUTPUT.exists():
        try:
            read_csv(PRIORITY_TEMPLATE_OUTPUT)
            priority_template_readable = True
        except Exception:
            priority_template_readable = False
    add_check(
        checks,
        "priority_template_created",
        priority_template_readable,
        str(PRIORITY_TEMPLATE_OUTPUT),
    )
    priority_only_high = priority_template.empty or (
        priority_template["优先级"] == "高"
    ).all()
    add_check(
        checks,
        "priority_template_only_high_need_fill",
        priority_only_high,
        f"priority template rows: {len(priority_template)}",
    )
    foreign_high_count = int(
        (
            (rule_gap["品种类型"].apply(normalize_code) == "FOREIGN_FUTURE")
            & (rule_gap["优先级"] == "高")
        ).sum()
    )
    add_check(
        checks,
        "foreign_future_not_high_priority_by_default",
        foreign_high_count == 0,
        f"FOREIGN_FUTURE high priority rows: {foreign_high_count}",
    )
    return pd.DataFrame(checks, columns=["check_item", "status", "message"])


def main():
    detail_sheet = read_csv(DETAIL_SHEET_INPUT, required=True)
    mapping_sheet = read_csv(MAPPING_SHEET_INPUT, required=True)

    rule_gap = build_rule_gap_list(detail_sheet, mapping_sheet)
    template = build_template(rule_gap)
    priority_template = build_priority_template(rule_gap)
    summary = build_summary(detail_sheet, rule_gap, priority_template)

    RULE_GAP_LIST_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    TEMPLATE_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    rule_gap.to_csv(RULE_GAP_LIST_OUTPUT, index=False, encoding="utf-8-sig")
    template.to_csv(TEMPLATE_OUTPUT, index=False, encoding="utf-8-sig")
    priority_template.to_csv(
        PRIORITY_TEMPLATE_OUTPUT, index=False, encoding="utf-8-sig"
    )

    checks = build_checks(
        detail_sheet, mapping_sheet, rule_gap, template, priority_template
    )
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")
    checks.to_csv(CHECK_OUTPUT, index=False, encoding="utf-8-sig")

    need_fill_count = int((rule_gap["是否需要补规则"] == True).sum())
    need_fill = rule_gap[rule_gap["是否需要补规则"] == True]
    high_priority_count = int((need_fill["优先级"] == "高").sum())
    medium_priority_count = int((need_fill["优先级"] == "中").sum())
    low_priority_count = int((need_fill["优先级"] == "低").sum())

    print("rule gap list generated successfully")
    print(f"rule units: {len(rule_gap)}")
    print(f"need rule fill: {need_fill_count}")
    print(f"high priority need fill: {high_priority_count}")
    print(f"medium priority need fill: {medium_priority_count}")
    print(f"low priority need fill: {low_priority_count}")
    print(f"priority template saved to {PRIORITY_TEMPLATE_OUTPUT}")
    print(f"full template saved to {TEMPLATE_OUTPUT}")


if __name__ == "__main__":
    main()
