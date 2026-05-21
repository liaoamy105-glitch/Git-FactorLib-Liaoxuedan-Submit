from pathlib import Path

import pandas as pd


INPUT_PATH = Path("data/manual/exchange_rule_mapping_template_priority_clean.csv")
OUTPUT_PATH = Path("data/manual/core_exchange_rule_seed.csv")
SUMMARY_OUTPUT = Path("data/processed/core_exchange_rule_seed_summary.csv")
CHECK_OUTPUT = Path("data/processed/core_exchange_rule_seed_check.csv")

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

INVALID_OR_EMPTY = {"", "待核验", "不适用", "nan", "None"}


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


def clear_value(value):
    text = clean_text(value)
    return text if text not in INVALID_OR_EMPTY else ""


def original_or_default(row, column, default):
    return clear_value(row.get(column, "")) or default


def base_row(row):
    return {
        "品种类型": row["品种类型"],
        "交易所代码": row["交易所代码"],
        "交易所名称": row["交易所名称"],
        "品种代码": row["品种代码"],
        "品种名称": row["品种名称"],
        "合约行数": row["合约行数"],
        "优先级": row["优先级"],
        "合约月份规则": "",
        "合约数量规则": "",
        "合约数量规则标准化代码": "",
        "到期日规则": "",
        "最后交易日规则": "",
        "是否节假日顺延": "",
        "行权方式": "",
        "交割方式": "",
        "规则来源": "",
        "核验状态": "预填待人工核验",
        "备注": "",
    }


def seed_cffex_index_future(row, output):
    output.update(
        {
            "合约月份规则": "当月、下月及随后两个季月",
            "合约数量规则": "二近二季",
            "合约数量规则标准化代码": "INDEX_FUTURE_2_NEAR_2_QUARTER",
            "到期日规则": "合约到期月份的第三个星期五，遇国家法定假日顺延",
            "最后交易日规则": "合约到期月份的第三个星期五，遇国家法定假日顺延",
            "是否节假日顺延": "是",
            "行权方式": original_or_default(row, "行权方式", "不适用"),
            "交割方式": original_or_default(row, "交割方式", "现金交割"),
            "规则来源": "中国金融期货交易所股指期货合约规则，待人工复核",
            "备注": "股指期货规则预填；请人工核对中金所官网",
        }
    )


def seed_shfe_ine_future(row, output):
    output.update(
        {
            "合约月份规则": "具体合约月份按交易所各品种合约文本执行",
            "合约数量规则": "待人工按品种核验",
            "合约数量规则标准化代码": "FUTURE_RULE_PENDING_MANUAL",
            "到期日规则": "Tushare已提供具体合约最后交易日/最后交割日；规则文本待人工按交易所品种合约核验",
            "最后交易日规则": "待人工按品种核验",
            "是否节假日顺延": "待人工核验",
            "行权方式": original_or_default(row, "行权方式", "不适用"),
            "交割方式": original_or_default(row, "交割方式", "实物交割"),
            "规则来源": "上海期货交易所/上海国际能源交易中心品种合约文本，待人工复核",
            "备注": "商品期货规则预填；需按具体品种合约文本核验",
        }
    )


def seed_dce_future(row, output):
    output.update(
        {
            "合约月份规则": "具体合约月份按大连商品交易所各品种合约文本执行",
            "合约数量规则": "待人工按品种核验",
            "合约数量规则标准化代码": "FUTURE_RULE_PENDING_MANUAL",
            "到期日规则": "Tushare已提供具体合约日期；规则文本待人工按交易所品种合约核验",
            "最后交易日规则": "待人工按品种核验",
            "是否节假日顺延": "待人工核验",
            "行权方式": original_or_default(row, "行权方式", "不适用"),
            "交割方式": original_or_default(row, "交割方式", "实物交割"),
            "规则来源": "大连商品交易所品种合约文本，待人工复核",
            "备注": "商品期货规则预填；需按具体品种合约文本核验",
        }
    )


def seed_czce_future(row, output):
    output.update(
        {
            "合约月份规则": "具体合约月份按郑州商品交易所各品种合约文本执行",
            "合约数量规则": "待人工按品种核验",
            "合约数量规则标准化代码": "FUTURE_RULE_PENDING_MANUAL",
            "到期日规则": "Tushare已提供具体合约日期；规则文本待人工按交易所品种合约核验",
            "最后交易日规则": "待人工按品种核验",
            "是否节假日顺延": "待人工核验",
            "行权方式": original_or_default(row, "行权方式", "不适用"),
            "交割方式": original_or_default(row, "交割方式", "实物交割"),
            "规则来源": "郑州商品交易所品种合约文本，待人工复核",
            "备注": "商品期货规则预填；需按具体品种合约文本核验",
        }
    )


def seed_option(row, output):
    output.update(
        {
            "合约月份规则": "期货期权合约月份规则，待人工按交易所品种核验",
            "合约数量规则": "待人工按交易所品种核验",
            "合约数量规则标准化代码": "FUTURE_OPTION_RULE_PENDING_MANUAL",
            "到期日规则": "Tushare已提供具体到期日/最后交易日；规则文本待人工按交易所品种核验",
            "最后交易日规则": "Tushare已提供具体最后交易日；规则文本待人工按交易所品种核验",
            "是否节假日顺延": "待人工核验",
            "行权方式": original_or_default(row, "行权方式", "美式"),
            "交割方式": original_or_default(
                row, "交割方式", "期货期权行权后转为相应期货持仓，待人工核验"
            ),
            "规则来源": "交易所期权合约规则，待人工复核",
            "备注": "商品期权规则预填；需按具体品种期权合约规则核验",
        }
    )


def seed_default(row, output):
    output.update(
        {
            "合约月份规则": row.get("合约月份规则", ""),
            "合约数量规则": row.get("合约数量规则", ""),
            "合约数量规则标准化代码": row.get("合约数量规则标准化代码", ""),
            "到期日规则": row.get("到期日规则", ""),
            "最后交易日规则": row.get("最后交易日规则", ""),
            "是否节假日顺延": row.get("是否节假日顺延", ""),
            "行权方式": row.get("行权方式", ""),
            "交割方式": row.get("交割方式", ""),
            "规则来源": "待人工补充规则来源",
            "备注": "未匹配预填规则；请人工补充",
        }
    )


def build_seed(input_template):
    rows = []
    for _, row in input_template.iterrows():
        output = base_row(row)
        instrument_type = normalize_code(row["品种类型"])
        exchange_code = normalize_code(row["交易所代码"])
        product_code = normalize_code(row["品种代码"])

        if (
            instrument_type == "FUTURE"
            and exchange_code == "CFFEX"
            and product_code in {"IF", "IH", "IC", "IM"}
        ):
            seed_cffex_index_future(row, output)
        elif instrument_type == "FUTURE" and exchange_code in {"SHFE", "INE"}:
            seed_shfe_ine_future(row, output)
        elif instrument_type == "FUTURE" and exchange_code == "DCE":
            seed_dce_future(row, output)
        elif instrument_type == "FUTURE" and exchange_code == "CZCE":
            seed_czce_future(row, output)
        elif instrument_type == "OPTION":
            seed_option(row, output)
        else:
            seed_default(row, output)

        rows.append(output)

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


def add_check(checks, check_item, passed, message):
    checks.append(
        {
            "check_item": check_item,
            "status": "passed" if passed else "failed",
            "message": message,
        }
    )


def build_checks(input_template, seed):
    checks = []
    add_check(checks, "input_template_exists", INPUT_PATH.exists(), str(INPUT_PATH))
    add_check(checks, "seed_created", OUTPUT_PATH.exists(), str(OUTPUT_PATH))
    add_check(
        checks,
        "seed_rows_equal_input_rows",
        len(seed) == len(input_template),
        f"input rows: {len(input_template)}, seed rows: {len(seed)}",
    )
    missing_columns = [column for column in OUTPUT_COLUMNS if column not in seed.columns]
    add_check(
        checks,
        "required_columns_exist",
        not missing_columns,
        "required columns exist"
        if not missing_columns
        else "missing columns: " + "；".join(missing_columns),
    )
    add_check(
        checks,
        "no_empty_product_code",
        not (seed["品种代码"].apply(clean_text) == "").any(),
        "no empty product code",
    )
    add_check(
        checks,
        "no_empty_exchange_code",
        not (seed["交易所代码"].apply(clean_text) == "").any(),
        "no empty exchange code",
    )
    add_check(
        checks,
        "no_empty_rule_source",
        not (seed["规则来源"].apply(clean_text) == "").any(),
        "no empty rule source",
    )
    add_check(
        checks,
        "all_status_prefill_pending",
        (seed["核验状态"] == "预填待人工核验").all(),
        "all rows are 预填待人工核验",
    )
    cffex_futures = seed[
        (seed["品种类型"].apply(normalize_code) == "FUTURE")
        & (seed["交易所代码"].apply(normalize_code) == "CFFEX")
    ]
    add_check(
        checks,
        "contains_IF_IH_IC_IM",
        {"IF", "IH", "IC", "IM"}.issubset(
            set(cffex_futures["品种代码"].apply(normalize_code))
        ),
        "IF/IH/IC/IM exist",
    )
    add_check(
        checks,
        "contains_core_options",
        (seed["品种类型"].apply(normalize_code) == "OPTION").any(),
        "OPTION rows exist",
    )
    return pd.DataFrame(checks, columns=["check_item", "status", "message"])


def build_summary(input_template, seed, check_report):
    cffex_index_future_rows = seed[
        (seed["品种类型"].apply(normalize_code) == "FUTURE")
        & (seed["交易所代码"].apply(normalize_code) == "CFFEX")
        & (seed["品种代码"].apply(normalize_code).isin(["IF", "IH", "IC", "IM"]))
    ]
    manual_pending_rows = seed[
        seed["合约数量规则标准化代码"].str.contains("PENDING", na=False)
    ]
    return pd.DataFrame(
        [
            {"metric": "input_rows", "value": len(input_template)},
            {"metric": "seed_rows", "value": len(seed)},
            {
                "metric": "future_rows",
                "value": int((seed["品种类型"].apply(normalize_code) == "FUTURE").sum()),
            },
            {
                "metric": "option_rows",
                "value": int((seed["品种类型"].apply(normalize_code) == "OPTION").sum()),
            },
            {
                "metric": "cffex_index_future_rows",
                "value": len(cffex_index_future_rows),
            },
            {"metric": "manual_pending_rows", "value": len(manual_pending_rows)},
            {
                "metric": "check_failed_count",
                "value": int((check_report["status"] == "failed").sum()),
            },
        ],
        columns=["metric", "value"],
    )


def main():
    input_template = read_csv(INPUT_PATH, required=True)
    seed = build_seed(input_template)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    SUMMARY_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    seed.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

    check_report = build_checks(input_template, seed)
    summary = build_summary(input_template, seed, check_report)
    check_report.to_csv(CHECK_OUTPUT, index=False, encoding="utf-8-sig")
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    check_failed_count = int((check_report["status"] == "failed").sum())

    print("core exchange rule seed created successfully")
    print(f"input rows: {len(input_template)}")
    print(f"seed rows: {len(seed)}")
    print(f"check failed count: {check_failed_count}")
    print(f"seed saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
