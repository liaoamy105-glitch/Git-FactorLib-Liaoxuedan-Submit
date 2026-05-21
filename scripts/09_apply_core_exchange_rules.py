from pathlib import Path

import pandas as pd


DETAIL_INPUT = Path("data/processed/detail_sheet.csv")
SEED_INPUT = Path("data/manual/core_exchange_rule_seed.csv")
OUTPUT_PATH = Path("data/processed/detail_sheet_rule_applied.csv")
SUMMARY_OUTPUT = Path("data/processed/apply_core_rules_summary.csv")
CHECK_OUTPUT = Path("data/processed/apply_core_rules_check.csv")

KEY_COLUMNS = ["品种类型", "交易所代码", "品种代码"]
APPLY_FIELDS = [
    "合约月份规则",
    "合约数量规则",
    "合约数量规则标准化代码",
    "到期日规则",
    "最后交易日规则",
    "行权方式",
    "交割方式",
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


def make_key(row):
    return (
        clean_text(row.get("品种类型", "")),
        normalize_code(row.get("交易所代码", "")),
        normalize_code(row.get("品种代码", "")),
    )


def build_seed_lookup(seed):
    lookup = {}
    for _, row in seed.iterrows():
        key = make_key(row)
        if key not in lookup:
            lookup[key] = row
    return lookup


def append_remark(original_remark, seed_remark):
    seed_remark = clean_text(seed_remark)
    if not seed_remark:
        return original_remark
    addition = f"核心规则预填已回填：{seed_remark}"
    original_remark = clean_text(original_remark)
    if not original_remark:
        return addition
    return f"{original_remark}；{addition}"


def apply_rules(detail, seed):
    output = detail.copy()
    original_columns = detail.columns.tolist()
    available_fields = [field for field in APPLY_FIELDS if field in output.columns]
    skipped_fields = [field for field in APPLY_FIELDS if field not in output.columns]
    seed_lookup = build_seed_lookup(seed)
    matched_seed_keys = set()
    updated_indices = set()

    for index, row in output.iterrows():
        key = make_key(row)
        seed_row = seed_lookup.get(key)
        if seed_row is None:
            continue

        matched_seed_keys.add(key)
        row_updated = False
        for field in available_fields:
            seed_value = clean_text(seed_row.get(field, ""))
            if not seed_value:
                continue
            if field == "备注":
                output.at[index, field] = append_remark(output.at[index, field], seed_value)
            else:
                output.at[index, field] = seed_value
            row_updated = True

        if row_updated:
            updated_indices.add(index)

    output = output.reindex(columns=original_columns)
    return output, matched_seed_keys, updated_indices, skipped_fields


def add_check(checks, check_item, passed, message):
    checks.append(
        {
            "check_item": check_item,
            "status": "passed" if passed else "failed",
            "message": message,
        }
    )


def key_filter(data, instrument_type=None, exchange_code=None, product_codes=None):
    filtered = data.copy()
    if instrument_type:
        filtered = filtered[filtered["品种类型"].apply(clean_text) == instrument_type]
    if exchange_code:
        filtered = filtered[filtered["交易所代码"].apply(normalize_code) == exchange_code]
    if product_codes:
        product_code_set = {normalize_code(code) for code in product_codes}
        filtered = filtered[filtered["品种代码"].apply(normalize_code).isin(product_code_set)]
    return filtered


def all_seed_units_matched(seed, matched_seed_keys):
    seed_keys = {make_key(row) for _, row in seed.iterrows()}
    unmatched = sorted(seed_keys - matched_seed_keys)
    return unmatched


def build_checks(detail, output, seed, matched_seed_keys, skipped_fields):
    checks = []
    add_check(checks, "input_detail_exists", DETAIL_INPUT.exists(), str(DETAIL_INPUT))
    add_check(checks, "input_seed_exists", SEED_INPUT.exists(), str(SEED_INPUT))
    add_check(checks, "output_created", OUTPUT_PATH.exists(), str(OUTPUT_PATH))
    add_check(
        checks,
        "output_rows_equal_input_rows",
        len(output) == len(detail),
        f"input rows: {len(detail)}, output rows: {len(output)}",
    )
    add_check(
        checks,
        "output_columns_equal_input_columns",
        output.columns.tolist() == detail.columns.tolist(),
        "output columns equal input columns"
        if output.columns.tolist() == detail.columns.tolist()
        else "output columns differ from input columns",
    )

    unmatched_seed = all_seed_units_matched(seed, matched_seed_keys)
    add_check(
        checks,
        "all_seed_units_matched_detail",
        not unmatched_seed,
        "all seed units matched detail"
        if not unmatched_seed
        else "unmatched seed units: " + "；".join("/".join(item) for item in unmatched_seed),
    )

    index_futures = key_filter(output, "FUTURE", "CFFEX", ["IF", "IH", "IC", "IM"])
    index_future_passed = (
        not index_futures.empty
        and index_futures["合约月份规则"].str.contains("当月、下月及随后两个季月", na=False).all()
        and (
            index_futures["合约数量规则标准化代码"]
            == "INDEX_FUTURE_2_NEAR_2_QUARTER"
        ).all()
    )
    add_check(
        checks,
        "IF_IH_IC_IM_rules_applied",
        index_future_passed,
        f"checked rows: {len(index_futures)}",
    )

    seed_future_codes = seed[
        seed["品种类型"].apply(clean_text) == "FUTURE"
    ][["交易所代码", "品种代码"]].drop_duplicates()
    core_future = output[output["品种类型"].apply(clean_text) == "FUTURE"].copy()
    if not seed_future_codes.empty:
        seed_future_keys = set(
            zip(
                seed_future_codes["交易所代码"].apply(normalize_code),
                seed_future_codes["品种代码"].apply(normalize_code),
            )
        )
        core_future = core_future[
            list(
                zip(
                    core_future["交易所代码"].apply(normalize_code),
                    core_future["品种代码"].apply(normalize_code),
                )
            )
        ]
    core_future_rules_applied = (
        not core_future.empty
        and (core_future["合约月份规则"].apply(clean_text) != "").all()
    )
    add_check(
        checks,
        "core_future_rules_applied",
        core_future_rules_applied,
        f"core FUTURE rows checked: {len(core_future)}; skipped fields: {skipped_fields}",
    )

    seed_option_codes = seed[
        seed["品种类型"].apply(clean_text) == "OPTION"
    ][["交易所代码", "品种代码"]].drop_duplicates()
    core_option = output[output["品种类型"].apply(clean_text) == "OPTION"].copy()
    if not seed_option_codes.empty:
        seed_option_keys = set(
            zip(
                seed_option_codes["交易所代码"].apply(normalize_code),
                seed_option_codes["品种代码"].apply(normalize_code),
            )
        )
        option_keys = list(
            zip(
                core_option["交易所代码"].apply(normalize_code),
                core_option["品种代码"].apply(normalize_code),
            )
        )
        core_option = core_option[[key in seed_option_keys for key in option_keys]]
    core_option_rules_applied = (
        not core_option.empty
        and (
            core_option["合约数量规则标准化代码"]
            == "FUTURE_OPTION_RULE_PENDING_MANUAL"
        ).all()
    )
    add_check(
        checks,
        "core_option_rules_applied",
        core_option_rules_applied,
        f"core OPTION rows checked: {len(core_option)}; skipped fields: {skipped_fields}",
    )

    add_check(
        checks,
        "no_empty_product_code_after_apply",
        not (output["品种代码"].apply(clean_text) == "").any(),
        "no empty product code",
    )
    add_check(
        checks,
        "no_empty_exchange_code_after_apply",
        not (output["交易所代码"].apply(clean_text) == "").any(),
        "no empty exchange code",
    )
    return pd.DataFrame(checks, columns=["check_item", "status", "message"])


def build_summary(detail, output, seed, matched_seed_keys, updated_indices, check_report):
    seed_keys = {make_key(row) for _, row in seed.iterrows()}
    unmatched_seed_units = len(seed_keys - matched_seed_keys)
    updated_rows = output.loc[list(updated_indices)] if updated_indices else output.iloc[0:0]
    return pd.DataFrame(
        [
            {"metric": "input_detail_rows", "value": len(detail)},
            {"metric": "output_detail_rows", "value": len(output)},
            {"metric": "seed_rows", "value": len(seed)},
            {"metric": "matched_seed_units", "value": len(matched_seed_keys)},
            {"metric": "unmatched_seed_units", "value": unmatched_seed_units},
            {"metric": "updated_detail_rows", "value": len(updated_indices)},
            {
                "metric": "updated_future_rows",
                "value": int((updated_rows["品种类型"].apply(clean_text) == "FUTURE").sum()),
            },
            {
                "metric": "updated_option_rows",
                "value": int((updated_rows["品种类型"].apply(clean_text) == "OPTION").sum()),
            },
            {
                "metric": "check_failed_count",
                "value": int((check_report["status"] == "failed").sum()),
            },
        ],
        columns=["metric", "value"],
    )


def main():
    detail = read_csv(DETAIL_INPUT, required=True)
    seed = read_csv(SEED_INPUT, required=True)

    output, matched_seed_keys, updated_indices, skipped_fields = apply_rules(detail, seed)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    output.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

    check_report = build_checks(detail, output, seed, matched_seed_keys, skipped_fields)
    summary = build_summary(
        detail, output, seed, matched_seed_keys, updated_indices, check_report
    )
    check_report.to_csv(CHECK_OUTPUT, index=False, encoding="utf-8-sig")
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    check_failed_count = int((check_report["status"] == "failed").sum())

    print("core exchange rules applied successfully")
    print(f"input detail rows: {len(detail)}")
    print(f"output detail rows: {len(output)}")
    print(f"seed rows: {len(seed)}")
    print(f"updated detail rows: {len(updated_indices)}")
    print(f"check failed count: {check_failed_count}")
    print(f"output saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
