from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
STAGE_DIR = ROOT / "data" / "stage2_options"
FINAL_DIR = STAGE_DIR / "final"
PROCESSED_DIR = STAGE_DIR / "processed"
MANUAL_DIR = STAGE_DIR / "manual"
OUTPUT_DIR = ROOT / "output"
ENCODING = "utf-8-sig"

OPTIONS_IN = FINAL_DIR / "options_master_tushare_primary_cleaned.csv"
RULE_IN = FINAL_DIR / "option_contract_rule_detail_tushare_primary_cleaned.csv"
ASSET_IN = FINAL_DIR / "asset_master_with_options_tushare_primary_cleaned.csv"
REVIEW_IN = PROCESSED_DIR / "options_underlying_review_list.csv"
CONTRACT_RAW_IN = PROCESSED_DIR / "options_contract_raw_tushare_primary.csv"
QUALITY_IN = PROCESSED_DIR / "options_quality_report_tushare_primary.csv"
FIX_LOG_IN = PROCESSED_DIR / "options_underlying_fix_log.csv"

PATCH_FILE = MANUAL_DIR / "manual_option_underlying_group_patch.csv"
OPTIONS_OUT = FINAL_DIR / "options_master_tushare_primary_final.csv"
RULE_OUT = FINAL_DIR / "option_contract_rule_detail_tushare_primary_final.csv"
ASSET_OUT = FINAL_DIR / "asset_master_with_options_tushare_primary_final.csv"
PATCH_LOG_OUT = PROCESSED_DIR / "options_underlying_group_patch_log.csv"
REMAINING_REVIEW_OUT = PROCESSED_DIR / "options_underlying_group_remaining_review.csv"
SUMMARY_OUT = OUTPUT_DIR / "options_underlying_group_patch_summary.csv"
EXCEL_OUT = OUTPUT_DIR / "mapping_info_stage2_options_final.xlsx"

GROUP_MAPPING = {
    "AP": "APPLE",
    "CJ": "JUJUBE",
    "CY": "COTTON_YARN",
    "ER": "EARLY_INDICA_RICE",
    "FG": "GLASS",
    "JR": "JAPONICA_RICE",
    "LR": "LATE_INDICA_RICE",
    "MA": "METHANOL",
    "ME": "METHANOL",
    "OI": "RAPESEED_OIL",
    "PF": "POLYESTER_STAPLE_FIBER",
    "PK": "PEANUT",
    "PL": "PROPYLENE",
    "PM": "COMMON_WHEAT",
    "PR": "BOTTLE_GRADE_PET",
    "PX": "PARAXYLENE",
    "RI": "EARLY_INDICA_RICE",
    "RM": "RAPESEED_MEAL",
    "RO": "RAPESEED_OIL",
    "RS": "RAPESEED",
    "SA": "SODA_ASH",
    "SF": "FERROSILICON",
    "SH": "CAUSTIC_SODA",
    "SM": "SILICOMANGANESE",
    "SR": "SUGAR",
    "TA": "PTA",
    "TC": "THERMAL_COAL",
    "UR": "UREA",
    "WH": "STRONG_WHEAT",
    "WS": "STRONG_WHEAT",
    "WT": "HARD_WHEAT",
    "ZC": "THERMAL_COAL",
    "BB": "PLYWOOD",
    "BZ": "BENZENE",
    "FB": "FIBERBOARD",
    "LG": "LOG",
    "PG": "LPG",
    "RR": "JAPONICA_RICE",
    "AD": "CAST_ALUMINUM_ALLOY",
    "AO": "ALUMINA",
    "BR": "BUTADIENE_RUBBER",
    "OP": "OFFSET_PAPER",
    "SP": "PULP",
    "SS": "STAINLESS_STEEL",
    "WR": "WIRE_ROD",
    "EC": "CONTAINER_FREIGHT_INDEX_EUROPE",
    "SCTAS": "CRUDE_OIL",
    "PD": "PALLADIUM",
    "PS": "POLYSILICON",
    "PT": "PLATINUM",
}
MEDIUM_CONFIDENCE_SYMBOLS = {"ER", "LR", "ME", "RO", "TC", "WS", "WT", "SCTAS"}
PATCH_COLUMNS = [
    "asset_id",
    "option_symbol",
    "option_name_cn",
    "exchange_code",
    "underlying_asset_id",
    "underlying_symbol",
    "old_underlying_group",
    "new_underlying_group",
    "underlying_name_cn",
    "patch_method",
    "mapping_confidence",
    "review_status",
    "notes",
]


def ensure_dirs() -> None:
    for path in [FINAL_DIR, PROCESSED_DIR, MANUAL_DIR, OUTPUT_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def read_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"缺少输入文件: {path}")
    return pd.read_csv(path, dtype=str).fillna("")


def read_optional_csv(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_csv(path, dtype=str).fillna("")
    return pd.DataFrame({"message": [f"{path.name} not found"]})


def save_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding=ENCODING)


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def is_missing(value: Any) -> bool:
    text = clean_text(value)
    return text == "" or text.upper() in {"CHECK", "TODO", "NA", "NAN", "NONE"}


def append_note(old_note: Any, addition: str) -> str:
    old = clean_text(old_note)
    if not old:
        return addition
    if addition in old:
        return old
    return f"{old} | {addition}"


def standard_symbol(*values: Any) -> str:
    for value in values:
        text = clean_text(value).upper()
        if not text or text == "CHECK":
            continue
        match = re.match(r"([A-Z]+)", text.split(".")[0])
        if match:
            return match.group(1)
    return ""


def confidence_for(symbol: str) -> str:
    return "MEDIUM" if symbol in MEDIUM_CONFIDENCE_SYMBOLS else "HIGH"


def notes_for(symbol: str) -> str:
    if symbol in MEDIUM_CONFIDENCE_SYMBOLS:
        return "historical or special symbol, need review"
    return "conservative option symbol mapping"


def build_or_update_patch(review: pd.DataFrame) -> pd.DataFrame:
    candidates = review[review["underlying_group"].apply(is_missing)].copy() if not review.empty else pd.DataFrame()
    new_rows = []
    for _, row in candidates.iterrows():
        symbol = standard_symbol(row.get("option_symbol"), row.get("underlying_symbol"))
        mapped_group = GROUP_MAPPING.get(symbol, "CHECK")
        confidence = confidence_for(symbol) if mapped_group != "CHECK" else "CHECK"
        new_rows.append(
            {
                "asset_id": row.get("asset_id", ""),
                "option_symbol": row.get("option_symbol", ""),
                "option_name_cn": row.get("option_name_cn", ""),
                "exchange_code": row.get("exchange_code", ""),
                "underlying_asset_id": row.get("underlying_asset_id", ""),
                "underlying_symbol": row.get("underlying_symbol", ""),
                "old_underlying_group": row.get("underlying_group", ""),
                "new_underlying_group": mapped_group,
                "underlying_name_cn": row.get("option_name_cn", "").replace("期权", ""),
                "patch_method": "OPTION_SYMBOL_CONSERVATIVE_RULE" if mapped_group != "CHECK" else "NOT_PATCHED_NEED_REVIEW",
                "mapping_confidence": confidence,
                "review_status": "TODO",
                "notes": notes_for(symbol) if mapped_group != "CHECK" else "no conservative mapping found",
            }
        )
    generated = pd.DataFrame(new_rows, columns=PATCH_COLUMNS)
    if PATCH_FILE.exists():
        existing = pd.read_csv(PATCH_FILE, dtype=str).fillna("")
        for col in PATCH_COLUMNS:
            if col not in existing.columns:
                existing[col] = ""
        existing = existing[PATCH_COLUMNS].copy()
        existing_ids = set(existing["asset_id"])
        additions = generated[~generated["asset_id"].isin(existing_ids)].copy()
        patch = pd.concat([existing, additions], ignore_index=True)
    else:
        patch = generated
    save_csv(patch, PATCH_FILE)
    return patch


def apply_patch(options: pd.DataFrame, patch: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    cleaned = options.copy()
    patch_by_id = patch.set_index("asset_id").to_dict("index") if not patch.empty else {}
    log_rows = []
    for idx, row in cleaned.iterrows():
        old_group = clean_text(row.get("underlying_group"))
        if not is_missing(old_group):
            continue
        asset_id = clean_text(row.get("asset_id"))
        patch_row = patch_by_id.get(asset_id, {})
        new_group = clean_text(patch_row.get("new_underlying_group", "CHECK"))
        old_status = clean_text(row.get("source_status"))
        applied = new_group != "" and new_group != "CHECK"
        if applied:
            cleaned.at[idx, "underlying_group"] = new_group
            cleaned.at[idx, "source_status"] = "TUSHARE_PRIMARY_UNDERLYING_GROUP_PATCHED_NEED_REVIEW"
            cleaned.at[idx, "notes"] = append_note(
                row.get("notes"),
                "underlying_group patched by conservative option symbol mapping, need review",
            )
        else:
            cleaned.at[idx, "underlying_group"] = "CHECK"
            cleaned.at[idx, "source_status"] = "UNDERLYING_GROUP_NEED_REVIEW"
        log_rows.append(
            {
                "asset_id": asset_id,
                "option_symbol": row.get("option_symbol", ""),
                "option_name_cn": row.get("option_name_cn", ""),
                "exchange_code": row.get("exchange_code", ""),
                "old_underlying_group": old_group,
                "new_underlying_group": cleaned.at[idx, "underlying_group"],
                "patch_applied": "Y" if applied else "N",
                "patch_method": patch_row.get("patch_method", "NOT_PATCHED_NEED_REVIEW"),
                "mapping_confidence": patch_row.get("mapping_confidence", "CHECK"),
                "old_source_status": old_status,
                "new_source_status": cleaned.at[idx, "source_status"],
                "notes": patch_row.get("notes", ""),
            }
        )
    log_df = pd.DataFrame(log_rows)
    save_csv(cleaned, OPTIONS_OUT)
    save_csv(log_df, PATCH_LOG_OUT)
    return cleaned, log_df


def sync_rule_detail(rule: pd.DataFrame, final_options: pd.DataFrame) -> pd.DataFrame:
    result = rule.copy()
    for field in ["underlying_asset_id", "underlying_symbol", "underlying_group"]:
        if field not in result.columns:
            result[field] = ""
    mapping = final_options.set_index("asset_id")[["underlying_asset_id", "underlying_symbol", "underlying_group"]].to_dict("index")
    for idx, row in result.iterrows():
        values = mapping.get(clean_text(row.get("asset_id")), {})
        for field, value in values.items():
            result.at[idx, field] = value
    save_csv(result, RULE_OUT)
    return result


def sync_asset_master(asset: pd.DataFrame, final_options: pd.DataFrame) -> pd.DataFrame:
    result = asset.copy()
    mapping = final_options.set_index("asset_id")[["underlying_group", "source_status", "notes"]].to_dict("index")
    for idx, row in result.iterrows():
        if clean_text(row.get("asset_type")) != "OPTION":
            continue
        values = mapping.get(clean_text(row.get("asset_id")), {})
        for field, value in values.items():
            if field not in result.columns:
                result[field] = ""
            result.at[idx, field] = value
    save_csv(result, ASSET_OUT)
    return result


def build_remaining_review(final_options: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "asset_id",
        "option_symbol",
        "option_name_cn",
        "exchange_code",
        "underlying_asset_id",
        "underlying_symbol",
        "underlying_group",
        "issue_reason",
        "suggested_next_step",
        "notes",
    ]
    rows = []
    for _, row in final_options.iterrows():
        if not is_missing(row.get("underlying_group")):
            continue
        rows.append(
            {
                "asset_id": row.get("asset_id", ""),
                "option_symbol": row.get("option_symbol", ""),
                "option_name_cn": row.get("option_name_cn", ""),
                "exchange_code": row.get("exchange_code", ""),
                "underlying_asset_id": row.get("underlying_asset_id", ""),
                "underlying_symbol": row.get("underlying_symbol", ""),
                "underlying_group": row.get("underlying_group", ""),
                "issue_reason": "underlying_group still missing after conservative patch",
                "suggested_next_step": "需要人工确认该期权品种对应的统一标的，或在 ETF/现货阶段补充标的资产后重新映射。",
                "notes": row.get("notes", ""),
            }
        )
    review = pd.DataFrame(rows, columns=columns)
    save_csv(review, REMAINING_REVIEW_OUT)
    return review


def missing_group_count(df: pd.DataFrame) -> int:
    return int(df["underlying_group"].apply(is_missing).sum()) if "underlying_group" in df.columns else len(df)


def build_summary(before: pd.DataFrame, final_options: pd.DataFrame, log_df: pd.DataFrame, remaining: pd.DataFrame) -> pd.DataFrame:
    rows = [
        {
            "section": "overall",
            "subtype": "ALL",
            "exchange_code": "ALL",
            "total_options_count": len(final_options),
            "underlying_group_missing_before": missing_group_count(before),
            "underlying_group_missing_after": missing_group_count(final_options),
            "patch_candidate_count": len(log_df),
            "patch_applied_count": int((log_df["patch_applied"] == "Y").sum()) if not log_df.empty else 0,
            "remaining_review_count": len(remaining),
            "high_confidence_patch_count": int(((log_df["patch_applied"] == "Y") & (log_df["mapping_confidence"] == "HIGH")).sum()) if not log_df.empty else 0,
            "medium_confidence_patch_count": int(((log_df["patch_applied"] == "Y") & (log_df["mapping_confidence"] == "MEDIUM")).sum()) if not log_df.empty else 0,
            "total_count": len(final_options),
            "underlying_group_missing_after_by_group": missing_group_count(final_options),
        }
    ]
    remaining_ids = set(remaining["asset_id"]) if not remaining.empty else set()
    applied_ids = set(log_df.loc[log_df["patch_applied"] == "Y", "asset_id"]) if not log_df.empty else set()
    for (subtype, exchange), group in final_options.groupby(["subtype", "exchange_code"], dropna=False):
        rows.append(
            {
                "section": "by_subtype_exchange",
                "subtype": subtype,
                "exchange_code": exchange,
                "total_options_count": "",
                "underlying_group_missing_before": "",
                "underlying_group_missing_after": "",
                "patch_candidate_count": "",
                "patch_applied_count": int(group["asset_id"].isin(applied_ids).sum()),
                "remaining_review_count": int(group["asset_id"].isin(remaining_ids).sum()),
                "high_confidence_patch_count": "",
                "medium_confidence_patch_count": "",
                "total_count": len(group),
                "underlying_group_missing_after_by_group": missing_group_count(group),
            }
        )
    summary = pd.DataFrame(rows)
    save_csv(summary, SUMMARY_OUT)
    return summary


def readme_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ("项目", "Step 3 期权基础表 final draft"),
            ("主数据源", "Tushare opt_basic 为主数据源。"),
            ("覆盖范围", "已覆盖股指期权、ETF期权、商品期权。"),
            ("标的映射", "已根据 Stage 1 futures_master 匹配 commodity option 的 underlying_asset_id。"),
            ("标的组补丁", "已根据保守字典补齐部分 Stage 1 中 CHECK 的商品 underlying_group。"),
            ("ETF期权", "ETF期权的 underlying_asset_id 将在 ETF现货阶段进一步补齐。"),
            ("规则复核", "正式期权交易规则仍需交易所规则复核。"),
        ],
        columns=["item", "description"],
    )


def write_excel(
    asset: pd.DataFrame,
    options: pd.DataFrame,
    rule: pd.DataFrame,
    log_df: pd.DataFrame,
    remaining: pd.DataFrame,
    summary: pd.DataFrame,
) -> None:
    sheets = {
        "README": readme_df(),
        "asset_master_with_options": asset,
        "options_master": options,
        "option_contract_rule_detail": rule,
        "options_underlying_group_patch_log": log_df,
        "options_underlying_group_remaining_review": remaining,
        "options_quality_summary": summary,
    }
    if CONTRACT_RAW_IN.exists():
        sheets["options_contract_raw_tushare"] = read_optional_csv(CONTRACT_RAW_IN)
    if QUALITY_IN.exists():
        sheets["quality_report_tushare_primary"] = read_optional_csv(QUALITY_IN)
    if FIX_LOG_IN.exists():
        sheets["options_underlying_fix_log"] = read_optional_csv(FIX_LOG_IN)
    with pd.ExcelWriter(EXCEL_OUT, engine="openpyxl") as writer:
        for name, df in sheets.items():
            sheet = name[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.book[sheet]
            ws.freeze_panes = "A2"
            if ws.max_row > 1 and ws.max_column > 1:
                ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col_cells in ws.columns:
                width = 10
                letter = col_cells[0].column_letter
                for cell in col_cells[:200]:
                    width = max(width, min(len("" if cell.value is None else str(cell.value)) + 2, 40))
                ws.column_dimensions[letter].width = width


def main() -> None:
    ensure_dirs()
    options = read_csv(OPTIONS_IN)
    rule = read_csv(RULE_IN)
    asset = read_csv(ASSET_IN)
    review = read_csv(REVIEW_IN)
    before_missing = missing_group_count(options)

    patch = build_or_update_patch(review)
    final_options, patch_log = apply_patch(options, patch)
    final_rule = sync_rule_detail(rule, final_options)
    final_asset = sync_asset_master(asset, final_options)
    remaining = build_remaining_review(final_options)
    summary = build_summary(options, final_options, patch_log, remaining)
    write_excel(final_asset, final_options, final_rule, patch_log, remaining, summary)

    high_count = int(((patch_log["patch_applied"] == "Y") & (patch_log["mapping_confidence"] == "HIGH")).sum()) if not patch_log.empty else 0
    medium_count = int(((patch_log["patch_applied"] == "Y") & (patch_log["mapping_confidence"] == "MEDIUM")).sum()) if not patch_log.empty else 0
    print(f"options_master 行数: {len(options)}")
    print(f"underlying_group 缺失修复前数量: {before_missing}")
    print(f"underlying_group 缺失修复后数量: {missing_group_count(final_options)}")
    print(f"patch_applied_count: {int((patch_log['patch_applied'] == 'Y').sum()) if not patch_log.empty else 0}")
    print(f"remaining_review_count: {len(remaining)}")
    print(f"high_confidence_patch_count: {high_count}")
    print(f"medium_confidence_patch_count: {medium_count}")
    print(f"Excel 输出路径: {EXCEL_OUT}")


if __name__ == "__main__":
    main()
