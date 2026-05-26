from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
STAGE4_DIR = ROOT / "data" / "stage4_etf_index"
FINAL_DIR = STAGE4_DIR / "final"
PROCESSED_DIR = STAGE4_DIR / "processed"
OUTPUT_DIR = ROOT / "output"
ENCODING = "utf-8-sig"

ETF_MASTER_IN = FINAL_DIR / "etf_master.csv"
ASSET_MASTER_IN = FINAL_DIR / "asset_master_with_etf_index.csv"
OPTIONS_IN = FINAL_DIR / "options_master_with_etf_underlying.csv"
MAPPING_IN = FINAL_DIR / "etf_index_mapping.csv"
INDEX_MASTER_IN = FINAL_DIR / "index_master.csv"
RAW_QUALITY_IN = PROCESSED_DIR / "etf_index_quality_report.csv"

ETF_MASTER_OUT = FINAL_DIR / "etf_master_cleaned.csv"
ASSET_MASTER_OUT = FINAL_DIR / "asset_master_with_etf_index_cleaned.csv"
REPORT_OUT = PROCESSED_DIR / "etf_quality_classification_report.csv"
REVIEW_OUT = PROCESSED_DIR / "etf_underlying_group_review_list.csv"
EXCEL_OUT = OUTPUT_DIR / "mapping_info_stage4_etf_index_cleaned.xlsx"

CORE_ETF_CODES = {
    "510050", "510300", "510500", "510880", "588000", "588080", "588050",
    "159919", "159922", "159915", "159901", "159949", "159995", "159601",
    "159605", "159845", "159629",
}


def ensure_dirs() -> None:
    FINAL_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


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


def contains_any(text: str, keywords: list[str]) -> bool:
    return any(keyword in text for keyword in keywords)


def classify_category(row: pd.Series) -> tuple[str, str, str]:
    name = clean_text(row.get("name_cn"))
    fund_type = clean_text(row.get("fund_type"))
    invest_type = clean_text(row.get("invest_type"))
    tracking = clean_text(row.get("tracking_index_name"))
    group = clean_text(row.get("underlying_group"))
    text = f"{name} {fund_type} {invest_type} {tracking} {group}"

    if contains_any(text, ["沪深300", "上证50", "中证500", "中证1000", "创业板", "科创50", "深证100", "红利", "A500", "中证A500"]):
        return "BROAD_BASED_ETF", "IMPORTANT", ""
    if contains_any(text, ["黄金", "白银", "有色", "豆粕", "能源", "原油", "商品"]):
        return "COMMODITY_ETF", "IMPORTANT", "COMMODITY_ETF_NEED_UNDERLYING_GROUP_REVIEW"
    if contains_any(text, ["债", "国债", "政金债", "信用债", "可转债", "货币", "现金", "添利"]):
        return "BOND_OR_MONEY_ETF", "LOW_FOR_CURRENT_ARBITRAGE_MAPPING", "NOT_CORE_FOR_CURRENT_STAGE"
    if contains_any(text, ["纳指", "标普", "德国", "法国", "日经", "恒生", "港股", "中概", "海外", "QDII", "印度", "东南亚"]):
        return "CROSS_BORDER_ETF", "LATER_STAGE", "FOREIGN_MAPPING_STAGE"
    if contains_any(
        text,
        [
            "证券", "银行", "保险", "军工", "芯片", "半导体", "医药", "新能源", "电池", "光伏", "消费",
            "酒", "食品", "传媒", "游戏", "人工智能", "机器人", "数据", "云计算", "软件", "汽车",
            "煤炭", "钢铁", "房地产", "基建", "农业", "畜牧", "养殖", "旅游", "物流",
        ],
    ):
        return "THEMATIC_OR_SECTOR_ETF", "LATER_STAGE", "THEMATIC_ETF_LATER_STAGE"
    return "OTHER_ETF", "LATER_STAGE", "NOT_CORE_FOR_CURRENT_STAGE"


def etf_option_underlying_ids(options: pd.DataFrame) -> set[str]:
    if options.empty or "subtype" not in options.columns:
        return set()
    etf_options = options[options["subtype"] == "ETF_OPTION"]
    return {clean_text(v) for v in etf_options.get("underlying_asset_id", pd.Series(dtype=str)).tolist() if clean_text(v)}


def classify_etfs(etf_master: pd.DataFrame, options: pd.DataFrame) -> pd.DataFrame:
    cleaned = etf_master.copy()
    option_underlying_ids = etf_option_underlying_ids(options)
    new_cols = [
        "etf_importance_level",
        "etf_quality_status",
        "etf_category_refined",
        "underlying_group_review_required",
        "underlying_group_review_reason",
        "is_core_arbitrage_etf",
        "is_etf_option_underlying",
        "classification_method",
    ]
    for col in new_cols:
        cleaned[col] = ""

    for idx, row in cleaned.iterrows():
        code = clean_text(row.get("etf_code"))
        asset_id = clean_text(row.get("asset_id"))
        group_missing = is_missing(row.get("underlying_group"))
        category, importance, default_reason = classify_category(row)
        methods = []
        is_core = code in CORE_ETF_CODES
        is_option_underlying = asset_id in option_underlying_ids

        if is_core:
            importance = "CORE"
            methods.append("CORE_ETF_CODE_RULE")
        if is_option_underlying:
            importance = "CORE"
            is_core = True
            methods.append("ETF_OPTION_UNDERLYING_RULE")
        if not methods:
            methods.append("NAME_KEYWORD_RULE")

        review_required = "N"
        review_reason = default_reason
        if group_missing and importance in {"CORE", "IMPORTANT"}:
            review_required = "Y"
            if importance == "CORE":
                status = "CORE_ETF_UNDERLYING_GROUP_MISSING"
                review_reason = "CORE_ETF_UNDERLYING_GROUP_MISSING"
            else:
                status = "IMPORTANT_UNDERLYING_GROUP_NEED_REVIEW"
                review_reason = review_reason or "IMPORTANT_UNDERLYING_GROUP_NEED_REVIEW"
        elif group_missing and importance in {"LOW_FOR_CURRENT_ARBITRAGE_MAPPING", "LATER_STAGE"}:
            status = "NON_CORE_UNDERLYING_GROUP_NOT_REQUIRED_NOW"
            review_required = "N"
            review_reason = review_reason or "NOT_CORE_FOR_CURRENT_STAGE"
        else:
            status = "CORE_MAPPED" if importance == "CORE" else "MAPPED"
            review_required = "N"
            review_reason = ""

        cleaned.at[idx, "etf_category_refined"] = category
        cleaned.at[idx, "etf_importance_level"] = importance
        cleaned.at[idx, "etf_quality_status"] = status
        cleaned.at[idx, "underlying_group_review_required"] = review_required
        cleaned.at[idx, "underlying_group_review_reason"] = review_reason
        cleaned.at[idx, "is_core_arbitrage_etf"] = "Y" if is_core else "N"
        cleaned.at[idx, "is_etf_option_underlying"] = "Y" if is_option_underlying else "N"
        cleaned.at[idx, "classification_method"] = ";".join(methods)
    save_csv(cleaned, ETF_MASTER_OUT)
    return cleaned


def sync_asset_master(asset_master: pd.DataFrame, etf_cleaned: pd.DataFrame) -> pd.DataFrame:
    result = asset_master.copy()
    mapping = etf_cleaned.set_index("asset_id").to_dict("index")
    for idx, row in result.iterrows():
        if clean_text(row.get("asset_type")) != "ETF":
            continue
        asset_id = clean_text(row.get("asset_id"))
        etf = mapping.get(asset_id)
        if not etf:
            continue
        result.at[idx, "underlying_group"] = etf.get("underlying_group", "")
        result.at[idx, "source_status"] = etf.get("source_status", "")
        notes = clean_text(etf.get("notes", ""))
        if is_missing(etf.get("underlying_group")) and etf.get("underlying_group_review_required") == "N":
            notes = append_note(notes, "non-core ETF underlying_group deferred to later mapping stage")
        result.at[idx, "notes"] = notes
    save_csv(result, ASSET_MASTER_OUT)
    return result


def build_review_list(etf_cleaned: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in etf_cleaned[etf_cleaned["underlying_group_review_required"] == "Y"].iterrows():
        reason = clean_text(row.get("underlying_group_review_reason"))
        if "CORE" in reason:
            step = "需要补充核心ETF对应的指数或统一标的。"
        elif "COMMODITY" in reason:
            step = "需要确认商品ETF对应的商品标的或跟踪指数。"
        else:
            step = "需要确认ETF对应的指数或统一标的。"
        rows.append(
            {
                "asset_id": row.get("asset_id", ""),
                "etf_code": row.get("etf_code", ""),
                "name_cn": row.get("name_cn", ""),
                "exchange_code": row.get("exchange_code", ""),
                "etf_category_refined": row.get("etf_category_refined", ""),
                "etf_importance_level": row.get("etf_importance_level", ""),
                "underlying_group": row.get("underlying_group", ""),
                "underlying_group_review_reason": reason,
                "tracking_index_code": row.get("tracking_index_code", ""),
                "tracking_index_name": row.get("tracking_index_name", ""),
                "source_status": row.get("source_status", ""),
                "suggested_next_step": step,
                "notes": row.get("notes", ""),
            }
        )
    review = pd.DataFrame(
        rows,
        columns=[
            "asset_id", "etf_code", "name_cn", "exchange_code", "etf_category_refined", "etf_importance_level",
            "underlying_group", "underlying_group_review_reason", "tracking_index_code", "tracking_index_name",
            "source_status", "suggested_next_step", "notes",
        ],
    )
    save_csv(review, REVIEW_OUT)
    return review


def build_report(etf_cleaned: pd.DataFrame, review: pd.DataFrame, raw_missing: int) -> pd.DataFrame:
    rows = [
        {
            "section": "overall",
            "category": "ALL",
            "count": len(etf_cleaned),
            "underlying_group_missing_count": raw_missing,
            "review_required_count": len(review),
            "total_etf_count": len(etf_cleaned),
            "core_etf_count": int((etf_cleaned["etf_importance_level"] == "CORE").sum()),
            "important_etf_count": int((etf_cleaned["etf_importance_level"] == "IMPORTANT").sum()),
            "later_stage_etf_count": int((etf_cleaned["etf_importance_level"] == "LATER_STAGE").sum()),
            "bond_or_money_etf_count": int((etf_cleaned["etf_category_refined"] == "BOND_OR_MONEY_ETF").sum()),
            "cross_border_etf_count": int((etf_cleaned["etf_category_refined"] == "CROSS_BORDER_ETF").sum()),
            "thematic_or_sector_etf_count": int((etf_cleaned["etf_category_refined"] == "THEMATIC_OR_SECTOR_ETF").sum()),
            "commodity_etf_count": int((etf_cleaned["etf_category_refined"] == "COMMODITY_ETF").sum()),
            "raw_underlying_group_missing_count": raw_missing,
            "review_required_underlying_group_missing_count": len(review),
            "non_core_missing_deferred_count": int(
                (etf_cleaned["underlying_group"].apply(is_missing) & (etf_cleaned["underlying_group_review_required"] == "N")).sum()
            ),
            "etf_option_underlying_count": int((etf_cleaned["is_etf_option_underlying"] == "Y").sum()),
            "core_mapped_count": int((etf_cleaned["etf_quality_status"] == "CORE_MAPPED").sum()),
        }
    ]
    for category, group in etf_cleaned.groupby("etf_category_refined", dropna=False):
        rows.append(
            {
                "section": "by_category",
                "category": category,
                "count": len(group),
                "underlying_group_missing_count": int(group["underlying_group"].apply(is_missing).sum()),
                "review_required_count": int((group["underlying_group_review_required"] == "Y").sum()),
            }
        )
    for level, group in etf_cleaned.groupby("etf_importance_level", dropna=False):
        rows.append(
            {
                "section": "by_importance",
                "category": level,
                "count": len(group),
                "underlying_group_missing_count": int(group["underlying_group"].apply(is_missing).sum()),
                "review_required_count": int((group["underlying_group_review_required"] == "Y").sum()),
            }
        )
    report = pd.DataFrame(rows)
    save_csv(report, REPORT_OUT)
    return report


def readme_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ("项目", "Step 4.1 ETF质量分层 cleaned 版本"),
            ("说明", "ETF总量较大，原始 underlying_group 缺失不全部视为错误。"),
            ("核心优先", "核心ETF和ETF期权标的ETF已优先映射。"),
            ("后续阶段", "行业ETF、债券ETF、货币ETF、跨境ETF等暂归入后续阶段。"),
            ("复核清单", "review list 只保留当前阶段真正需要处理的核心/重要 ETF。"),
        ],
        columns=["item", "description"],
    )


def write_excel(
    asset_cleaned: pd.DataFrame,
    etf_cleaned: pd.DataFrame,
    index_master: pd.DataFrame,
    options: pd.DataFrame,
    mapping: pd.DataFrame,
    report: pd.DataFrame,
    review: pd.DataFrame,
) -> None:
    sheets = {
        "README": readme_df(),
        "asset_master_with_etf_index_cleaned": asset_cleaned,
        "etf_master_cleaned": etf_cleaned,
        "index_master": index_master,
        "options_master_with_etf_underlying": options,
        "etf_index_mapping": mapping,
        "etf_quality_classification_report": report,
        "etf_underlying_group_review_list": review,
    }
    if RAW_QUALITY_IN.exists():
        sheets["etf_index_quality_report"] = read_optional_csv(RAW_QUALITY_IN)
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
    etf_master = read_csv(ETF_MASTER_IN)
    asset_master = read_csv(ASSET_MASTER_IN)
    options = read_csv(OPTIONS_IN)
    mapping = read_csv(MAPPING_IN)
    index_master = read_csv(INDEX_MASTER_IN)

    raw_missing = int(etf_master["underlying_group"].apply(is_missing).sum())
    etf_cleaned = classify_etfs(etf_master, options)
    asset_cleaned = sync_asset_master(asset_master, etf_cleaned)
    review = build_review_list(etf_cleaned)
    report = build_report(etf_cleaned, review, raw_missing)
    write_excel(asset_cleaned, etf_cleaned, index_master, options, mapping, report, review)

    non_core_deferred = int(
        (etf_cleaned["underlying_group"].apply(is_missing) & (etf_cleaned["underlying_group_review_required"] == "N")).sum()
    )
    print(f"etf_master 原始行数: {len(etf_master)}")
    print(f"etf_master_cleaned 行数: {len(etf_cleaned)}")
    print(f"原始 underlying_group 缺失数量: {raw_missing}")
    print(f"需要 review 的 underlying_group 缺失数量: {len(review)}")
    print(f"非核心 deferred 数量: {non_core_deferred}")
    print(f"CORE ETF 数量: {int((etf_cleaned['etf_importance_level'] == 'CORE').sum())}")
    print(f"ETF期权标的 ETF 数量: {int((etf_cleaned['is_etf_option_underlying'] == 'Y').sum())}")
    print(f"review list 行数: {len(review)}")
    print(f"Excel 输出路径: {EXCEL_OUT}")


if __name__ == "__main__":
    main()
