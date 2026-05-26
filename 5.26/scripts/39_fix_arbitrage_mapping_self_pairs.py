from __future__ import annotations

import importlib.util
import shutil
import sys
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
V3_IN = ROOT / "output" / "mapping_info_full_v3.xlsx"
FULL_OUT = ROOT / "output" / "mapping_info_full.xlsx"
V4_OUT = ROOT / "output" / "mapping_info_full_v4.xlsx"
BACKUP_OUT = ROOT / "output" / "mapping_info_full_backup_before_v4.xlsx"
REPORT_OUT = ROOT / "output" / "mapping_info_full_v4_check_report.txt"

STEP9_PATH = ROOT / "scripts" / "36_optimize_mapping_info_full.py"
STEP10_PATH = ROOT / "scripts" / "37_rebuild_mapping_and_detail_tables.py"
STEP11_PATH = ROOT / "scripts" / "38_add_commodity_spot_layer.py"

spec9 = importlib.util.spec_from_file_location("step9", STEP9_PATH)
step9 = importlib.util.module_from_spec(spec9)
assert spec9 and spec9.loader
spec9.loader.exec_module(step9)

spec10 = importlib.util.spec_from_file_location("step10", STEP10_PATH)
step10 = importlib.util.module_from_spec(spec10)
assert spec10 and spec10.loader
spec10.loader.exec_module(step10)

spec11 = importlib.util.spec_from_file_location("step11", STEP11_PATH)
step11 = importlib.util.module_from_spec(spec11)
assert spec11 and spec11.loader
spec11.loader.exec_module(step11)

SHEET_ORDER = [
    "README", "relation_summary", "validation_summary", "arbitrage_mapping_pairs",
    "mapping_display_wide", "detail_display", "current_contract_detail", "contract_rule_detail",
    "core_asset_master", "asset_master", "futures_master", "options_master", "core_etf_master",
    "etf_master", "index_master", "spot_master", "spot_mapping", "underlying_mapping",
    "domestic_cross_market_mapping", "industry_chain_mapping", "industry_chain_asset_pool",
    "foreign_asset_master", "foreign_cross_market_mapping",
]


def norm(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def read_workbook(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(f"缺少输入工作簿: {path}")
    return pd.read_excel(path, sheet_name=None, dtype=str).copy()


def self_mask(pairs: pd.DataFrame) -> pd.Series:
    return pairs["base_future_asset_id"].map(norm).str.upper() == pairs["counterparty_asset_id"].map(norm).str.upper()


def invalid_counterparty_mask(pairs: pd.DataFrame) -> pd.Series:
    return pairs["counterparty_asset_id"].map(norm).isin(["", "空"])


def priority(role: str) -> int:
    order = {
        "OPTION_UNDERLYING": 0,
        "ETF_OPTION": 1,
        "ETF_SPOT": 2,
        "COMMODITY_SPOT": 3,
        "INDEX_SPOT": 4,
        "FOREIGN_FUTURE": 5,
        "FOREIGN_ETF": 6,
        "FOREIGN_REFERENCE": 7,
        "INDUSTRY_CHAIN_RELATED": 8,
        "DOMESTIC_CROSS_MARKET": 9,
        "BASE_FUTURE_REFERENCE": 99,
    }
    return order.get(norm(role), 50)


def sanitize_id(value: Any) -> str:
    text = norm(value)
    for old, new in {" ": "_", "，": "_", "；": "_", ",": "_", ";": "_"}.items():
        text = text.replace(old, new)
    return text


def rebuild_pair_id(pairs: pd.DataFrame) -> pd.DataFrame:
    out = pairs.copy()
    out["pair_id"] = out.apply(
        lambda r: "PAIR_{}_{}_{}_{}_{}".format(
            sanitize_id(r["relation_source"]),
            sanitize_id(r["base_future_asset_id"]),
            sanitize_id(r["counterparty_asset_id"]),
            sanitize_id(r["relation_type"]),
            sanitize_id(r["mapping_role"]),
        ),
        axis=1,
    )
    dup = out["pair_id"].duplicated(keep=False)
    if dup.any():
        out.loc[dup, "pair_id"] = out.loc[dup, "pair_id"] + "_" + out.groupby("pair_id").cumcount().astype(str)
    return out


def fix_pairs(pairs: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    before_self = int(self_mask(pairs).sum())
    out = pairs.loc[~self_mask(pairs) & ~invalid_counterparty_mask(pairs)].copy()
    out["_priority"] = out["mapping_role"].map(priority)
    out = out.sort_values("_priority")
    out = out.drop_duplicates(
        ["base_future_asset_id", "counterparty_asset_id", "relation_source", "relation_type", "strategy_type", "mapping_role"],
        keep="first",
    )
    out = out.drop(columns=["_priority"])
    out = rebuild_pair_id(out)
    after_self = int(self_mask(out).sum())
    return out.reset_index(drop=True), before_self, after_self


def relation_type_cn(value: str) -> str:
    if value == "FUTURE_SPOT":
        return "期货-商品现货"
    return step9.REL_CN.get(value, value)


def relation_summary(pairs: pd.DataFrame) -> pd.DataFrame:
    grouped = pairs.groupby(["relation_source", "relation_type", "strategy_type", "mapping_role"], dropna=False).size().reset_index(name="count")
    grouped["relation_type_cn"] = grouped["relation_type"].map(relation_type_cn)
    return grouped


def mapping_display_wide(pairs: pd.DataFrame, futures: pd.DataFrame) -> pd.DataFrame:
    return step11.mapping_display_wide_v3(pairs, futures)


def non_empty_count(df: pd.DataFrame, col: str) -> int:
    if col not in df.columns:
        return 0
    return int(~df[col].astype(str).str.strip().isin(["", "空"]).sum()) if False else int((~df[col].astype(str).str.strip().isin(["", "空"])).sum())


def multi_asset_count(pairs: pd.DataFrame) -> int:
    return step10.has_multi_asset(pairs)


def validation_summary(sheets: dict[str, pd.DataFrame], pairs: pd.DataFrame) -> pd.DataFrame:
    futures = sheets["futures_master"]
    options = sheets["options_master"]
    detail = sheets["detail_display"]
    contract = sheets["contract_rule_detail"]
    spot_mapping = sheets["spot_mapping"]
    commodity_count = len(spot_mapping)
    spot_coverage = spot_mapping["future_asset_id"].nunique() if "future_asset_id" in spot_mapping else 0
    spot_rate = round(spot_coverage / commodity_count, 4) if commodity_count else 0
    residual = 0
    for name, df in sheets.items():
        if name == "validation_summary" or df.empty:
            continue
        residual += int(df.astype(str).apply(lambda s: s.str.contains("CHECK|TODO|NEED_REVIEW", regex=True, na=False)).sum().sum())
    etf_missing = 0
    if {"subtype", "underlying_asset_id"}.issubset(options.columns):
        mask = options["subtype"].eq("ETF_OPTION")
        etf_missing = int(options.loc[mask, "underlying_asset_id"].astype(str).str.strip().isin(["", "空"]).sum())
    check_wait = int(options.astype(str).apply(lambda s: s.str.contains("CHECK_WAIT_ETF_MASTER", regex=False, na=False)).sum().sum()) if not options.empty else 0
    contract_src = int(contract.astype(str).apply(lambda s: s.str.contains("AKShare|OpenCTP", regex=True, na=False)).sum().sum()) if not contract.empty else 0
    rows = [
        ("arbitrage_mapping_pairs 行数", "PASS" if len(pairs) else "FAIL", len(pairs), "套利映射长表记录数。"),
        ("pair_id 重复数", "PASS" if not pairs["pair_id"].duplicated().any() else "FAIL", int(pairs["pair_id"].duplicated().sum()), "pair_id 必须唯一。"),
        ("self_mapping_count", "PASS" if int(self_mask(pairs).sum()) == 0 else "FAIL", int(self_mask(pairs).sum()), "base_future_asset_id 不得等于 counterparty_asset_id。"),
        ("base_future 覆盖数量", "PASS" if pairs["base_future_asset_id"].nunique() >= 98 else "FAIL", pairs["base_future_asset_id"].nunique(), "应覆盖 98 个国内期货。"),
        ("是否存在一个格子多个资产", "PASS" if multi_asset_count(pairs) == 0 else "FAIL", multi_asset_count(pairs), "counterparty 字段不得拼接多个资产。"),
        ("SPOT 映射行数", "PASS" if int((pairs["relation_source"] == "SPOT").sum()) else "FAIL", int((pairs["relation_source"] == "SPOT").sum()), "商品现货映射行数。"),
        ("商品期货 spot_mapping 覆盖数量", "PASS", spot_coverage, "覆盖到 spot_mapping 的商品期货数量。"),
        ("商品期货 spot_mapping 覆盖率", "PASS" if spot_rate >= 0.8 else "FAIL", spot_rate, "覆盖率低于 80% 为 FAIL。"),
        ("detail_display 行数", "PASS" if len(detail) >= 98 else "FAIL", len(detail), "详情展示表至少覆盖 98 个期货。"),
        ("detail_display contract_multiplier 非空数量", "PASS" if non_empty_count(detail, "contract_multiplier") >= 78 else "WARN", non_empty_count(detail, "contract_multiplier"), "低于 80% 为 WARN。"),
        ("detail_display tick_size 非空数量", "PASS" if non_empty_count(detail, "tick_size") >= 78 else "WARN", non_empty_count(detail, "tick_size"), "低于 80% 为 WARN。"),
        ("detail_display contract_unit 非空数量", "PASS" if non_empty_count(detail, "contract_unit") >= 78 else "WARN", non_empty_count(detail, "contract_unit"), "低于 80% 为 WARN。"),
        ("futures_master underlying_group 空值数量", "PASS" if int(futures["underlying_group"].astype(str).str.strip().isin(["", "空"]).sum()) == 0 else "FAIL", int(futures["underlying_group"].astype(str).str.strip().isin(["", "空"]).sum()), "期货展示表 underlying_group 应不为空。"),
        ("options_master CHECK_WAIT_ETF_MASTER 残留数量", "PASS" if check_wait == 0 else "FAIL", check_wait, "ETF期权标的应使用最终修复表。"),
        ("ETF_OPTION underlying_asset_id 空值数量", "PASS" if etf_missing == 0 else "FAIL", etf_missing, "ETF_OPTION 标的资产不能为空。"),
        ("contract_rule_detail AKShare/OpenCTP 残留数量", "PASS" if contract_src == 0 else "FAIL", contract_src, "规则展示表不应保留过程来源字样。"),
        ("CHECK/TODO/NEED_REVIEW 残留数量", "PASS" if residual == 0 else "WARN", residual, "展示层无效占位残留。"),
    ]
    return pd.DataFrame(rows, columns=["check_item", "result", "count", "description"])


def report(before_rows: int, before_self: int, after_rows: int, after_self: int, pairs: pd.DataFrame, spot_mapping: pd.DataFrame, val: pd.DataFrame, deleted: int) -> str:
    spot_rate = spot_mapping["future_asset_id"].nunique() / len(spot_mapping) if len(spot_mapping) else 0
    fail = int((val["result"] == "FAIL").sum())
    warn = int((val["result"] == "WARN").sum())
    conclusion = "NEED_MAJOR_FIX" if fail else ("NEED_MINOR_FIX" if warn else "PASS")
    lines = [
        f"修复前 arbitrage_mapping_pairs 行数: {before_rows}",
        f"修复前 self_mapping_count: {before_self}",
        f"修复后 arbitrage_mapping_pairs 行数: {after_rows}",
        f"修复后 self_mapping_count: {after_self}",
        f"删除 self mapping 行数: {deleted}",
        f"pair_id 重复数: {int(pairs['pair_id'].duplicated().sum())}",
        f"base_future 覆盖数量: {pairs['base_future_asset_id'].nunique()}",
        f"SPOT 映射行数: {int((pairs['relation_source'] == 'SPOT').sum())}",
        f"商品期货 spot_mapping 覆盖率: {spot_rate:.2%}",
        f"是否存在多资产塞一个格子: {'Y' if multi_asset_count(pairs) else 'N'}",
        f"validation_summary FAIL 数量: {fail}",
        f"validation_summary WARN 数量: {warn}",
        f"最终结论: {conclusion}",
    ]
    return "\n".join(lines) + "\n"


def main() -> None:
    try:
        if FULL_OUT.exists():
            shutil.copy2(FULL_OUT, BACKUP_OUT)
        sheets = read_workbook(V3_IN if V3_IN.exists() else FULL_OUT)
        pairs_before = sheets["arbitrage_mapping_pairs"].fillna("空")
        before_rows = len(pairs_before)
        pairs_fixed, before_self, after_self = fix_pairs(pairs_before)

        sheets["arbitrage_mapping_pairs"] = pairs_fixed
        sheets["relation_summary"] = relation_summary(pairs_fixed)
        sheets["mapping_display_wide"] = mapping_display_wide(pairs_fixed, sheets["futures_master"].fillna("空"))
        sheets["validation_summary"] = validation_summary(sheets, pairs_fixed)

        ordered = {name: sheets[name] for name in SHEET_ORDER if name in sheets}
        step9.write_excel(FULL_OUT, ordered)
        step9.write_excel(V4_OUT, ordered)
        deleted = before_rows - len(pairs_fixed)
        REPORT_OUT.write_text(report(before_rows, before_self, len(pairs_fixed), after_self, pairs_fixed, sheets["spot_mapping"], sheets["validation_summary"], deleted), encoding="utf-8")

        fail = int((sheets["validation_summary"]["result"] == "FAIL").sum())
        warn = int((sheets["validation_summary"]["result"] == "WARN").sum())
        conclusion = "NEED_MAJOR_FIX" if fail else ("NEED_MINOR_FIX" if warn else "PASS")
        print(f"mapping_info_full.xlsx 是否更新: {'Y' if FULL_OUT.exists() else 'N'}")
        print(f"mapping_info_full_v4.xlsx 是否生成: {'Y' if V4_OUT.exists() else 'N'}")
        print(f"修复前 arbitrage_mapping_pairs 行数: {before_rows}")
        print(f"修复前 self_mapping_count: {before_self}")
        print(f"修复后 arbitrage_mapping_pairs 行数: {len(pairs_fixed)}")
        print(f"修复后 self_mapping_count: {after_self}")
        print(f"删除 self mapping 行数: {deleted}")
        print(f"pair_id 重复数: {int(pairs_fixed['pair_id'].duplicated().sum())}")
        print(f"base_future 覆盖数量: {pairs_fixed['base_future_asset_id'].nunique()}")
        print(f"SPOT 映射行数: {int((pairs_fixed['relation_source'] == 'SPOT').sum())}")
        print(f"商品期货 spot_mapping 覆盖率: {sheets['spot_mapping']['future_asset_id'].nunique()/len(sheets['spot_mapping']):.2%}")
        print(f"是否存在多资产塞一个格子: {'Y' if multi_asset_count(pairs_fixed) else 'N'}")
        print(f"validation_summary FAIL 数量: {fail}")
        print(f"validation_summary WARN 数量: {warn}")
        print(f"最终结论: {conclusion}")
        print(f"检查报告路径: {REPORT_OUT}")
    except Exception as exc:
        print(f"Step 12 自映射修复失败: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
