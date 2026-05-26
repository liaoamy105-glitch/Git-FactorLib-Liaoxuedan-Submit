from __future__ import annotations

import importlib.util
import shutil
import sys
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
STEP9_PATH = ROOT / "scripts" / "36_optimize_mapping_info_full.py"
spec = importlib.util.spec_from_file_location("step9", STEP9_PATH)
step9 = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(step9)

FULL_OUT = ROOT / "output" / "mapping_info_full.xlsx"
V2_OUT = ROOT / "output" / "mapping_info_full_v2.xlsx"
BACKUP_OUT = ROOT / "output" / "mapping_info_full_backup_before_v2.xlsx"
REPORT_OUT = ROOT / "output" / "mapping_info_full_v2_check_report.txt"


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def invalid(value: Any) -> bool:
    return step9.is_invalid(value) or clean_text(value) in {"空"}


def first_valid(*values: Any) -> str:
    for value in values:
        text = clean_text(value)
        if text and not invalid(text):
            return text
    return ""


def clean_sheet(sheet: str, df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    clean, dropped = step9.clean_sheet(sheet, df)
    # Step 9's process-field cleanup intentionally drops confidence-like columns. In v2,
    # mapping_confidence is a core modeling field for arbitrage_mapping_pairs.
    if sheet == "arbitrage_mapping_pairs" and "mapping_confidence" not in clean.columns and "mapping_confidence" in df.columns:
        clean["mapping_confidence"] = df["mapping_confidence"].map(step9.clean_value)
    return step9.order_cols(sheet, clean), dropped


def backfill_futures_from_rules(futures: pd.DataFrame, rules: pd.DataFrame) -> pd.DataFrame:
    out = futures.copy()
    rule_cols = [
        "asset_id", "contract_multiplier", "contract_unit", "tick_size", "quote_unit",
        "settlement_type", "trading_hours", "night_trading", "contract_month_rule",
        "last_trading_day_rule", "delivery_day_rule", "rule_parseable", "sample_last_ddate",
    ]
    available = [c for c in rule_cols if c in rules.columns]
    rule = rules[available].drop_duplicates("asset_id")
    out = out.merge(rule, on="asset_id", how="left", suffixes=("", "_rule"))
    for col in rule_cols:
        rc = col + "_rule"
        if rc not in out.columns:
            continue
        if col not in out.columns:
            out[col] = ""
        mask = out[col].map(invalid) | out[col].astype(str).str.strip().eq("")
        out.loc[mask, col] = out.loc[mask, rc]
        out = out.drop(columns=[rc])
    return out


def contract_code(ts_code: str, fallback: str) -> str:
    ts = clean_text(ts_code)
    if ts:
        return ts.split(".")[0]
    fb = clean_text(fallback)
    return fb.split(".")[0] if fb else ""


def detail_display(futures: pd.DataFrame, rules: pd.DataFrame, options: pd.DataFrame) -> pd.DataFrame:
    fut = backfill_futures_from_rules(futures, rules)
    rows = []
    for _, row in fut.iterrows():
        ts = first_valid(row.get("sample_ts_code"), row.get("sample_contract_name"), row.get("symbol"))
        code = contract_code(ts, row.get("sample_contract_name", ""))
        name = first_valid(row.get("sample_contract_name"), f"{row.get('name_cn', '')}{code}", row.get("name_cn"))
        rows.append({
            "symbol": row.get("symbol", ""),
            "asset_type": "FUTURE",
            "ts_code": ts,
            "contract_code": code,
            "contract_name": name,
            "exchange_code": row.get("exchange_code", ""),
            "exchange_name": row.get("exchange_name", ""),
            "name_cn": row.get("name_cn", ""),
            "contract_unit": row.get("contract_unit", ""),
            "contract_multiplier": row.get("contract_multiplier", ""),
            "quote_unit": row.get("quote_unit", ""),
            "tick_size": row.get("tick_size", ""),
            "settlement_type": row.get("settlement_type", ""),
            "list_date": row.get("list_date_min", ""),
            "delist_date": row.get("delist_date_max", ""),
            "last_trading_day": first_valid(row.get("sample_last_ddate"), row.get("delist_date_max")),
            "last_trading_day_rule": row.get("last_trading_day_rule", ""),
            "delivery_day_rule": row.get("delivery_day_rule", ""),
            "contract_month_rule": row.get("contract_month_rule", ""),
            "trading_hours": row.get("trading_hours", ""),
            "night_trading": row.get("night_trading", ""),
            "underlying_group": row.get("underlying_group", ""),
            "subtype": row.get("subtype", ""),
            "country": row.get("country", ""),
            "currency": row.get("currency", ""),
        })
    # Append option product-level rows where useful. They are supplemental and do not
    # affect the 98-future coverage checks.
    for _, row in options.iterrows():
        rows.append({
            "symbol": row.get("option_symbol", row.get("symbol", "")),
            "asset_type": "OPTION",
            "ts_code": row.get("option_symbol", row.get("symbol", "")),
            "contract_code": row.get("option_symbol", row.get("symbol", "")),
            "contract_name": row.get("option_name_cn", row.get("name_cn", "")),
            "exchange_code": row.get("exchange_code", ""),
            "exchange_name": row.get("exchange_name", ""),
            "name_cn": row.get("option_name_cn", row.get("name_cn", "")),
            "contract_unit": row.get("contract_unit", ""),
            "contract_multiplier": row.get("contract_multiplier", ""),
            "quote_unit": row.get("quote_unit", ""),
            "tick_size": row.get("tick_size", ""),
            "settlement_type": row.get("settlement_type", ""),
            "list_date": "",
            "delist_date": "",
            "last_trading_day": "",
            "last_trading_day_rule": row.get("last_trading_day_rule", ""),
            "delivery_day_rule": row.get("delivery_day_rule", ""),
            "contract_month_rule": row.get("listed_contract_rule", ""),
            "trading_hours": "",
            "night_trading": "",
            "underlying_group": row.get("underlying_group", ""),
            "subtype": row.get("subtype", ""),
            "country": row.get("country", ""),
            "currency": row.get("currency", ""),
        })
    return pd.DataFrame(rows)


def current_contract_detail(futures: pd.DataFrame, rules: pd.DataFrame) -> pd.DataFrame:
    fut = backfill_futures_from_rules(futures, rules)
    rows = []
    for _, row in fut.iterrows():
        ts = first_valid(row.get("sample_ts_code"), row.get("sample_contract_name"), row.get("symbol"))
        code = contract_code(ts, row.get("sample_contract_name", ""))
        rows.append({
            "asset_id": row.get("asset_id", ""),
            "symbol": row.get("symbol", ""),
            "name_cn": row.get("name_cn", ""),
            "asset_type": row.get("asset_type", ""),
            "subtype": row.get("subtype", ""),
            "exchange_code": row.get("exchange_code", ""),
            "exchange_name": row.get("exchange_name", ""),
            "ts_code": ts,
            "contract_code": code,
            "contract_name": first_valid(row.get("sample_contract_name"), f"{row.get('name_cn', '')}{code}", row.get("name_cn")),
            "contract_unit": row.get("contract_unit", ""),
            "contract_multiplier": row.get("contract_multiplier", ""),
            "quote_unit": row.get("quote_unit", ""),
            "tick_size": row.get("tick_size", ""),
            "settlement_type": row.get("settlement_type", ""),
            "list_date": row.get("list_date_min", ""),
            "delist_date": row.get("delist_date_max", ""),
            "last_trading_day": first_valid(row.get("sample_last_ddate"), row.get("delist_date_max")),
            "contract_month_rule": row.get("contract_month_rule", ""),
            "last_trading_day_rule": row.get("last_trading_day_rule", ""),
            "trading_hours": row.get("trading_hours", ""),
            "night_trading": row.get("night_trading", ""),
        })
    return pd.DataFrame(rows)


def relation_desc(role: str, relation_type: str) -> str:
    if role == "OPTION_UNDERLYING":
        return "期权与对应标的之间的映射，可用于期权套利和定价分析。"
    if role == "ETF_OPTION":
        return "ETF期权与同一统一标的期货之间的映射，可用于期权和标的联动分析。"
    if role == "ETF_SPOT":
        return "期货与同一标的 ETF 之间的映射，可用于期现或基差分析。"
    if role == "INDEX_SPOT":
        return "期货与指数或现货参考之间的映射，可用于基差和指数参考分析。"
    if role == "FOREIGN_FUTURE":
        return "国内期货与国外同标的期货之间的映射，可用于内外盘套利。"
    if role in {"FOREIGN_ETF", "FOREIGN_REFERENCE"}:
        return "国内期货与国外 ETF 或参考资产之间的映射，可用于跨市场联动观察。"
    if role == "INDUSTRY_CHAIN_RELATED":
        return "产业链相关品种之间的映射，可用于跨品种套利。"
    return step9.DOM_DESC.get(relation_type, "映射候选关系，可用于套利组合构建。")


def pair_record(base: pd.Series, counter: dict[str, Any], source: str, role: str, relation_type: str, strategy: str, confidence: str = "HIGH", direction: str = "BIDIRECTIONAL", tradable: str = "", fx: str = "N", unit: str = "", note: str = "") -> dict[str, Any]:
    cid = clean_text(counter.get("asset_id"))
    return {
        "pair_id": f"PAIR_{source}_{base.get('asset_id')}_{cid}_{relation_type}",
        "base_future_asset_id": base.get("asset_id", ""),
        "base_future_symbol": base.get("symbol", ""),
        "base_future_name": base.get("name_cn", ""),
        "base_exchange_code": base.get("exchange_code", ""),
        "base_exchange_name": base.get("exchange_name", ""),
        "underlying_group": base.get("underlying_group", ""),
        "counterparty_asset_id": cid,
        "counterparty_symbol": counter.get("symbol", ""),
        "counterparty_name": counter.get("name", ""),
        "counterparty_asset_type": counter.get("asset_type", ""),
        "counterparty_subtype": counter.get("subtype", ""),
        "counterparty_exchange_code": counter.get("exchange_code", ""),
        "counterparty_exchange_name": counter.get("exchange_name", ""),
        "counterparty_country": counter.get("country", ""),
        "counterparty_currency": counter.get("currency", ""),
        "relation_source": source,
        "relation_type": relation_type,
        "strategy_type": strategy,
        "mapping_role": role,
        "mapping_confidence": confidence,
        "direction_supported": direction,
        "tradable_check": tradable,
        "fx_conversion_needed": fx,
        "unit_conversion": unit,
        "relation_desc_cn": relation_desc(role, relation_type),
        "long_short_note": note,
    }


def build_arbitrage_pairs(futures: pd.DataFrame, options: pd.DataFrame, etf_core: pd.DataFrame, etf_all: pd.DataFrame, index: pd.DataFrame, domestic: pd.DataFrame, industry: pd.DataFrame, foreign: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, base in futures.iterrows():
        aid, ug = base.get("asset_id", ""), base.get("underlying_group", "")
        opt_mask = (options.get("underlying_asset_id", "") == aid) | ((options.get("underlying_group", "") == ug) & options.get("subtype", "").astype(str).str.contains("COMMODITY_OPTION|INDEX_OPTION", na=False))
        for _, r in options.loc[opt_mask].iterrows():
            rows.append(pair_record(base, {"asset_id": r.get("asset_id"), "symbol": r.get("option_symbol"), "name": r.get("option_name_cn"), "asset_type": "OPTION", "subtype": r.get("subtype"), "exchange_code": r.get("exchange_code"), "exchange_name": r.get("exchange_name"), "country": r.get("country"), "currency": r.get("currency")}, "OPTION", "OPTION_UNDERLYING", "OPTION_UNDERLYING", "OPTION_FUTURE_ARBITRAGE" if "COMMODITY" in r.get("subtype", "") else "OPTION_ARBITRAGE"))
        etf_opt = options[(options.get("subtype", "") == "ETF_OPTION") & (options.get("underlying_group", "") == ug)]
        for _, r in etf_opt.iterrows():
            rows.append(pair_record(base, {"asset_id": r.get("asset_id"), "symbol": r.get("option_symbol"), "name": r.get("option_name_cn"), "asset_type": "OPTION", "subtype": r.get("subtype"), "exchange_code": r.get("exchange_code"), "exchange_name": r.get("exchange_name"), "country": r.get("country"), "currency": r.get("currency")}, "OPTION", "ETF_OPTION", "OPTION_UNDERLYING", "OPTION_ARBITRAGE"))
        etfs = (etf_core if not etf_core.empty else etf_all)
        for _, r in etfs[etfs.get("underlying_group", "") == ug].iterrows():
            rows.append(pair_record(base, {"asset_id": r.get("asset_id"), "symbol": r.get("etf_code", r.get("symbol", "")), "name": r.get("name_cn"), "asset_type": "ETF", "subtype": r.get("subtype"), "exchange_code": r.get("exchange_code"), "exchange_name": r.get("exchange_name"), "country": r.get("country"), "currency": r.get("currency")}, "ETF", "ETF_SPOT", "FUTURE_ETF", "BASIS_ARBITRAGE"))
        for _, r in index[index.get("underlying_group", "") == ug].iterrows():
            rows.append(pair_record(base, {"asset_id": r.get("asset_id"), "symbol": r.get("index_code", r.get("symbol", "")), "name": r.get("name_cn"), "asset_type": "SPOT", "subtype": "INDEX_SPOT", "exchange_code": r.get("exchange_code"), "exchange_name": r.get("exchange_name"), "country": r.get("country"), "currency": r.get("currency")}, "INDEX", "INDEX_SPOT", "FUTURE_INDEX", "INDEX_REFERENCE"))
        dom = domestic[(domestic.get("asset_id_a", "") == aid) | (domestic.get("asset_id_b", "") == aid)]
        for _, r in dom.iterrows():
            side_b = r.get("asset_id_a") == aid
            rows.append(pair_record(base, {"asset_id": r.get("asset_id_b" if not side_b else "asset_id_a"), "symbol": r.get("symbol_b" if not side_b else "symbol_a"), "name": r.get("name_b" if not side_b else "name_a"), "asset_type": r.get("asset_type_b" if not side_b else "asset_type_a"), "subtype": "", "exchange_code": r.get("exchange_b" if not side_b else "exchange_a"), "exchange_name": "", "country": "CN", "currency": "CNY"}, "DOMESTIC", "DOMESTIC_CROSS_MARKET", r.get("relation_type"), r.get("strategy_type"), r.get("mapping_confidence", "HIGH"), r.get("direction_supported", ""), r.get("tradable_check", ""), note=r.get("long_short_note", "")))
        for _, r in foreign[foreign.get("domestic_asset_id", "") == aid].iterrows():
            ftype = r.get("foreign_asset_type")
            role = "FOREIGN_FUTURE" if ftype == "FUTURE" else ("FOREIGN_ETF" if ftype == "ETF" else "FOREIGN_REFERENCE")
            rows.append(pair_record(base, {"asset_id": r.get("foreign_asset_id"), "symbol": r.get("foreign_symbol"), "name": r.get("foreign_name"), "asset_type": ftype, "subtype": "", "exchange_code": r.get("foreign_exchange"), "exchange_name": r.get("foreign_exchange"), "country": r.get("foreign_country"), "currency": r.get("currency_foreign")}, "FOREIGN", role, r.get("relation_type"), r.get("strategy_type"), r.get("mapping_confidence"), r.get("direction_supported"), "", r.get("fx_conversion_needed"), r.get("unit_conversion")))
        ind = industry[(industry.get("asset_id_a", "") == aid) | (industry.get("asset_id_b", "") == aid)]
        for _, r in ind.iterrows():
            side_b = r.get("asset_id_a") == aid
            rows.append(pair_record(base, {"asset_id": r.get("asset_id_b" if not side_b else "asset_id_a"), "symbol": r.get("symbol_b" if not side_b else "symbol_a"), "name": r.get("name_b" if not side_b else "name_a"), "asset_type": r.get("asset_type_b" if not side_b else "asset_type_a"), "subtype": "", "exchange_code": r.get("exchange_b" if not side_b else "exchange_a"), "exchange_name": "", "country": "CN", "currency": "CNY"}, "INDUSTRY_CHAIN", "INDUSTRY_CHAIN_RELATED", r.get("relation_type"), r.get("strategy_type"), first_valid(r.get("mapping_confidence"), r.get("relation_strength"), "MEDIUM"), r.get("direction_supported"), r.get("tradable_check"), note=r.get("long_short_note", "")))
    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame()
    covered = set(df["base_future_asset_id"]) if not df.empty else set()
    fallback_rows = []
    for _, base in futures.iterrows():
        if base.get("asset_id") in covered:
            continue
        fallback_rows.append(pair_record(
            base,
            {
                "asset_id": base.get("asset_id"),
                "symbol": base.get("symbol"),
                "name": base.get("name_cn"),
                "asset_type": "FUTURE",
                "subtype": base.get("subtype"),
                "exchange_code": base.get("exchange_code"),
                "exchange_name": base.get("exchange_name"),
                "country": base.get("country", "CN"),
                "currency": base.get("currency", "CNY"),
            },
            "DOMESTIC",
            "BASE_FUTURE_REFERENCE",
            "BASE_FUTURE_REFERENCE",
            "REFERENCE_ONLY",
            "LOW",
            "REFERENCE_ONLY",
            base.get("tradable", ""),
            "N",
            "",
            "暂无独立对手资产映射，保留基础期货参考行，便于程序覆盖全部期货品种。",
        ))
    if fallback_rows:
        df = pd.concat([df, pd.DataFrame(fallback_rows)], ignore_index=True)
    if df.empty:
        return df
    priority = {"OPTION_UNDERLYING": 0, "ETF_OPTION": 1, "ETF_SPOT": 2, "INDEX_SPOT": 3, "DOMESTIC_CROSS_MARKET": 4, "FOREIGN_FUTURE": 5, "FOREIGN_ETF": 6, "FOREIGN_REFERENCE": 7, "INDUSTRY_CHAIN_RELATED": 8}
    df["_p"] = df["mapping_role"].map(priority).fillna(99)
    df = df.sort_values("_p").drop_duplicates(["base_future_asset_id", "counterparty_asset_id", "relation_source", "relation_type", "strategy_type"]).drop(columns=["_p"])
    df["pair_id"] = df["pair_id"].where(~df["pair_id"].duplicated(), df["pair_id"] + "_" + df.groupby("pair_id").cumcount().astype(str))
    return df.sort_values(["base_exchange_code", "base_future_symbol", "relation_source", "mapping_role", "counterparty_asset_type", "counterparty_symbol"]).reset_index(drop=True)


def mapping_display_wide(pairs: pd.DataFrame, futures: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, f in futures.sort_values(["exchange_code", "symbol"]).iterrows():
        p = pairs[pairs["base_future_asset_id"] == f["asset_id"]]
        def sample(mask):
            vals = p.loc[mask, "counterparty_name"].tolist()
            return vals[:2] + [""] * max(0, 2 - len(vals))
        foreign_f = sample(p["mapping_role"].eq("FOREIGN_FUTURE"))
        etfs = sample(p["mapping_role"].eq("ETF_SPOT"))
        inds = sample(p["mapping_role"].eq("INDUSTRY_CHAIN_RELATED"))
        rows.append({
            "期货": f"{f['symbol']}_{f['exchange_code']}_中国",
            "商品名称": f["name_cn"],
            "交易所": f["exchange_name"],
            "统一标的": f["underlying_group"],
            "期权数量": int((p["mapping_role"] == "OPTION_UNDERLYING").sum()),
            "ETF期权数量": int((p["mapping_role"] == "ETF_OPTION").sum()),
            "ETF现货数量": int((p["mapping_role"] == "ETF_SPOT").sum()),
            "指数/现货数量": int((p["mapping_role"] == "INDEX_SPOT").sum()),
            "国内跨市场关系数量": int((p["relation_source"] == "DOMESTIC").sum()),
            "国外期货数量": int((p["mapping_role"] == "FOREIGN_FUTURE").sum()),
            "国外ETF/参考资产数量": int(p["mapping_role"].isin(["FOREIGN_ETF", "FOREIGN_REFERENCE"]).sum()),
            "产业链相关品种数量": int((p["mapping_role"] == "INDUSTRY_CHAIN_RELATED").sum()),
            "主要国外期货1": foreign_f[0], "主要国外期货2": foreign_f[1],
            "主要ETF现货1": etfs[0], "主要ETF现货2": etfs[1],
            "主要产业链品种1": inds[0], "主要产业链品种2": inds[1],
            "支持策略数量": p["strategy_type"].nunique(),
        })
    return pd.DataFrame(rows)


def relation_summary_from_pairs(pairs: pd.DataFrame) -> pd.DataFrame:
    if pairs.empty:
        return pd.DataFrame(columns=["relation_source", "relation_type", "strategy_type", "mapping_role", "count", "relation_type_cn"])
    out = pairs.groupby(["relation_source", "relation_type", "strategy_type", "mapping_role"], dropna=False).size().reset_index(name="count")
    out["relation_type_cn"] = out["relation_type"].map(lambda x: step9.REL_CN.get(x, x))
    return out


def has_multi_asset(pairs: pd.DataFrame) -> int:
    if pairs.empty:
        return 0
    total = 0
    for col in ["counterparty_asset_id", "counterparty_symbol", "counterparty_name"]:
        total += int(pairs[col].astype(str).str.contains("；|,", regex=True, na=False).sum())
    return total


def validation(pairs: pd.DataFrame, futures: pd.DataFrame, detail: pd.DataFrame, options: pd.DataFrame, domestic: pd.DataFrame, industry: pd.DataFrame, foreign: pd.DataFrame, contract: pd.DataFrame, residual_invalid: int) -> pd.DataFrame:
    base_cov = pairs["base_future_asset_id"].nunique() if not pairs.empty else 0
    multi = has_multi_asset(pairs)
    fut_ug_empty = int(futures["underlying_group"].map(invalid).sum())
    check_wait = int(options.astype(str).apply(lambda s: s.str.contains("CHECK_WAIT_ETF_MASTER", regex=False, na=False)).sum().sum())
    etf_missing = int(((options["subtype"] == "ETF_OPTION") & (options["underlying_asset_id"].map(invalid) | options["underlying_asset_id"].astype(str).str.strip().eq(""))).sum()) if {"subtype", "underlying_asset_id"}.issubset(options.columns) else 0
    src_resid = int(contract.astype(str).apply(lambda s: s.str.contains("AKShare|OpenCTP", regex=True, na=False)).sum().sum()) if not contract.empty else 0
    metrics = [
        ("arbitrage_mapping_pairs 行数", "PASS", len(pairs), "套利映射长表记录数。"),
        ("arbitrage_mapping_pairs pair_id 重复数", "PASS" if not pairs["pair_id"].duplicated().any() else "FAIL", int(pairs["pair_id"].duplicated().sum()), "pair_id 必须唯一。"),
        ("arbitrage_mapping_pairs 是否存在一个单元格多个资产拼接", "PASS" if multi == 0 else "FAIL", multi, "counterparty 字段不得包含多个资产拼接。"),
        ("base_future 覆盖数量", "PASS" if base_cov >= 98 else "FAIL", base_cov, "应覆盖 98 个国内期货品种。"),
        ("detail_display 行数", "PASS" if len(detail) >= 98 else "FAIL", len(detail), "至少覆盖 98 个国内期货品种。"),
        ("detail_display contract_multiplier 非空数量", "PASS" if (detail["contract_multiplier"] != "空").sum() >= 78 else "WARN", int((detail["contract_multiplier"] != "空").sum()), "低于 80% 为 WARN。"),
        ("detail_display tick_size 非空数量", "PASS" if (detail["tick_size"] != "空").sum() >= 78 else "WARN", int((detail["tick_size"] != "空").sum()), "低于 80% 为 WARN。"),
        ("detail_display contract_unit 非空数量", "PASS" if (detail["contract_unit"] != "空").sum() >= 78 else "WARN", int((detail["contract_unit"] != "空").sum()), "低于 80% 为 WARN。"),
        ("futures_master underlying_group 空值数量", "PASS" if fut_ug_empty == 0 else "FAIL", fut_ug_empty, "期货展示表 underlying_group 应已同步修复。"),
        ("options_master CHECK_WAIT_ETF_MASTER 残留数量", "PASS" if check_wait == 0 else "FAIL", check_wait, "ETF期权标的应使用最终修复表。"),
        ("ETF_OPTION underlying_asset_id 空值数量", "PASS" if etf_missing == 0 else "FAIL", etf_missing, "ETF_OPTION 标的资产不能为空。"),
        ("contract_rule_detail AKShare/OpenCTP 残留数量", "PASS" if src_resid == 0 else "FAIL", src_resid, "规则展示表不应保留过程来源字样。"),
        ("domestic mapping_id 重复数", "PASS" if not domestic["mapping_id"].duplicated().any() else "FAIL", int(domestic["mapping_id"].duplicated().sum()), "国内映射 ID 重复数。"),
        ("industry 自映射数量", "PASS" if int((industry["asset_id_a"] == industry["asset_id_b"]).sum()) == 0 else "FAIL", int((industry["asset_id_a"] == industry["asset_id_b"]).sum()), "产业链自映射数量。"),
        ("foreign mapping_id 重复数", "PASS" if not foreign["mapping_id"].duplicated().any() else "FAIL", int(foreign["mapping_id"].duplicated().sum()), "国外映射 ID 重复数。"),
        ("CHECK/TODO/NEED_REVIEW 残留数量", "PASS" if residual_invalid == 0 else "WARN", residual_invalid, "展示层无效占位残留。"),
    ]
    return pd.DataFrame(metrics, columns=["check_item", "result", "count", "description"])


def report(metrics: dict[str, Any], val: pd.DataFrame, conclusion: str) -> str:
    lines = [
        f"输出文件路径: {FULL_OUT}",
        f"v2 文件路径: {V2_OUT}",
        f"arbitrage_mapping_pairs 行数: {metrics['pairs_rows']}",
        f"pair_id 重复数: {metrics['pair_dup']}",
        f"是否存在多资产塞一个格子: {'Y' if metrics['multi_asset'] else 'N'}",
        f"base_future 覆盖数量: {metrics['base_coverage']}",
        f"detail_display 行数: {metrics['detail_rows']}",
        f"detail_display contract_multiplier 非空数量: {metrics['mult_non_empty']}",
        f"detail_display tick_size 非空数量: {metrics['tick_non_empty']}",
        f"detail_display contract_unit 非空数量: {metrics['unit_non_empty']}",
        f"futures_master underlying_group 空值数量: {metrics['fut_ug_empty']}",
        f"options_master CHECK_WAIT_ETF_MASTER 残留数量: {metrics['check_wait']}",
        f"ETF_OPTION underlying_asset_id 空值数量: {metrics['etf_under_missing']}",
        f"contract_rule_detail AKShare/OpenCTP 残留数量: {metrics['src_resid']}",
        f"validation_summary FAIL 数量: {int((val['result']=='FAIL').sum())}",
        f"validation_summary WARN 数量: {int((val['result']=='WARN').sum())}",
        f"最终结论: {conclusion}",
        "",
        "validation_summary:",
    ]
    for _, r in val.iterrows():
        lines.append(f"- {r['result']} | {r['check_item']} | count={r['count']} | {r['description']}")
    return "\n".join(lines) + "\n"


def main() -> None:
    try:
        if FULL_OUT.exists():
            shutil.copy2(FULL_OUT, BACKUP_OUT)
        asset = step9.display_index_as_spot(step9.read_csv(step9.ASSET_PRIMARY if step9.ASSET_PRIMARY.exists() else step9.ASSET_FALLBACK, True))
        futures = step9.sync_futures(step9.read_csv(step9.FUTURES, True), step9.read_csv(step9.UNDERLYING, True), asset)
        rules = step9.merge_contract_rules()
        futures = backfill_futures_from_rules(futures, rules)
        options, _ = step9.choose_options()
        options = step9.sync_options(options, step9.read_csv(step9.UNDERLYING, True))
        etf = step9.read_csv(step9.ETF, True)
        core_etf = step9.core_etf(etf)
        index = step9.display_index_as_spot(step9.read_csv(step9.INDEX, True))
        domestic = step9.add_domestic_desc(step9.read_csv(step9.DOMESTIC, True))
        industry = step9.add_industry_desc(step9.read_csv(step9.INDUSTRY, True))
        industry_pool = step9.read_csv(step9.INDUSTRY_POOL, True)
        foreign_asset = step9.read_csv(step9.FOREIGN_ASSET, True)
        foreign = step9.add_foreign_desc(step9.read_csv(step9.FOREIGN, True))
        contract = step9.merge_contract_rules()

        pairs = build_arbitrage_pairs(futures, options, core_etf, etf, index, domestic, industry, foreign)
        wide = mapping_display_wide(pairs, futures)
        detail = detail_display(futures, rules, options)
        current = current_contract_detail(futures, rules)

        sheets_raw = {
            "README": step9.readme_df([]),
            "relation_summary": relation_summary_from_pairs(pairs),
            "arbitrage_mapping_pairs": pairs,
            "mapping_display_wide": wide,
            "detail_display": detail,
            "current_contract_detail": current,
            "contract_rule_detail": contract,
            "core_asset_master": step9.core_assets(asset, options, core_etf),
            "asset_master": asset,
            "futures_master": futures,
            "options_master": options,
            "core_etf_master": core_etf,
            "etf_master": etf,
            "index_master": index,
            "underlying_mapping": step9.read_csv(step9.UNDERLYING, True),
            "domestic_cross_market_mapping": domestic,
            "industry_chain_mapping": industry,
            "industry_chain_asset_pool": industry_pool,
            "foreign_asset_master": foreign_asset,
            "foreign_cross_market_mapping": foreign,
        }
        temp_clean = {k: clean_sheet(k, v)[0] for k, v in sheets_raw.items()}
        residual_invalid = step9.count_residual({k: v for k, v in temp_clean.items() if k != "validation_summary"}, ["CHECK", "TODO", "NEED_REVIEW"])
        val = validation(pairs, futures, clean_sheet("detail_display", detail)[0], options, domestic, industry, foreign, clean_sheet("contract_rule_detail", contract)[0], residual_invalid)
        sheets_raw["validation_summary"] = val

        order = [
            "README", "relation_summary", "validation_summary", "arbitrage_mapping_pairs", "mapping_display_wide",
            "detail_display", "current_contract_detail", "contract_rule_detail", "core_asset_master", "asset_master",
            "futures_master", "options_master", "core_etf_master", "etf_master", "index_master", "underlying_mapping",
            "domestic_cross_market_mapping", "industry_chain_mapping", "industry_chain_asset_pool",
            "foreign_asset_master", "foreign_cross_market_mapping",
        ]
        final_sheets = {name: clean_sheet(name, sheets_raw[name])[0] for name in order if name in sheets_raw}
        step9.write_excel(FULL_OUT, final_sheets)
        step9.write_excel(V2_OUT, final_sheets)

        detail_clean = final_sheets["detail_display"]
        contract_clean = final_sheets["contract_rule_detail"]
        metrics = {
            "pairs_rows": len(final_sheets["arbitrage_mapping_pairs"]),
            "pair_dup": int(final_sheets["arbitrage_mapping_pairs"]["pair_id"].duplicated().sum()),
            "multi_asset": has_multi_asset(final_sheets["arbitrage_mapping_pairs"]),
            "base_coverage": final_sheets["arbitrage_mapping_pairs"]["base_future_asset_id"].nunique(),
            "detail_rows": len(detail_clean),
            "mult_non_empty": int((detail_clean["contract_multiplier"] != "空").sum()),
            "tick_non_empty": int((detail_clean["tick_size"] != "空").sum()),
            "unit_non_empty": int((detail_clean["contract_unit"] != "空").sum()),
            "fut_ug_empty": int(futures["underlying_group"].map(invalid).sum()),
            "check_wait": int(options.astype(str).apply(lambda s: s.str.contains("CHECK_WAIT_ETF_MASTER", regex=False, na=False)).sum().sum()),
            "etf_under_missing": int(((options["subtype"] == "ETF_OPTION") & (options["underlying_asset_id"].map(invalid) | options["underlying_asset_id"].astype(str).str.strip().eq(""))).sum()),
            "src_resid": int(contract_clean.astype(str).apply(lambda s: s.str.contains("AKShare|OpenCTP", regex=True, na=False)).sum().sum()),
        }
        fail = int((final_sheets["validation_summary"]["result"] == "FAIL").sum())
        warn = int((final_sheets["validation_summary"]["result"] == "WARN").sum())
        conclusion = "NEED_MAJOR_FIX" if fail else ("NEED_MINOR_FIX" if warn else "PASS")
        REPORT_OUT.write_text(report(metrics, final_sheets["validation_summary"], conclusion), encoding="utf-8")

        print(f"mapping_info_full.xlsx 是否更新: {'Y' if FULL_OUT.exists() else 'N'}")
        print(f"mapping_info_full_v2.xlsx 是否生成: {'Y' if V2_OUT.exists() else 'N'}")
        print(f"arbitrage_mapping_pairs 行数: {metrics['pairs_rows']}")
        print(f"pair_id 重复数: {metrics['pair_dup']}")
        print(f"base_future 覆盖数量: {metrics['base_coverage']}")
        print(f"是否存在多资产塞一个格子: {'Y' if metrics['multi_asset'] else 'N'}")
        print(f"detail_display 行数: {metrics['detail_rows']}")
        print(f"detail_display contract_multiplier 非空数量: {metrics['mult_non_empty']}")
        print(f"detail_display tick_size 非空数量: {metrics['tick_non_empty']}")
        print(f"detail_display contract_unit 非空数量: {metrics['unit_non_empty']}")
        print(f"validation_summary FAIL 数量: {fail}")
        print(f"validation_summary WARN 数量: {warn}")
        print(f"最终结论: {conclusion}")
        print(f"检查报告路径: {REPORT_OUT}")
    except Exception as exc:
        print(f"v2 映射与详情表重建失败: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
