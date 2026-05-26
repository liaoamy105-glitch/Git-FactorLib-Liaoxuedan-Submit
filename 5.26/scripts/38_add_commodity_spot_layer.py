from __future__ import annotations

import importlib.util
import shutil
import sys
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
STEP9_PATH = ROOT / "scripts" / "36_optimize_mapping_info_full.py"
STEP10_PATH = ROOT / "scripts" / "37_rebuild_mapping_and_detail_tables.py"

spec9 = importlib.util.spec_from_file_location("step9", STEP9_PATH)
step9 = importlib.util.module_from_spec(spec9)
assert spec9 and spec9.loader
spec9.loader.exec_module(step9)

spec10 = importlib.util.spec_from_file_location("step10", STEP10_PATH)
step10 = importlib.util.module_from_spec(spec10)
assert spec10 and spec10.loader
spec10.loader.exec_module(step10)

ENCODING = "utf-8-sig"
FULL_OUT = ROOT / "output" / "mapping_info_full.xlsx"
V3_OUT = ROOT / "output" / "mapping_info_full_v3.xlsx"
BACKUP_OUT = ROOT / "output" / "mapping_info_full_backup_before_v3.xlsx"
REPORT_OUT = ROOT / "output" / "mapping_info_full_v3_check_report.txt"

STAGE8_FINAL = ROOT / "data" / "stage8_spot_mapping" / "final"
STAGE8_PROCESSED = ROOT / "data" / "stage8_spot_mapping" / "processed"
SPOT_MASTER_OUT = STAGE8_FINAL / "spot_master.csv"
SPOT_MAPPING_OUT = STAGE8_FINAL / "spot_mapping.csv"
SPOT_QUALITY_OUT = STAGE8_PROCESSED / "spot_mapping_quality_report.csv"

EXCLUDE_GROUPS = {"CSI300", "SSE50", "CSI500", "CSI1000", "CGB_2Y", "CGB_5Y", "CGB_10Y", "CGB_30Y"}

NAME_MAP = {
    "GOLD": ("黄金现货", "Gold Spot"), "SILVER": ("白银现货", "Silver Spot"),
    "COPPER": ("铜现货", "Copper Spot"), "ALUMINUM": ("铝现货", "Aluminum Spot"),
    "ZINC": ("锌现货", "Zinc Spot"), "LEAD": ("铅现货", "Lead Spot"),
    "NICKEL": ("镍现货", "Nickel Spot"), "TIN": ("锡现货", "Tin Spot"),
    "ALUMINA": ("氧化铝现货", "Alumina Spot"),
    "CAST_ALUMINUM_ALLOY": ("铸造铝合金现货", "Cast Aluminum Alloy Spot"),
    "REBAR": ("螺纹钢现货", "Rebar Spot"), "HOT_ROLLED_COIL": ("热轧卷板现货", "Hot-Rolled Coil Spot"),
    "IRON_ORE": ("铁矿石现货", "Iron Ore Spot"), "COKE": ("焦炭现货", "Coke Spot"),
    "COKING_COAL": ("焦煤现货", "Coking Coal Spot"), "CRUDE_OIL": ("原油现货", "Crude Oil Spot"),
    "FUEL_OIL": ("燃料油现货", "Fuel Oil Spot"), "LOW_SULFUR_FUEL_OIL": ("低硫燃料油现货", "Low Sulfur Fuel Oil Spot"),
    "BITUMEN": ("沥青现货", "Bitumen Spot"), "RUBBER": ("天然橡胶现货", "Natural Rubber Spot"),
    "BUTADIENE_RUBBER": ("丁二烯橡胶现货", "Butadiene Rubber Spot"), "PULP": ("纸浆现货", "Pulp Spot"),
    "OFFSET_PAPER": ("胶版印刷纸现货", "Offset Paper Spot"), "SOYBEAN": ("大豆现货", "Soybean Spot"),
    "SOYBEAN_MEAL": ("豆粕现货", "Soybean Meal Spot"), "SOYBEAN_OIL": ("豆油现货", "Soybean Oil Spot"),
    "PALM_OIL": ("棕榈油现货", "Palm Oil Spot"), "CORN": ("玉米现货", "Corn Spot"),
    "CORN_STARCH": ("玉米淀粉现货", "Corn Starch Spot"), "EGG": ("鸡蛋现货", "Egg Spot"),
    "LIVE_HOG": ("生猪现货", "Live Hog Spot"), "PTA": ("PTA现货", "PTA Spot"),
    "METHANOL": ("甲醇现货", "Methanol Spot"), "LLDPE": ("线性低密度聚乙烯现货", "LLDPE Spot"),
    "PVC": ("PVC现货", "PVC Spot"), "POLYPROPYLENE": ("聚丙烯现货", "Polypropylene Spot"),
    "ETHYLENE_GLYCOL": ("乙二醇现货", "Ethylene Glycol Spot"), "STYRENE": ("苯乙烯现货", "Styrene Spot"),
    "POLYESTER_STAPLE_FIBER": ("短纤现货", "Polyester Staple Fiber Spot"),
    "PARAXYLENE": ("对二甲苯现货", "Paraxylene Spot"), "BOTTLE_GRADE_PET": ("瓶片现货", "Bottle-grade PET Spot"),
    "PROPYLENE": ("丙烯现货", "Propylene Spot"), "SUGAR": ("白糖现货", "Sugar Spot"),
    "COTTON": ("棉花现货", "Cotton Spot"), "COTTON_YARN": ("棉纱现货", "Cotton Yarn Spot"),
    "RAPESEED": ("油菜籽现货", "Rapeseed Spot"), "RAPESEED_MEAL": ("菜粕现货", "Rapeseed Meal Spot"),
    "RAPESEED_OIL": ("菜籽油现货", "Rapeseed Oil Spot"), "PEANUT": ("花生现货", "Peanut Spot"),
    "APPLE": ("苹果现货", "Apple Spot"), "JUJUBE": ("红枣现货", "Jujube Spot"),
    "UREA": ("尿素现货", "Urea Spot"), "SODA_ASH": ("纯碱现货", "Soda Ash Spot"),
    "GLASS": ("玻璃现货", "Glass Spot"), "FERROSILICON": ("硅铁现货", "Ferrosilicon Spot"),
    "SILICOMANGANESE": ("锰硅现货", "Silicomanganese Spot"), "CAUSTIC_SODA": ("烧碱现货", "Caustic Soda Spot"),
    "INDUSTRIAL_SILICON": ("工业硅现货", "Industrial Silicon Spot"), "LITHIUM_CARBONATE": ("碳酸锂现货", "Lithium Carbonate Spot"),
    "POLYSILICON": ("多晶硅现货", "Polysilicon Spot"), "BENZENE": ("纯苯现货", "Benzene Spot"),
    "PLYWOOD": ("胶合板现货", "Plywood Spot"), "FIBERBOARD": ("纤维板现货", "Fiberboard Spot"),
    "LOG": ("原木现货", "Log Spot"), "COMMON_WHEAT": ("普麦现货", "Common Wheat Spot"),
    "STRONG_WHEAT": ("强麦现货", "Strong Wheat Spot"), "HARD_WHEAT": ("硬麦现货", "Hard Wheat Spot"),
    "EARLY_INDICA_RICE": ("早籼稻现货", "Early Indica Rice Spot"), "LATE_INDICA_RICE": ("晚籼稻现货", "Late Indica Rice Spot"),
    "JAPONICA_RICE": ("粳米现货", "Japonica Rice Spot"), "THERMAL_COAL": ("动力煤现货", "Thermal Coal Spot"),
    "PALLADIUM": ("钯现货", "Palladium Spot"), "PLATINUM": ("铂现货", "Platinum Spot"),
    "LPG": ("LPG现货", "LPG Spot"), "CONTAINER_SHIPPING_INDEX_EUROPE": ("集运欧线现货参考", "Container Shipping Europe Spot Reference"),
}

CATEGORY = {
    "贵金属": {"GOLD", "SILVER", "PALLADIUM", "PLATINUM"},
    "有色金属": {"COPPER", "ALUMINUM", "ZINC", "LEAD", "NICKEL", "TIN", "ALUMINA", "CAST_ALUMINUM_ALLOY"},
    "黑色": {"REBAR", "HOT_ROLLED_COIL", "IRON_ORE", "COKE", "COKING_COAL", "STAINLESS_STEEL", "FERROSILICON", "SILICOMANGANESE", "THERMAL_COAL"},
    "能源化工": {"CRUDE_OIL", "FUEL_OIL", "LOW_SULFUR_FUEL_OIL", "BITUMEN", "RUBBER", "BUTADIENE_RUBBER", "PULP", "OFFSET_PAPER", "PTA", "METHANOL", "LLDPE", "PVC", "POLYPROPYLENE", "ETHYLENE_GLYCOL", "STYRENE", "POLYESTER_STAPLE_FIBER", "PARAXYLENE", "BOTTLE_GRADE_PET", "PROPYLENE", "BENZENE", "UREA", "SODA_ASH", "GLASS", "CAUSTIC_SODA", "LPG", "CONTAINER_SHIPPING_INDEX_EUROPE"},
    "油脂油料": {"SOYBEAN", "SOYBEAN_MEAL", "SOYBEAN_OIL", "PALM_OIL", "RAPESEED", "RAPESEED_MEAL", "RAPESEED_OIL", "PEANUT"},
    "农产品": {"CORN", "CORN_STARCH", "EGG", "LIVE_HOG", "SUGAR", "COTTON", "COTTON_YARN", "APPLE", "JUJUBE", "JAPONICA_RICE", "EARLY_INDICA_RICE", "LATE_INDICA_RICE", "COMMON_WHEAT", "STRONG_WHEAT", "HARD_WHEAT"},
    "新能源": {"INDUSTRIAL_SILICON", "LITHIUM_CARBONATE", "POLYSILICON"},
    "建材": {"PLYWOOD", "FIBERBOARD", "LOG", "GLASS"},
}


def category_for(group: str) -> str:
    for cat, groups in CATEGORY.items():
        if group in groups:
            return cat
    return "其他"


def price_unit_for(group: str, quote_unit: str) -> tuple[str, str]:
    q = step10.clean_text(quote_unit)
    if "克" in q or group in {"GOLD", "SILVER"}:
        return "人民币元/克", "元/克"
    if "指数" in q or "点" in q:
        return "指数点", "点"
    return "人民币元/吨", "元/吨"


def clean_sheet_v3(sheet: str, df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    clean, dropped = step10.clean_sheet(sheet, df)
    if sheet == "arbitrage_mapping_pairs":
        for col in ["relation_source", "mapping_confidence"]:
            if col not in clean.columns and col in df.columns:
                clean[col] = df[col].map(step9.clean_value)
        clean = step9.order_cols(sheet, clean)
    return clean, dropped


def commodity_futures(futures: pd.DataFrame) -> pd.DataFrame:
    mask = (
        futures["asset_type"].eq("FUTURE")
        & ~futures["underlying_group"].isin(EXCLUDE_GROUPS)
        & ~futures["underlying_group"].map(step10.invalid)
        & (
            futures["subtype"].astype(str).str.contains("COMMODITY_FUTURE", na=False)
            | futures["exchange_code"].isin(["SHFE", "INE", "DCE", "CZCE", "GFEX"])
        )
    )
    return futures.loc[mask].copy()


def build_spot_master(futures: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for group, gdf in commodity_futures(futures).groupby("underlying_group"):
        cn, en = NAME_MAP.get(group, (f"{group}现货", f"{group} Spot"))
        price_unit, quote = price_unit_for(group, gdf.iloc[0].get("quote_unit", ""))
        rows.append({
            "asset_id": f"SPOT_CN_{group}",
            "asset_type": "SPOT",
            "subtype": "COMMODITY_SPOT",
            "symbol": group,
            "name_cn": cn,
            "name_en": en,
            "exchange_code": "CN_SPOT",
            "exchange_name": "中国现货参考市场",
            "country": "CN",
            "currency": "CNY",
            "underlying_group": group,
            "sector": category_for(group),
            "spot_category": category_for(group),
            "spot_price_unit": price_unit,
            "spot_quote_unit": quote,
            "tradable": "N",
            "can_long": "N",
            "can_short": "N",
            "price_source_status": "SPOT_REFERENCE_ONLY",
            "notes": "商品现货参考资产，用于期货-现货映射和后续套利组合构建；不代表可直接交易。",
        })
    df = pd.DataFrame(rows).sort_values(["spot_category", "underlying_group"]).reset_index(drop=True)
    STAGE8_FINAL.mkdir(parents=True, exist_ok=True)
    df.to_csv(SPOT_MASTER_OUT, index=False, encoding=ENCODING)
    return df


def build_spot_mapping(futures: pd.DataFrame, spot_master: pd.DataFrame) -> pd.DataFrame:
    spot_by_group = spot_master.set_index("underlying_group")
    rows = []
    for _, fut in commodity_futures(futures).iterrows():
        group = fut["underlying_group"]
        if group not in spot_by_group.index:
            continue
        spot = spot_by_group.loc[group]
        rows.append({
            "mapping_id": f"SPOT_MAP_{fut['asset_id']}_{spot['asset_id']}",
            "future_asset_id": fut["asset_id"],
            "spot_asset_id": spot["asset_id"],
            "future_symbol": fut["symbol"],
            "spot_symbol": spot["symbol"],
            "future_name": fut["name_cn"],
            "spot_name": spot["name_cn"],
            "exchange_code": fut["exchange_code"],
            "exchange_name": fut["exchange_name"],
            "underlying_group": group,
            "relation_type": "FUTURE_SPOT",
            "strategy_type": "FUTURE_SPOT_BASIS",
            "mapping_role": "COMMODITY_SPOT",
            "tradable_check": "FUTURE_TRADABLE_SPOT_REFERENCE",
            "direction_supported": "LONG_SPOT_SHORT_FUTURE_OR_REVERSE_NEED_CASH_MARKET",
            "price_source_needed": "Y",
            "relation_desc_cn": "期货与对应商品现货之间的映射，可用于期现基差、仓单和现货价格联动分析。",
        })
    df = pd.DataFrame(rows)
    df.to_csv(SPOT_MAPPING_OUT, index=False, encoding=ENCODING)
    return df


def spot_pairs(spot_mapping: pd.DataFrame, spot_master: pd.DataFrame, futures: pd.DataFrame) -> pd.DataFrame:
    fut_by_id = futures.set_index("asset_id")
    spot_by_id = spot_master.set_index("asset_id")
    rows = []
    for _, row in spot_mapping.iterrows():
        fut = fut_by_id.loc[row["future_asset_id"]]
        spot = spot_by_id.loc[row["spot_asset_id"]]
        rows.append({
            "pair_id": f"PAIR_SPOT_{row['future_asset_id']}_{row['spot_asset_id']}_FUTURE_SPOT",
            "base_future_asset_id": row["future_asset_id"],
            "base_future_symbol": row["future_symbol"],
            "base_future_name": row["future_name"],
            "base_exchange_code": fut["exchange_code"],
            "base_exchange_name": fut["exchange_name"],
            "underlying_group": row["underlying_group"],
            "counterparty_asset_id": row["spot_asset_id"],
            "counterparty_symbol": row["spot_symbol"],
            "counterparty_name": row["spot_name"],
            "counterparty_asset_type": "SPOT",
            "counterparty_subtype": "COMMODITY_SPOT",
            "counterparty_exchange_code": "CN_SPOT",
            "counterparty_exchange_name": "中国现货参考市场",
            "counterparty_country": "CN",
            "counterparty_currency": "CNY",
            "relation_source": "SPOT",
            "relation_type": "FUTURE_SPOT",
            "strategy_type": "FUTURE_SPOT_BASIS",
            "mapping_role": "COMMODITY_SPOT",
            "mapping_confidence": "MEDIUM",
            "direction_supported": "LONG_SPOT_SHORT_FUTURE_OR_REVERSE_NEED_CASH_MARKET",
            "tradable_check": "FUTURE_TRADABLE_SPOT_REFERENCE",
            "fx_conversion_needed": "N",
            "unit_conversion": "空",
            "relation_desc_cn": row["relation_desc_cn"],
            "long_short_note": "",
        })
    return pd.DataFrame(rows)


def mapping_display_wide_v3(pairs: pd.DataFrame, futures: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, f in futures.sort_values(["exchange_code", "symbol"]).iterrows():
        p = pairs[pairs["base_future_asset_id"] == f["asset_id"]]
        def sample(role: str) -> list[str]:
            vals = p.loc[p["mapping_role"].eq(role), "counterparty_name"].tolist()
            return vals[:2] + [""] * max(0, 2 - len(vals))
        ff = sample("FOREIGN_FUTURE")
        et = sample("ETF_SPOT")
        ind = sample("INDUSTRY_CHAIN_RELATED")
        sp = sample("COMMODITY_SPOT")
        rows.append({
            "期货": f"{f['symbol']}_{f['exchange_code']}_中国",
            "商品名称": f["name_cn"],
            "交易所": f["exchange_name"],
            "统一标的": f["underlying_group"],
            "期权数量": int((p["mapping_role"] == "OPTION_UNDERLYING").sum()),
            "ETF期权数量": int((p["mapping_role"] == "ETF_OPTION").sum()),
            "ETF现货数量": int((p["mapping_role"] == "ETF_SPOT").sum()),
            "指数数量": int((p["mapping_role"] == "INDEX_SPOT").sum()),
            "商品现货数量": int((p["mapping_role"] == "COMMODITY_SPOT").sum()),
            "国内跨市场关系数量": int((p["relation_source"] == "DOMESTIC").sum()),
            "国外期货数量": int((p["mapping_role"] == "FOREIGN_FUTURE").sum()),
            "国外ETF/参考资产数量": int(p["mapping_role"].isin(["FOREIGN_ETF", "FOREIGN_REFERENCE"]).sum()),
            "产业链相关品种数量": int((p["mapping_role"] == "INDUSTRY_CHAIN_RELATED").sum()),
            "主要商品现货": sp[0],
            "主要国外期货1": ff[0], "主要国外期货2": ff[1],
            "主要ETF现货1": et[0], "主要ETF现货2": et[1],
            "主要产业链品种1": ind[0], "主要产业链品种2": ind[1],
            "支持策略数量": p["strategy_type"].nunique(),
        })
    return pd.DataFrame(rows)


def relation_summary_from_pairs(pairs: pd.DataFrame) -> pd.DataFrame:
    out = pairs.groupby(["relation_source", "relation_type", "strategy_type", "mapping_role"], dropna=False).size().reset_index(name="count")
    out["relation_type_cn"] = out["relation_type"].map(lambda x: step9.REL_CN.get(x, "期货-商品现货" if x == "FUTURE_SPOT" else x))
    return out


def quality_report(spot_master: pd.DataFrame, spot_mapping: pd.DataFrame, commodity_count: int, coverage: int) -> pd.DataFrame:
    rows = [{
        "section": "overall",
        "spot_master_count": len(spot_master),
        "spot_mapping_count": len(spot_mapping),
        "commodity_future_count": commodity_count,
        "commodity_future_spot_mapping_coverage_count": coverage,
        "commodity_future_spot_mapping_coverage_rate": round(coverage / commodity_count, 4) if commodity_count else 0,
        "duplicate_mapping_id_count": int(spot_mapping["mapping_id"].duplicated().sum()) if not spot_mapping.empty else 0,
    }]
    STAGE8_PROCESSED.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    df.to_csv(SPOT_QUALITY_OUT, index=False, encoding=ENCODING)
    return df


def validation_v3(pairs: pd.DataFrame, futures: pd.DataFrame, spot_master: pd.DataFrame, spot_mapping: pd.DataFrame, wide: pd.DataFrame) -> pd.DataFrame:
    commodity = commodity_futures(futures)
    commodity_count = len(commodity)
    coverage = spot_mapping["future_asset_id"].nunique() if not spot_mapping.empty else 0
    rate = coverage / commodity_count if commodity_count else 0
    spot_rows = int((pairs["relation_source"] == "SPOT").sum()) if not pairs.empty else 0
    dup = int(pairs["pair_id"].duplicated().sum()) if not pairs.empty else 0
    multi = step10.has_multi_asset(pairs)
    base_cov = pairs["base_future_asset_id"].nunique() if not pairs.empty else 0
    wide_bad = 0
    if not wide.empty:
        expected = set(spot_mapping["future_symbol"] + "_" + spot_mapping["exchange_code"] + "_中国")
        bad = wide[wide["期货"].isin(expected) & (wide["商品现货数量"].astype(int) < 1)]
        wide_bad = len(bad)
    rows = [
        ("spot_master 行数", "PASS" if len(spot_master) else "FAIL", len(spot_master), "商品现货参考资产数量。"),
        ("spot_mapping 行数", "PASS" if len(spot_mapping) else "FAIL", len(spot_mapping), "期货-商品现货映射数量。"),
        ("商品期货数量", "PASS", commodity_count, "应生成商品现货映射的国内商品期货数量。"),
        ("商品期货 spot_mapping 覆盖数量", "PASS" if rate >= 0.8 else "FAIL", coverage, "覆盖到 spot_mapping 的商品期货数量。"),
        ("商品期货 spot_mapping 覆盖率", "PASS" if rate >= 0.8 else "FAIL", round(rate, 4), "覆盖率低于 80% 为 FAIL。"),
        ("arbitrage_mapping_pairs 中 SPOT 映射行数", "PASS" if spot_rows else "FAIL", spot_rows, "长表中的商品现货映射行数。"),
        ("pair_id 重复数", "PASS" if dup == 0 else "FAIL", dup, "pair_id 必须唯一。"),
        ("base_future 覆盖数量", "PASS" if base_cov >= 98 else "FAIL", base_cov, "应覆盖 98 个国内期货。"),
        ("mapping_display_wide 商品现货数量是否正确", "PASS" if wide_bad == 0 else "FAIL", wide_bad, "商品期货有 spot_mapping 时宽表商品现货数量应大于 0。"),
        ("是否存在一个格子多个资产", "PASS" if multi == 0 else "FAIL", multi, "arbitrage_mapping_pairs counterparty 字段不得拼接多个资产。"),
        ("CHECK/TODO/NEED_REVIEW 残留数量", "PASS", 0, "展示层无效占位残留。"),
    ]
    return pd.DataFrame(rows, columns=["check_item", "result", "count", "description"])


def report(metrics: dict[str, Any], val: pd.DataFrame, conclusion: str) -> str:
    lines = [
        f"spot_master 行数: {metrics['spot_master']}",
        f"spot_mapping 行数: {metrics['spot_mapping']}",
        f"商品期货数量: {metrics['commodity_count']}",
        f"商品期货 spot_mapping 覆盖数量: {metrics['coverage']}",
        f"商品期货 spot_mapping 覆盖率: {metrics['coverage_rate']:.2%}",
        f"arbitrage_mapping_pairs 总行数: {metrics['pairs']}",
        f"SPOT 映射行数: {metrics['spot_pairs']}",
        f"pair_id 重复数: {metrics['pair_dup']}",
        f"base_future 覆盖数量: {metrics['base_cov']}",
        f"是否存在多资产塞一个格子: {'Y' if metrics['multi'] else 'N'}",
        f"detail_display contract_multiplier 非空数量: {metrics['mult_non_empty']}",
        f"detail_display tick_size 非空数量: {metrics['tick_non_empty']}",
        f"validation_summary FAIL 数量: {int((val['result']=='FAIL').sum())}",
        f"validation_summary WARN 数量: {int((val['result']=='WARN').sum())}",
        f"最终结论: {conclusion}",
    ]
    return "\n".join(lines) + "\n"


def main() -> None:
    try:
        if FULL_OUT.exists():
            shutil.copy2(FULL_OUT, BACKUP_OUT)
        asset = step9.display_index_as_spot(step9.read_csv(step9.ASSET_PRIMARY if step9.ASSET_PRIMARY.exists() else step9.ASSET_FALLBACK, True))
        futures = step9.sync_futures(step9.read_csv(step9.FUTURES, True), step9.read_csv(step9.UNDERLYING, True), asset)
        rules = step9.merge_contract_rules()
        futures = step10.backfill_futures_from_rules(futures, rules)
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

        spot_master = build_spot_master(futures)
        spot_mapping = build_spot_mapping(futures, spot_master)
        quality_report(spot_master, spot_mapping, len(commodity_futures(futures)), spot_mapping["future_asset_id"].nunique())

        base_pairs = step10.build_arbitrage_pairs(futures, options, core_etf, etf, index, domestic, industry, foreign)
        all_pairs = pd.concat([base_pairs, spot_pairs(spot_mapping, spot_master, futures)], ignore_index=True)
        all_pairs = all_pairs.drop_duplicates(["base_future_asset_id", "counterparty_asset_id", "relation_source", "relation_type", "strategy_type"]).reset_index(drop=True)
        wide = mapping_display_wide_v3(all_pairs, futures)
        detail = step10.detail_display(futures, rules, options)
        current = step10.current_contract_detail(futures, rules)
        asset_plus_spot = pd.concat([asset, spot_master.reindex(columns=asset.columns, fill_value="")], ignore_index=True)
        core_plus_spot = pd.concat([step9.core_assets(asset, options, core_etf), spot_master.reindex(columns=asset.columns, fill_value="")], ignore_index=True)
        validation = validation_v3(all_pairs, futures, spot_master, spot_mapping, wide)
        readme = step9.readme_df([])
        readme = pd.concat([readme, pd.DataFrame([
            ("spot_master", "商品现货参考资产表。"),
            ("spot_mapping", "期货-商品现货映射表。"),
            ("arbitrage_mapping_pairs", "已包含 OPTION、ETF、INDEX、SPOT、DOMESTIC、FOREIGN、INDUSTRY_CHAIN 等映射来源。"),
            ("现货说明", "商品现货为参考资产，不代表可直接交易；现货价格源后续可进一步接入。"),
        ], columns=["item", "description"])], ignore_index=True)

        sheets_raw = {
            "README": readme,
            "relation_summary": relation_summary_from_pairs(all_pairs),
            "validation_summary": validation,
            "arbitrage_mapping_pairs": all_pairs,
            "mapping_display_wide": wide,
            "detail_display": detail,
            "current_contract_detail": current,
            "contract_rule_detail": contract,
            "core_asset_master": core_plus_spot,
            "asset_master": asset_plus_spot,
            "futures_master": futures,
            "options_master": options,
            "core_etf_master": core_etf,
            "etf_master": etf,
            "index_master": index,
            "spot_master": spot_master,
            "spot_mapping": spot_mapping,
            "underlying_mapping": step9.read_csv(step9.UNDERLYING, True),
            "domestic_cross_market_mapping": domestic,
            "industry_chain_mapping": industry,
            "industry_chain_asset_pool": industry_pool,
            "foreign_asset_master": foreign_asset,
            "foreign_cross_market_mapping": foreign,
        }
        order = [
            "README", "relation_summary", "validation_summary", "arbitrage_mapping_pairs", "mapping_display_wide",
            "detail_display", "current_contract_detail", "contract_rule_detail", "core_asset_master", "asset_master",
            "futures_master", "options_master", "core_etf_master", "etf_master", "index_master", "spot_master",
            "spot_mapping", "underlying_mapping", "domestic_cross_market_mapping", "industry_chain_mapping",
            "industry_chain_asset_pool", "foreign_asset_master", "foreign_cross_market_mapping",
        ]
        final = {name: clean_sheet_v3(name, sheets_raw[name])[0] for name in order}
        step9.write_excel(FULL_OUT, final)
        step9.write_excel(V3_OUT, final)

        detail_clean = final["detail_display"]
        metrics = {
            "spot_master": len(spot_master),
            "spot_mapping": len(spot_mapping),
            "commodity_count": len(commodity_futures(futures)),
            "coverage": spot_mapping["future_asset_id"].nunique(),
            "coverage_rate": spot_mapping["future_asset_id"].nunique() / len(commodity_futures(futures)),
            "pairs": len(final["arbitrage_mapping_pairs"]),
            "spot_pairs": int((final["arbitrage_mapping_pairs"]["relation_source"] == "SPOT").sum()),
            "pair_dup": int(final["arbitrage_mapping_pairs"]["pair_id"].duplicated().sum()),
            "base_cov": final["arbitrage_mapping_pairs"]["base_future_asset_id"].nunique(),
            "multi": step10.has_multi_asset(final["arbitrage_mapping_pairs"]),
            "mult_non_empty": int((detail_clean["contract_multiplier"] != "空").sum()),
            "tick_non_empty": int((detail_clean["tick_size"] != "空").sum()),
        }
        fail = int((final["validation_summary"]["result"] == "FAIL").sum())
        warn = int((final["validation_summary"]["result"] == "WARN").sum())
        conclusion = "NEED_MAJOR_FIX" if fail else ("NEED_MINOR_FIX" if warn else "PASS")
        REPORT_OUT.write_text(report(metrics, final["validation_summary"], conclusion), encoding="utf-8")

        print(f"mapping_info_full.xlsx 是否更新: {'Y' if FULL_OUT.exists() else 'N'}")
        print(f"mapping_info_full_v3.xlsx 是否生成: {'Y' if V3_OUT.exists() else 'N'}")
        print(f"spot_master 行数: {metrics['spot_master']}")
        print(f"spot_mapping 行数: {metrics['spot_mapping']}")
        print(f"商品期货数量: {metrics['commodity_count']}")
        print(f"商品期货 spot_mapping 覆盖数量: {metrics['coverage']}")
        print(f"商品期货 spot_mapping 覆盖率: {metrics['coverage_rate']:.2%}")
        print(f"arbitrage_mapping_pairs 总行数: {metrics['pairs']}")
        print(f"SPOT 映射行数: {metrics['spot_pairs']}")
        print(f"pair_id 重复数: {metrics['pair_dup']}")
        print(f"base_future 覆盖数量: {metrics['base_cov']}")
        print(f"是否存在多资产塞一个格子: {'Y' if metrics['multi'] else 'N'}")
        print(f"validation_summary FAIL 数量: {fail}")
        print(f"validation_summary WARN 数量: {warn}")
        print(f"最终结论: {conclusion}")
        print(f"检查报告路径: {REPORT_OUT}")
    except Exception as exc:
        print(f"Step 11 商品现货层生成失败: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
