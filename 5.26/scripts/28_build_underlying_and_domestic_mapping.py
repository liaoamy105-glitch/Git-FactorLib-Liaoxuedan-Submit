from __future__ import annotations

import itertools
import re
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
FINAL_DIR = ROOT / "data" / "stage5_mapping" / "final"
PROCESSED_DIR = ROOT / "data" / "stage5_mapping" / "processed"
OUTPUT_DIR = ROOT / "output" / "stage5_mapping"
EXCEL_PATH = ROOT / "output" / "mapping_info_stage5_mapping_draft.xlsx"
ENCODING = "utf-8-sig"

ASSET_MASTER_PRIMARY = ROOT / "data" / "stage4_etf_index" / "final" / "asset_master_with_etf_index_cleaned.csv"
ASSET_MASTER_FALLBACK = ROOT / "data" / "stage4_etf_index" / "final" / "asset_master_with_etf_index.csv"
FUTURES_MASTER = ROOT / "data" / "final" / "futures_master.csv"
OPTIONS_MASTER = ROOT / "data" / "stage2_options" / "final" / "options_master_tushare_primary_final.csv"
ETF_MASTER = ROOT / "data" / "stage4_etf_index" / "final" / "etf_master_cleaned.csv"
INDEX_MASTER = ROOT / "data" / "stage4_etf_index" / "final" / "index_master.csv"
OPTIONS_WITH_ETF = ROOT / "data" / "stage4_etf_index" / "final" / "options_master_with_etf_underlying.csv"
ETF_INDEX_MAPPING = ROOT / "data" / "stage4_etf_index" / "final" / "etf_index_mapping.csv"

GROUP_MAP = {
    "IF": "CSI300", "IO": "CSI300", "510300": "CSI300", "159919": "CSI300", "000300": "CSI300",
    "IH": "SSE50", "HO": "SSE50", "510050": "SSE50", "000016": "SSE50",
    "IC": "CSI500", "510500": "CSI500", "159922": "CSI500", "000905": "CSI500",
    "IM": "CSI1000", "MO": "CSI1000", "159845": "CSI1000", "159629": "CSI1000", "000852": "CSI1000",
    "588000": "STAR50", "588080": "STAR50", "588050": "STAR50", "000688": "STAR50",
    "159915": "CHINEXT", "399006": "CHINEXT",
    "159949": "CHINEXT50", "399673": "CHINEXT50",
    "159901": "SZSE100", "399330": "SZSE100",
    "510880": "SSE_DIVIDEND", "000015": "SSE_DIVIDEND",
    "AU": "GOLD", "AG": "SILVER", "CU": "COPPER", "BC": "COPPER", "AL": "ALUMINUM",
    "ZN": "ZINC", "PB": "LEAD", "NI": "NICKEL", "SN": "TIN", "AO": "ALUMINA",
    "AD": "CAST_ALUMINUM_ALLOY", "RB": "REBAR", "HC": "HOT_ROLLED_COIL", "SS": "STAINLESS_STEEL",
    "WR": "WIRE_ROD", "I": "IRON_ORE", "J": "COKE", "JM": "COKING_COAL",
    "SC": "CRUDE_OIL", "SCTAS": "CRUDE_OIL", "FU": "FUEL_OIL", "LU": "LOW_SULFUR_FUEL_OIL",
    "BU": "BITUMEN", "RU": "RUBBER", "NR": "RUBBER", "BR": "BUTADIENE_RUBBER", "SP": "PULP",
    "OP": "OFFSET_PAPER", "M": "SOYBEAN_MEAL", "Y": "SOYBEAN_OIL", "P": "PALM_OIL",
    "A": "SOYBEAN", "B": "SOYBEAN", "C": "CORN", "CS": "CORN_STARCH", "JD": "EGG",
    "LH": "LIVE_HOG", "RR": "JAPONICA_RICE", "TA": "PTA", "MA": "METHANOL", "ME": "METHANOL",
    "L": "LLDPE", "V": "PVC", "PP": "POLYPROPYLENE", "EG": "ETHYLENE_GLYCOL",
    "EB": "STYRENE", "PF": "POLYESTER_STAPLE_FIBER", "PX": "PARAXYLENE",
    "PR": "BOTTLE_GRADE_PET", "PL": "PROPYLENE", "SR": "SUGAR", "CF": "COTTON",
    "CY": "COTTON_YARN", "OI": "RAPESEED_OIL", "RO": "RAPESEED_OIL", "RM": "RAPESEED_MEAL",
    "RS": "RAPESEED", "PK": "PEANUT", "AP": "APPLE", "CJ": "JUJUBE", "UR": "UREA",
    "SA": "SODA_ASH", "FG": "GLASS", "SF": "FERROSILICON", "SM": "SILICOMANGANESE",
    "SH": "CAUSTIC_SODA", "SI": "INDUSTRIAL_SILICON", "LC": "LITHIUM_CARBONATE",
    "PS": "POLYSILICON",
}

GROUP_NAME = {
    "CSI300": ("沪深300", "CSI 300"),
    "SSE50": ("上证50", "SSE 50"),
    "CSI500": ("中证500", "CSI 500"),
    "CSI1000": ("中证1000", "CSI 1000"),
    "STAR50": ("科创50", "STAR 50"),
    "CHINEXT": ("创业板指", "ChiNext Index"),
    "CHINEXT50": ("创业板50", "ChiNext 50"),
    "SZSE100": ("深证100", "SZSE 100"),
    "SSE_DIVIDEND": ("上证红利", "SSE Dividend"),
    "GOLD": ("黄金", "Gold"),
    "SILVER": ("白银", "Silver"),
    "COPPER": ("铜", "Copper"),
    "ALUMINUM": ("铝", "Aluminum"),
    "ZINC": ("锌", "Zinc"),
    "LEAD": ("铅", "Lead"),
    "NICKEL": ("镍", "Nickel"),
    "TIN": ("锡", "Tin"),
    "ALUMINA": ("氧化铝", "Alumina"),
    "CAST_ALUMINUM_ALLOY": ("铸造铝合金", "Cast Aluminum Alloy"),
    "REBAR": ("螺纹钢", "Rebar"),
    "HOT_ROLLED_COIL": ("热轧卷板", "Hot Rolled Coil"),
    "STAINLESS_STEEL": ("不锈钢", "Stainless Steel"),
    "WIRE_ROD": ("线材", "Wire Rod"),
    "IRON_ORE": ("铁矿石", "Iron Ore"),
    "COKE": ("焦炭", "Coke"),
    "COKING_COAL": ("焦煤", "Coking Coal"),
    "CRUDE_OIL": ("原油", "Crude Oil"),
    "FUEL_OIL": ("燃料油", "Fuel Oil"),
    "LOW_SULFUR_FUEL_OIL": ("低硫燃料油", "Low Sulfur Fuel Oil"),
    "BITUMEN": ("沥青", "Bitumen"),
    "RUBBER": ("橡胶", "Rubber"),
    "BUTADIENE_RUBBER": ("丁二烯橡胶", "Butadiene Rubber"),
    "PULP": ("纸浆", "Pulp"),
    "OFFSET_PAPER": ("双胶纸", "Offset Paper"),
    "SOYBEAN_MEAL": ("豆粕", "Soybean Meal"),
    "SOYBEAN_OIL": ("豆油", "Soybean Oil"),
    "PALM_OIL": ("棕榈油", "Palm Oil"),
    "SOYBEAN": ("大豆", "Soybean"),
    "CORN": ("玉米", "Corn"),
    "CORN_STARCH": ("玉米淀粉", "Corn Starch"),
    "EGG": ("鸡蛋", "Egg"),
    "LIVE_HOG": ("生猪", "Live Hog"),
    "JAPONICA_RICE": ("粳米", "Japonica Rice"),
    "PTA": ("PTA", "PTA"),
    "METHANOL": ("甲醇", "Methanol"),
    "LLDPE": ("线型低密度聚乙烯", "LLDPE"),
    "PVC": ("PVC", "PVC"),
    "POLYPROPYLENE": ("聚丙烯", "Polypropylene"),
    "ETHYLENE_GLYCOL": ("乙二醇", "Ethylene Glycol"),
    "STYRENE": ("苯乙烯", "Styrene"),
    "POLYESTER_STAPLE_FIBER": ("短纤", "Polyester Staple Fiber"),
    "PARAXYLENE": ("对二甲苯", "Paraxylene"),
    "BOTTLE_GRADE_PET": ("瓶片", "Bottle-grade PET"),
    "PROPYLENE": ("丙烯", "Propylene"),
    "SUGAR": ("白糖", "Sugar"),
    "COTTON": ("棉花", "Cotton"),
    "COTTON_YARN": ("棉纱", "Cotton Yarn"),
    "RAPESEED_OIL": ("菜籽油", "Rapeseed Oil"),
    "RAPESEED_MEAL": ("菜粕", "Rapeseed Meal"),
    "RAPESEED": ("油菜籽", "Rapeseed"),
    "PEANUT": ("花生", "Peanut"),
    "APPLE": ("苹果", "Apple"),
    "JUJUBE": ("红枣", "Jujube"),
    "UREA": ("尿素", "Urea"),
    "SODA_ASH": ("纯碱", "Soda Ash"),
    "GLASS": ("玻璃", "Glass"),
    "FERROSILICON": ("硅铁", "Ferrosilicon"),
    "SILICOMANGANESE": ("锰硅", "Silicomanganese"),
    "CAUSTIC_SODA": ("烧碱", "Caustic Soda"),
    "INDUSTRIAL_SILICON": ("工业硅", "Industrial Silicon"),
    "LITHIUM_CARBONATE": ("碳酸锂", "Lithium Carbonate"),
    "POLYSILICON": ("多晶硅", "Polysilicon"),
}


def ensure_dirs() -> None:
    for path in [FINAL_DIR, PROCESSED_DIR, OUTPUT_DIR, EXCEL_PATH.parent]:
        path.mkdir(parents=True, exist_ok=True)


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def is_missing(value: Any) -> bool:
    text = clean_text(value)
    return text == "" or text.upper() in {"CHECK", "TODO", "NA", "NAN", "NONE", "CHECK_WAIT_ETF_MASTER"}


def read_csv_optional(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_csv(path, dtype=str).fillna("")
    print(f"提示: 输入文件不存在，已跳过: {path}")
    return pd.DataFrame()


def save_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding=ENCODING)


def choose_asset_master() -> tuple[pd.DataFrame, str]:
    if ASSET_MASTER_PRIMARY.exists():
        return pd.read_csv(ASSET_MASTER_PRIMARY, dtype=str).fillna(""), str(ASSET_MASTER_PRIMARY)
    if ASSET_MASTER_FALLBACK.exists():
        return pd.read_csv(ASSET_MASTER_FALLBACK, dtype=str).fillna(""), str(ASSET_MASTER_FALLBACK)
    raise FileNotFoundError("未找到 Stage 4 asset_master 输入文件")


def infer_group(row: pd.Series) -> tuple[str, str, str]:
    existing = clean_text(row.get("underlying_group"))
    if not is_missing(existing):
        return existing, "HIGH", "EXISTING_UNDERLYING_GROUP"
    candidates = [
        clean_text(row.get("symbol")).upper(),
        clean_text(row.get("asset_id")).split("_")[-1].upper(),
    ]
    for value in candidates:
        if value in GROUP_MAP:
            return GROUP_MAP[value], "HIGH", "CONSERVATIVE_SYMBOL_RULE"
        match = re.search(r"(?<!\d)(\d{6}|\d{6})(?!\d)", value)
        if match and match.group(1) in GROUP_MAP:
            return GROUP_MAP[match.group(1)], "HIGH", "CONSERVATIVE_SYMBOL_RULE"
    name = clean_text(row.get("name_cn"))
    name_rules = [("沪深300", "CSI300"), ("上证50", "SSE50"), ("中证500", "CSI500"), ("中证1000", "CSI1000"), ("科创50", "STAR50"), ("创业板", "CHINEXT")]
    for keyword, group in name_rules:
        if keyword in name:
            return group, "HIGH", "CONSERVATIVE_NAME_RULE"
    return "CHECK", "CHECK", "NEED_MANUAL_REVIEW"


def build_underlying_mapping(asset_master: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in asset_master.iterrows():
        group, confidence, method = infer_group(row)
        cn, en = GROUP_NAME.get(group, ("CHECK", "CHECK"))
        rows.append(
            {
                "asset_id": row.get("asset_id", ""),
                "symbol": row.get("symbol", ""),
                "name_cn": row.get("name_cn", ""),
                "asset_type": row.get("asset_type", ""),
                "subtype": row.get("subtype", ""),
                "exchange_code": row.get("exchange_code", ""),
                "underlying_group": group,
                "underlying_name_cn": cn,
                "underlying_name_en": en,
                "mapping_confidence": confidence,
                "mapping_method": method,
                "data_source_1": row.get("data_source_1", ""),
                "data_source_2": row.get("data_source_2", ""),
                "source_status": "UNDERLYING_MAPPING_DRAFT_NEED_REVIEW" if confidence != "CHECK" else "UNDERLYING_MAPPING_NEED_REVIEW",
                "notes": "统一标的映射由现有 asset_master underlying_group 和保守规则生成。",
            }
        )
    mapping = pd.DataFrame(rows)
    save_csv(mapping, FINAL_DIR / "underlying_mapping.csv")
    return mapping


def norm_asset(row: pd.Series) -> dict[str, str]:
    return {k: clean_text(row.get(k, "")) for k in ["asset_id", "symbol", "name_cn", "asset_type", "subtype", "exchange_code", "underlying_group", "tradable", "can_long", "can_short"]}


def relation_for(a: dict[str, str], b: dict[str, str], options_by_id: dict[str, pd.Series]) -> tuple[str, str, str] | None:
    types = {a["asset_type"], b["asset_type"]}
    subtypes = {a["subtype"], b["subtype"]}
    if types == {"FUTURE", "ETF"}:
        return "FUTURE_ETF", "BASIS_ARBITRAGE", "HIGH" if a["underlying_group"] == b["underlying_group"] else "MEDIUM"
    if types == {"FUTURE", "INDEX"}:
        return "FUTURE_INDEX", "BASIS_ARBITRAGE", "HIGH"
    if types == {"ETF", "INDEX"}:
        return "ETF_INDEX", "ETF_PREMIUM_DISCOUNT_ARBITRAGE", "HIGH"
    if "INDEX_OPTION" in subtypes and "INDEX" in types:
        return "OPTION_INDEX", "OPTION_ARBITRAGE", "HIGH"
    if "INDEX_OPTION" in subtypes and "FUTURE" in types:
        return "OPTION_FUTURE", "OPTION_FUTURE_ARBITRAGE", "HIGH"
    if "ETF_OPTION" in subtypes and "ETF" in types:
        option = a if a["subtype"] == "ETF_OPTION" else b
        other = b if a["subtype"] == "ETF_OPTION" else a
        opt_row = options_by_id.get(option["asset_id"])
        if opt_row is not None and clean_text(opt_row.get("underlying_asset_id")) == other["asset_id"]:
            return "OPTION_UNDERLYING", "OPTION_ARBITRAGE", "HIGH"
        return "OPTION_UNDERLYING", "OPTION_ARBITRAGE", "MEDIUM"
    if "COMMODITY_OPTION" in subtypes and "FUTURE" in types:
        option = a if a["subtype"] == "COMMODITY_OPTION" else b
        other = b if a["subtype"] == "COMMODITY_OPTION" else a
        opt_row = options_by_id.get(option["asset_id"])
        if opt_row is not None and clean_text(opt_row.get("underlying_asset_id")) == other["asset_id"]:
            return "OPTION_UNDERLYING", "OPTION_FUTURE_ARBITRAGE", "HIGH"
        return "OPTION_UNDERLYING", "OPTION_FUTURE_ARBITRAGE", "MEDIUM"
    if a["asset_type"] == "FUTURE" and b["asset_type"] == "FUTURE" and a["exchange_code"] != b["exchange_code"]:
        return "SAME_UNDERLYING_CROSS_EXCHANGE", "CROSS_EXCHANGE_ARBITRAGE", "MEDIUM"
    return None


def direction_for(relation: str, a_type: str, b_type: str) -> tuple[str, str]:
    if a_type == "FUTURE" and b_type == "FUTURE":
        return "BIDIRECTIONAL", "两个期货资产均可双向交易，跨交易所方向需结合合约规则确认。"
    if relation == "FUTURE_ETF":
        return "LONG_ETF_AND_SHORT_FUTURE_NEED_CHECK", "常见方向为多 ETF 空期货或反向，需结合融券、保证金和基差状态确认。"
    if relation in {"ETF_INDEX", "FUTURE_INDEX", "OPTION_INDEX"}:
        return "INDEX_REFERENCE_ONLY", "指数为参考标的，不直接交易。"
    if relation in {"OPTION_FUTURE", "OPTION_UNDERLYING"}:
        return "BIDIRECTIONAL_WITH_OPTION_MARGIN_CONSTRAINT", "期权买卖方向受保证金、权限和行权规则约束。"
    return "CHECK", "方向需后续按策略规则确认。"


def build_domestic_mapping(asset_master: pd.DataFrame, underlying_mapping: pd.DataFrame, options_with_etf: pd.DataFrame) -> pd.DataFrame:
    assets = asset_master.merge(underlying_mapping[["asset_id", "underlying_group"]], on="asset_id", suffixes=("", "_mapped"), how="left")
    assets["underlying_group"] = assets["underlying_group_mapped"].where(~assets["underlying_group_mapped"].apply(is_missing), assets["underlying_group"])
    assets = assets[~assets["underlying_group"].apply(is_missing)].copy()
    options_by_id = {row["asset_id"]: row for _, row in options_with_etf.iterrows()} if not options_with_etf.empty and "asset_id" in options_with_etf else {}
    rows = []
    seen = set()
    for group, group_df in assets.groupby("underlying_group"):
        records = [norm_asset(row) for _, row in group_df.iterrows()]
        for a, b in itertools.combinations(records, 2):
            pair = tuple(sorted([a["asset_id"], b["asset_id"]]))
            if pair in seen:
                continue
            relation = relation_for(a, b, options_by_id)
            if relation is None:
                continue
            seen.add(pair)
            relation_type, strategy_type, confidence = relation
            direction, note = direction_for(relation_type, a["asset_type"], b["asset_type"])
            tradable = "BOTH_TRADABLE" if a["tradable"] == "Y" and b["tradable"] == "Y" else "CHECK"
            rows.append(
                {
                    "mapping_id": f"{relation_type}_{pair[0]}_{pair[1]}",
                    "asset_id_a": a["asset_id"],
                    "asset_id_b": b["asset_id"],
                    "symbol_a": a["symbol"],
                    "symbol_b": b["symbol"],
                    "name_a": a["name_cn"],
                    "name_b": b["name_cn"],
                    "asset_type_a": a["asset_type"],
                    "asset_type_b": b["asset_type"],
                    "exchange_a": a["exchange_code"],
                    "exchange_b": b["exchange_code"],
                    "underlying_group": group,
                    "relation_type": relation_type,
                    "strategy_type": strategy_type,
                    "market_relation": "CN_DOMESTIC",
                    "tradable_check": tradable,
                    "direction_supported": direction,
                    "long_short_note": note,
                    "mapping_confidence": confidence,
                    "data_source_1": "asset_master",
                    "data_source_2": "underlying_mapping",
                    "source_status": "DOMESTIC_MAPPING_CANDIDATE_NEED_REVIEW",
                    "notes": "根据统一 underlying_group 自动生成国内跨市场映射候选，非套利计算结果。",
                }
            )
    mapping = pd.DataFrame(rows).drop_duplicates("mapping_id") if rows else pd.DataFrame()
    save_csv(mapping, FINAL_DIR / "domestic_cross_market_mapping.csv")
    return mapping


def build_underlying_quality(asset_master: pd.DataFrame, underlying_mapping: pd.DataFrame) -> pd.DataFrame:
    rows = [
        {
            "section": "overall",
            "asset_type": "ALL",
            "total_asset_count": len(asset_master),
            "underlying_mapping_count": len(underlying_mapping),
            "underlying_group_missing_count": int(underlying_mapping["underlying_group"].apply(is_missing).sum()),
            "mapping_confidence_high_count": int((underlying_mapping["mapping_confidence"] == "HIGH").sum()),
            "mapping_confidence_medium_count": int((underlying_mapping["mapping_confidence"] == "MEDIUM").sum()),
            "mapping_confidence_check_count": int((underlying_mapping["mapping_confidence"] == "CHECK").sum()),
            "total_count": len(asset_master),
            "mapping_count": len(underlying_mapping),
            "high_confidence_count": int((underlying_mapping["mapping_confidence"] == "HIGH").sum()),
            "check_count": int((underlying_mapping["mapping_confidence"] == "CHECK").sum()),
        }
    ]
    for asset_type, group in underlying_mapping.groupby("asset_type", dropna=False):
        rows.append(
            {
                "section": "by_asset_type",
                "asset_type": asset_type,
                "total_asset_count": "",
                "underlying_mapping_count": "",
                "underlying_group_missing_count": int(group["underlying_group"].apply(is_missing).sum()),
                "mapping_confidence_high_count": "",
                "mapping_confidence_medium_count": "",
                "mapping_confidence_check_count": "",
                "total_count": len(group),
                "mapping_count": len(group),
                "high_confidence_count": int((group["mapping_confidence"] == "HIGH").sum()),
                "check_count": int((group["mapping_confidence"] == "CHECK").sum()),
            }
        )
    report = pd.DataFrame(rows)
    save_csv(report, PROCESSED_DIR / "underlying_mapping_quality_report.csv")
    return report


def build_domestic_quality(mapping: pd.DataFrame) -> pd.DataFrame:
    rows = [
        {
            "section": "overall",
            "relation_type": "ALL",
            "count": len(mapping),
            "high_confidence_count": int((mapping["mapping_confidence"] == "HIGH").sum()) if not mapping.empty else 0,
            "medium_confidence_count": int((mapping["mapping_confidence"] == "MEDIUM").sum()) if not mapping.empty else 0,
            "domestic_mapping_count": len(mapping),
            "future_etf_mapping_count": int((mapping["relation_type"] == "FUTURE_ETF").sum()) if not mapping.empty else 0,
            "future_index_mapping_count": int((mapping["relation_type"] == "FUTURE_INDEX").sum()) if not mapping.empty else 0,
            "etf_index_mapping_count": int((mapping["relation_type"] == "ETF_INDEX").sum()) if not mapping.empty else 0,
            "option_index_mapping_count": int((mapping["relation_type"] == "OPTION_INDEX").sum()) if not mapping.empty else 0,
            "option_future_mapping_count": int((mapping["relation_type"] == "OPTION_FUTURE").sum()) if not mapping.empty else 0,
            "option_underlying_mapping_count": int((mapping["relation_type"] == "OPTION_UNDERLYING").sum()) if not mapping.empty else 0,
            "cross_exchange_mapping_count": int((mapping["relation_type"] == "SAME_UNDERLYING_CROSS_EXCHANGE").sum()) if not mapping.empty else 0,
        }
    ]
    if not mapping.empty:
        for relation, group in mapping.groupby("relation_type"):
            rows.append(
                {
                    "section": "by_relation_type",
                    "relation_type": relation,
                    "count": len(group),
                    "high_confidence_count": int((group["mapping_confidence"] == "HIGH").sum()),
                    "medium_confidence_count": int((group["mapping_confidence"] == "MEDIUM").sum()),
                }
            )
    report = pd.DataFrame(rows)
    save_csv(report, PROCESSED_DIR / "domestic_mapping_quality_report.csv")
    return report


def readme_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ("项目", "Step 5 统一标的映射和国内跨市场映射草稿"),
            ("数据基础", "基于 Stage 1 期货、Stage 2 期权、Stage 4 ETF/指数。"),
            ("underlying_mapping", "用于统一资产标的。"),
            ("domestic_cross_market_mapping", "用于后续期现套利、ETF折溢价套利、期权标的套利和跨市场套利。"),
            ("复核说明", "mapping_confidence 为 CHECK 的记录需要后续人工复核。"),
            ("范围说明", "当前不包含国外跨市场映射和产业链映射。"),
        ],
        columns=["item", "description"],
    )


def write_excel(
    underlying_mapping: pd.DataFrame,
    domestic_mapping: pd.DataFrame,
    underlying_quality: pd.DataFrame,
    domestic_quality: pd.DataFrame,
    asset_master: pd.DataFrame,
) -> None:
    sheets = {
        "README": readme_df(),
        "underlying_mapping": underlying_mapping,
        "domestic_cross_market_mapping": domestic_mapping,
        "underlying_mapping_quality_report": underlying_quality,
        "domestic_mapping_quality_report": domestic_quality,
        "asset_master_snapshot": asset_master,
    }
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
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
    asset_master, source_path = choose_asset_master()
    read_csv_optional(FUTURES_MASTER)
    read_csv_optional(OPTIONS_MASTER)
    read_csv_optional(ETF_MASTER)
    read_csv_optional(INDEX_MASTER)
    options_with_etf = read_csv_optional(OPTIONS_WITH_ETF)
    read_csv_optional(ETF_INDEX_MAPPING)

    underlying_mapping = build_underlying_mapping(asset_master)
    domestic_mapping = build_domestic_mapping(asset_master, underlying_mapping, options_with_etf)
    underlying_quality = build_underlying_quality(asset_master, underlying_mapping)
    domestic_quality = build_domestic_quality(domestic_mapping)
    write_excel(underlying_mapping, domestic_mapping, underlying_quality, domestic_quality, asset_master)

    print(f"使用的 asset_master 来源: {source_path}")
    print(f"asset 总数: {len(asset_master)}")
    print(f"underlying_mapping 行数: {len(underlying_mapping)}")
    print(f"underlying_group 缺失数量: {int(underlying_mapping['underlying_group'].apply(is_missing).sum())}")
    print(f"domestic_cross_market_mapping 行数: {len(domestic_mapping)}")
    if not domestic_mapping.empty:
        print("各 relation_type 数量:")
        for relation, count in domestic_mapping["relation_type"].value_counts().sort_index().items():
            print(f"  {relation}: {count}")
    print(f"Excel 输出路径: {EXCEL_PATH}")


if __name__ == "__main__":
    main()
