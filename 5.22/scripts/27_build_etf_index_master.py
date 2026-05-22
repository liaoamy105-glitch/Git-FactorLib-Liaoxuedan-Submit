from __future__ import annotations

import importlib.util
import re
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
STAGE_DIR = ROOT / "data" / "stage4_etf_index"
RAW_TUSHARE = STAGE_DIR / "raw" / "tushare"
RAW_AKSHARE = STAGE_DIR / "raw" / "akshare"
PROCESSED_DIR = STAGE_DIR / "processed"
FINAL_DIR = STAGE_DIR / "final"
OUTPUT_DIR = ROOT / "output" / "stage4_etf_index"
EXCEL_PATH = ROOT / "output" / "mapping_info_stage4_etf_index.xlsx"
ENCODING = "utf-8-sig"
UPDATE_DATE = date.today().strftime("%Y-%m-%d")

STAGE2_ASSET_PATH = ROOT / "data" / "stage2_options" / "final" / "asset_master_with_options_tushare_primary_final.csv"
STAGE2_OPTIONS_PATH = ROOT / "data" / "stage2_options" / "final" / "options_master_tushare_primary_final.csv"

INDEX_MARKETS = ["SSE", "SZSE", "CSI", "CNI", "SW", "MSCI", "OTH"]
EXCHANGE_NAME = {"SSE": "上海证券交易所", "SZSE": "深圳证券交易所"}
ETF_GROUP_MAPPING = {
    "510050": "SSE50",
    "510300": "CSI300",
    "510500": "CSI500",
    "510880": "SSE_DIVIDEND",
    "588000": "STAR50",
    "588080": "STAR50",
    "588050": "STAR50",
    "159919": "CSI300",
    "159922": "CSI500",
    "159915": "CHINEXT",
    "159901": "SZSE100",
    "159949": "CHINEXT50",
    "159995": "CHIP",
    "159601": "CSI500",
    "159605": "CSI500",
    "159845": "CSI1000",
    "159629": "CSI1000",
}
CORE_INDEX = {
    "CSI300": {"ts_codes": ["000300.SH", "000300.CSI"], "name": "沪深300指数", "publisher": "中证指数有限公司"},
    "SSE50": {"ts_codes": ["000016.SH"], "name": "上证50指数", "publisher": "上海证券交易所"},
    "CSI500": {"ts_codes": ["000905.SH", "000905.CSI"], "name": "中证500指数", "publisher": "中证指数有限公司"},
    "CSI1000": {"ts_codes": ["000852.SH", "000852.CSI"], "name": "中证1000指数", "publisher": "中证指数有限公司"},
    "CHINEXT": {"ts_codes": ["399006.SZ"], "name": "创业板指数", "publisher": "深圳证券交易所"},
    "STAR50": {"ts_codes": ["000688.SH"], "name": "科创50指数", "publisher": "上海证券交易所"},
    "SZSE100": {"ts_codes": ["399330.SZ"], "name": "深证100指数", "publisher": "深圳证券交易所"},
    "SSE_DIVIDEND": {"ts_codes": ["000015.SH"], "name": "上证红利指数", "publisher": "上海证券交易所"},
    "CHINEXT50": {"ts_codes": ["399673.SZ"], "name": "创业板50指数", "publisher": "深圳证券交易所"},
    "CHIP": {"ts_codes": [], "name": "芯片指数", "publisher": "CHECK"},
}
FUND_FIELDS = ",".join(
    [
        "ts_code",
        "name",
        "management",
        "custodian",
        "fund_type",
        "found_date",
        "due_date",
        "list_date",
        "issue_date",
        "delist_date",
        "issue_amount",
        "m_fee",
        "c_fee",
        "duration_year",
        "p_value",
        "min_amount",
        "exp_return",
        "benchmark",
        "status",
        "invest_type",
        "type",
        "trustee",
        "purc_startdate",
        "redm_startdate",
        "market",
    ]
)


def ensure_dirs() -> None:
    for path in [RAW_TUSHARE, RAW_AKSHARE, PROCESSED_DIR, FINAL_DIR, OUTPUT_DIR, EXCEL_PATH.parent]:
        path.mkdir(parents=True, exist_ok=True)


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def is_missing(value: Any) -> bool:
    text = clean_text(value)
    return text == "" or text.upper() in {"CHECK", "TODO", "NA", "NAN", "NONE", "CHECK_WAIT_ETF_MASTER"}


def save_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding=ENCODING)


def read_csv(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_csv(path, dtype=str).fillna("")
    return pd.DataFrame()


def read_config() -> dict[str, Any]:
    config_path = ROOT / "config_local.py"
    if not config_path.exists():
        return {}
    spec = importlib.util.spec_from_file_location("config_local", config_path)
    if spec is None or spec.loader is None:
        return {}
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return {name: getattr(module, name) for name in dir(module) if name.isupper()}


def exchange_from_code(etf_code: str, ts_code: str = "") -> str:
    suffix = clean_text(ts_code).split(".")[-1].upper() if "." in clean_text(ts_code) else ""
    if suffix == "SH":
        return "SSE"
    if suffix == "SZ":
        return "SZSE"
    if etf_code.startswith(("5", "6")):
        return "SSE"
    if etf_code.startswith(("1", "0")):
        return "SZSE"
    return "CHECK"


def ts_code_from_etf_code(etf_code: str) -> str:
    exchange = exchange_from_code(etf_code)
    suffix = "SH" if exchange == "SSE" else "SZ" if exchange == "SZSE" else "CHECK"
    return f"{etf_code}.{suffix}" if suffix != "CHECK" else etf_code


def etf_code_from_ts(ts_code: Any) -> str:
    text = clean_text(ts_code)
    match = re.search(r"(?<!\d)(\d{6})(?!\d)", text)
    return match.group(1) if match else ""


def infer_underlying_group_from_name(code: str, name: str, benchmark: str = "") -> str:
    if code in ETF_GROUP_MAPPING:
        return ETF_GROUP_MAPPING[code]
    text = f"{name} {benchmark}"
    rules = [
        ("沪深300", "CSI300"),
        ("上证50", "SSE50"),
        ("50ETF", "SSE50"),
        ("中证500", "CSI500"),
        ("中证1000", "CSI1000"),
        ("创业板50", "CHINEXT50"),
        ("创业板", "CHINEXT"),
        ("科创50", "STAR50"),
        ("红利", "SSE_DIVIDEND"),
        ("黄金", "GOLD"),
        ("有色", "NONFERROUS"),
        ("证券", "SECURITIES"),
        ("军工", "MILITARY"),
        ("芯片", "CHIP"),
    ]
    for keyword, group in rules:
        if keyword in text:
            return group
    return "CHECK"


def infer_etf_subtype(name: str, group: str) -> str:
    text = name + " " + group
    if any(k in text for k in ["黄金", "有色", "商品"]):
        return "COMMODITY_ETF"
    if any(k in text for k in ["债", "国债", "信用"]):
        return "BOND_ETF"
    if group in {"CSI300", "SSE50", "CSI500", "CSI1000", "CHINEXT", "STAR50", "SSE_DIVIDEND", "SZSE100", "CHINEXT50"}:
        return "BROAD_BASED_ETF"
    return "ETF"


def fetch_tushare() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    config = read_config()
    token = clean_text(config.get("TUSHARE_TOKEN", ""))
    if not token or "请在这里" in token:
        msg = "缺少 TUSHARE_TOKEN，跳过 Tushare，继续尝试 AKShare 和占位记录。"
        print(msg)
        summary = pd.DataFrame([{"market": m, "success": False, "rows": 0, "columns": "", "error": msg} for m in INDEX_MARKETS])
        save_csv(pd.DataFrame(), RAW_TUSHARE / "fund_basic_market_E.csv")
        save_csv(pd.DataFrame(), RAW_TUSHARE / "index_basic_all.csv")
        save_csv(summary, RAW_TUSHARE / "index_basic_download_summary.csv")
        return pd.DataFrame(), pd.DataFrame(), summary
    try:
        import tushare as ts
    except ImportError:
        msg = "缺少 tushare，请运行: pip install tushare"
        print(msg)
        summary = pd.DataFrame([{"market": m, "success": False, "rows": 0, "columns": "", "error": msg} for m in INDEX_MARKETS])
        save_csv(pd.DataFrame(), RAW_TUSHARE / "fund_basic_market_E.csv")
        save_csv(pd.DataFrame(), RAW_TUSHARE / "index_basic_all.csv")
        save_csv(summary, RAW_TUSHARE / "index_basic_download_summary.csv")
        return pd.DataFrame(), pd.DataFrame(), summary

    pro = ts.pro_api(token)
    try:
        fund = pro.fund_basic(market="E", fields=FUND_FIELDS)
    except Exception as exc:
        try:
            fund = pro.fund_basic(market="E")
        except Exception as fallback_exc:
            print(f"Tushare fund_basic 失败: {exc}; fallback failed: {fallback_exc}")
            fund = pd.DataFrame()
    save_csv(fund, RAW_TUSHARE / "fund_basic_market_E.csv")

    frames = []
    summary_rows = []
    for market in INDEX_MARKETS:
        error = ""
        df = pd.DataFrame()
        try:
            df = pro.index_basic(market=market)
        except Exception as exc:
            error = str(exc)
        if not df.empty:
            df["source_market_request"] = market
            save_csv(df, RAW_TUSHARE / f"index_basic_{market}.csv")
            frames.append(df)
        summary_rows.append(
            {"market": market, "success": not df.empty, "rows": len(df), "columns": "|".join(map(str, df.columns)), "error": error}
        )
        print(f"Tushare index_basic {market}: {len(df)} 行")
    index_all = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    summary = pd.DataFrame(summary_rows)
    save_csv(index_all, RAW_TUSHARE / "index_basic_all.csv")
    save_csv(summary, RAW_TUSHARE / "index_basic_download_summary.csv")
    print(f"Tushare fund_basic market=E: {len(fund)} 行")
    return fund, index_all, summary


def fetch_akshare() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    try:
        import akshare as ak
    except ImportError:
        msg = "缺少 akshare，请运行: pip install akshare"
        print(msg)
        log = pd.DataFrame(
            [{"function_name": fn, "exists": False, "success": False, "rows": 0, "columns": "", "error": msg, "notes": ""} for fn in ["fund_name_em", "fund_etf_spot_em", "index_stock_info"]]
        )
        save_csv(log, RAW_AKSHARE / "akshare_etf_index_call_log.csv")
        save_csv(pd.DataFrame(), RAW_AKSHARE / "akshare_etf_index_function_list.csv")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), log

    function_list = pd.DataFrame({"function_name": sorted([name for name in dir(ak) if "etf" in name.lower() or "index" in name.lower()])})
    save_csv(function_list, RAW_AKSHARE / "akshare_etf_index_function_list.csv")
    results: dict[str, pd.DataFrame] = {}
    logs = []
    for fn in ["fund_name_em", "fund_etf_spot_em", "index_stock_info"]:
        exists = hasattr(ak, fn)
        success = False
        rows = 0
        columns = ""
        error = ""
        if exists:
            try:
                df = getattr(ak, fn)()
                if isinstance(df, pd.DataFrame):
                    results[fn] = df
                    rows = len(df)
                    columns = "|".join(map(str, df.columns))
                    success = True
                    save_csv(df, RAW_AKSHARE / f"{fn}.csv")
                else:
                    error = f"返回类型不是 DataFrame: {type(df)}"
            except Exception as exc:
                error = str(exc)
        else:
            error = "函数不存在"
        logs.append({"function_name": fn, "exists": exists, "success": success, "rows": rows, "columns": columns, "error": error, "notes": ""})
        print(f"AKShare {fn}: exists={exists}, success={success}, rows={rows}")
    log = pd.DataFrame(logs)
    save_csv(log, RAW_AKSHARE / "akshare_etf_index_call_log.csv")
    return results.get("fund_name_em", pd.DataFrame()), results.get("fund_etf_spot_em", pd.DataFrame()), results.get("index_stock_info", pd.DataFrame()), log


def stage2_etf_codes(options_master: pd.DataFrame) -> set[str]:
    if options_master.empty:
        return set(ETF_GROUP_MAPPING)
    etf = options_master[options_master["subtype"] == "ETF_OPTION"].copy()
    codes = set()
    for _, row in etf.iterrows():
        for col in ["underlying_symbol", "option_symbol"]:
            code = etf_code_from_ts(row.get(col))
            if code:
                codes.add(code)
    return codes | set(ETF_GROUP_MAPPING)


def akshare_name_lookup(fund_name_em: pd.DataFrame) -> dict[str, dict[str, str]]:
    lookup: dict[str, dict[str, str]] = {}
    if fund_name_em.empty:
        return lookup
    for _, row in fund_name_em.iterrows():
        row_dict = {str(k): clean_text(v) for k, v in row.to_dict().items()}
        text = " ".join(row_dict.values())
        code = etf_code_from_ts(text)
        if code and code not in lookup:
            lookup[code] = row_dict
    return lookup


def build_etf_master(fund: pd.DataFrame, fund_name_em: pd.DataFrame, fund_etf_spot_em: pd.DataFrame, stage2_options: pd.DataFrame) -> pd.DataFrame:
    required_codes = stage2_etf_codes(stage2_options)
    ak_lookup = akshare_name_lookup(fund_name_em)
    rows_by_code: dict[str, dict[str, Any]] = {}
    if not fund.empty:
        for _, row in fund.iterrows():
            name = clean_text(row.get("name"))
            ts_code = clean_text(row.get("ts_code"))
            code = etf_code_from_ts(ts_code)
            filter_text = " ".join(clean_text(row.get(c)) for c in ["fund_type", "type", "invest_type", "name", "ts_code"])
            is_etf = "ETF" in filter_text.upper() or (code and name and "ETF" in name.upper())
            if not is_etf:
                continue
            exchange = exchange_from_code(code, ts_code)
            group = infer_underlying_group_from_name(code, name, row.get("benchmark", ""))
            rows_by_code[code] = {
                "asset_id": f"ETF_{exchange}_{code}",
                "asset_type": "ETF",
                "subtype": infer_etf_subtype(name, group),
                "etf_code": code,
                "ts_code": ts_code,
                "name_cn": name or f"{code}ETF",
                "name_en": "TODO",
                "exchange_code": exchange,
                "exchange_name": EXCHANGE_NAME.get(exchange, "CHECK"),
                "country": "CN",
                "currency": "CNY",
                "fund_company": clean_text(row.get("management")),
                "fund_type": clean_text(row.get("fund_type")) or clean_text(row.get("type")),
                "invest_type": clean_text(row.get("invest_type")),
                "tracking_index_code": "CHECK",
                "tracking_index_name": clean_text(row.get("benchmark")) or "CHECK",
                "underlying_group": group,
                "nav_available": "CHECK",
                "premium_discount_available": "CHECK",
                "can_subscribe_redeem": "CHECK",
                "creation_redemption_unit": "CHECK",
                "pcfl_available": "CHECK",
                "tradable": "Y",
                "can_long": "Y",
                "can_short": "CHECK",
                "data_source_1": "Tushare",
                "data_source_2": "AKShare" if not fund_name_em.empty or not fund_etf_spot_em.empty else "",
                "source_status": "ETF_MASTER_TUSHARE_AKSHARE_NEED_REVIEW",
                "update_date": UPDATE_DATE,
                "notes": "基于 Tushare fund_basic 场内基金自动筛选 ETF，ETF 申赎清单和折溢价字段后续补充。",
            }
    for code in sorted(required_codes):
        if code in rows_by_code:
            continue
        exchange = exchange_from_code(code)
        ak_info = ak_lookup.get(code, {})
        name = next((v for k, v in ak_info.items() if "名称" in k or "简称" in k), "") or f"{code}ETF"
        group = infer_underlying_group_from_name(code, name)
        rows_by_code[code] = {
            "asset_id": f"ETF_{exchange}_{code}",
            "asset_type": "ETF",
            "subtype": infer_etf_subtype(name, group),
            "etf_code": code,
            "ts_code": ts_code_from_etf_code(code),
            "name_cn": name,
            "name_en": "TODO",
            "exchange_code": exchange,
            "exchange_name": EXCHANGE_NAME.get(exchange, "CHECK"),
            "country": "CN",
            "currency": "CNY",
            "fund_company": "CHECK",
            "fund_type": "ETF",
            "invest_type": "CHECK",
            "tracking_index_code": "CHECK",
            "tracking_index_name": group,
            "underlying_group": group,
            "nav_available": "CHECK",
            "premium_discount_available": "CHECK",
            "can_subscribe_redeem": "CHECK",
            "creation_redemption_unit": "CHECK",
            "pcfl_available": "CHECK",
            "tradable": "Y",
            "can_long": "Y",
            "can_short": "CHECK",
            "data_source_1": "Tushare" if code in rows_by_code else "",
            "data_source_2": "AKShare" if code in ak_lookup else "",
            "source_status": "ETF_OPTION_UNDERLYING_PLACEHOLDER_NEED_REVIEW",
            "update_date": UPDATE_DATE,
            "notes": "由 ETF期权标的需求强制纳入，占位记录，后续 ETF 现货阶段需复核基金基础信息。",
        }
    columns = [
        "asset_id", "asset_type", "subtype", "etf_code", "ts_code", "name_cn", "name_en", "exchange_code", "exchange_name",
        "country", "currency", "fund_company", "fund_type", "invest_type", "tracking_index_code", "tracking_index_name",
        "underlying_group", "nav_available", "premium_discount_available", "can_subscribe_redeem", "creation_redemption_unit",
        "pcfl_available", "tradable", "can_long", "can_short", "data_source_1", "data_source_2", "source_status", "update_date", "notes",
    ]
    etf_master = pd.DataFrame(rows_by_code.values(), columns=columns).sort_values(["exchange_code", "etf_code"])
    save_csv(etf_master, FINAL_DIR / "etf_master.csv")
    return etf_master


def index_exchange_from_ts(ts_code: str) -> tuple[str, str]:
    suffix = clean_text(ts_code).split(".")[-1].upper() if "." in clean_text(ts_code) else ""
    if suffix == "SH":
        return "SSE", "上海证券交易所"
    if suffix == "SZ":
        return "SZSE", "深圳证券交易所"
    if suffix == "CSI":
        return "CSI", "中证指数有限公司"
    return suffix or "CHECK", "CHECK"


def build_index_master(index_all: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for group, info in CORE_INDEX.items():
        matched = pd.DataFrame()
        if not index_all.empty and "ts_code" in index_all.columns:
            matched = index_all[index_all["ts_code"].isin(info["ts_codes"])]
        if matched.empty and not index_all.empty and "name" in index_all.columns:
            names = [info["name"].replace("指数", ""), info["name"]]
            matched = index_all[index_all["name"].astype(str).apply(lambda x: any(n in x for n in names))]
        if not matched.empty:
            row = matched.iloc[0]
            ts_code = clean_text(row.get("ts_code"))
            name = clean_text(row.get("name")) or info["name"]
            exchange, exchange_name = index_exchange_from_ts(ts_code)
            source_status = "INDEX_MASTER_TUSHARE_NEED_REVIEW"
            notes = "核心指数来自 Tushare index_basic。"
        else:
            ts_code = info["ts_codes"][0] if info["ts_codes"] else f"CHECK_{group}"
            name = info["name"]
            exchange, exchange_name = index_exchange_from_ts(ts_code)
            source_status = "INDEX_PLACEHOLDER_NEED_REVIEW"
            notes = "核心指数占位记录，Tushare 未匹配到，后续需复核。"
        index_code = ts_code.split(".")[0] if "." in ts_code else ts_code
        rows.append(
            {
                "asset_id": f"INDEX_{group}",
                "asset_type": "INDEX",
                "subtype": "INDEX_SPOT",
                "index_code": index_code,
                "ts_code": ts_code,
                "name_cn": name,
                "name_en": "TODO",
                "exchange_code": exchange,
                "exchange_name": exchange_name,
                "publisher": info["publisher"],
                "category": "BROAD_BASED_INDEX",
                "market": exchange,
                "country": "CN",
                "currency": "CNY",
                "underlying_group": group,
                "tradable": "N",
                "can_long": "N",
                "can_short": "N",
                "data_source_1": "Tushare" if source_status == "INDEX_MASTER_TUSHARE_NEED_REVIEW" else "",
                "data_source_2": "",
                "source_status": source_status,
                "update_date": UPDATE_DATE,
                "notes": notes,
            }
        )
    index_master = pd.DataFrame(rows).sort_values("asset_id")
    save_csv(index_master, FINAL_DIR / "index_master.csv")
    return index_master


def update_options_underlying(options: pd.DataFrame, etf_master: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, int, int]:
    result = options.copy()
    etf_by_code = etf_master.set_index("etf_code").to_dict("index") if not etf_master.empty else {}
    before = 0
    rows = []
    for idx, row in result.iterrows():
        if clean_text(row.get("subtype")) != "ETF_OPTION":
            continue
        old_asset = clean_text(row.get("underlying_asset_id"))
        if is_missing(old_asset):
            before += 1
        code = etf_code_from_ts(row.get("underlying_symbol")) or etf_code_from_ts(row.get("option_symbol"))
        etf = etf_by_code.get(code)
        if etf:
            result.at[idx, "underlying_asset_id"] = etf["asset_id"]
            result.at[idx, "underlying_symbol"] = code
            result.at[idx, "underlying_group"] = etf["underlying_group"]
            result.at[idx, "notes"] = clean_text(row.get("notes")) + " | ETF option underlying_asset_id filled from etf_master"
        else:
            rows.append(
                {
                    "asset_id": row.get("asset_id", ""),
                    "option_symbol": row.get("option_symbol", ""),
                    "option_name_cn": row.get("option_name_cn", ""),
                    "underlying_symbol": row.get("underlying_symbol", ""),
                    "old_underlying_asset_id": old_asset,
                    "new_underlying_asset_id": "CHECK",
                    "underlying_group": row.get("underlying_group", ""),
                    "issue_reason": "ETF underlying not found in etf_master",
                    "suggested_next_step": "后续在 ETF 现货阶段确认 ETF asset_id 后回填。",
                }
            )
    after = int(result[(result["subtype"] == "ETF_OPTION") & (result["underlying_asset_id"].apply(is_missing))].shape[0])
    review = pd.DataFrame(rows)
    save_csv(result, FINAL_DIR / "options_master_with_etf_underlying.csv")
    save_csv(review, PROCESSED_DIR / "options_etf_underlying_review_list.csv")
    return result, review, before, after


def asset_rows_from_etf(etf_master: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "asset_id": etf_master["asset_id"],
            "asset_type": etf_master["asset_type"],
            "subtype": etf_master["subtype"],
            "symbol": etf_master["etf_code"],
            "name_cn": etf_master["name_cn"],
            "name_en": etf_master["name_en"],
            "exchange_code": etf_master["exchange_code"],
            "exchange_name": etf_master["exchange_name"],
            "country": etf_master["country"],
            "currency": etf_master["currency"],
            "underlying_group": etf_master["underlying_group"],
            "sector": "ETF",
            "tradable": etf_master["tradable"],
            "can_long": etf_master["can_long"],
            "can_short": etf_master["can_short"],
            "data_source_1": etf_master["data_source_1"],
            "data_source_2": etf_master["data_source_2"],
            "source_status": etf_master["source_status"],
            "update_date": etf_master["update_date"],
            "notes": etf_master["notes"],
        }
    )


def asset_rows_from_index(index_master: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "asset_id": index_master["asset_id"],
            "asset_type": index_master["asset_type"],
            "subtype": index_master["subtype"],
            "symbol": index_master["index_code"],
            "name_cn": index_master["name_cn"],
            "name_en": index_master["name_en"],
            "exchange_code": index_master["exchange_code"],
            "exchange_name": index_master["exchange_name"],
            "country": index_master["country"],
            "currency": index_master["currency"],
            "underlying_group": index_master["underlying_group"],
            "sector": "INDEX",
            "tradable": index_master["tradable"],
            "can_long": index_master["can_long"],
            "can_short": index_master["can_short"],
            "data_source_1": index_master["data_source_1"],
            "data_source_2": index_master["data_source_2"],
            "source_status": index_master["source_status"],
            "update_date": index_master["update_date"],
            "notes": index_master["notes"],
        }
    )


def build_asset_master(stage2_assets: pd.DataFrame, etf_master: pd.DataFrame, index_master: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    cols = list(stage2_assets.columns)
    etf_assets = asset_rows_from_etf(etf_master)
    index_assets = asset_rows_from_index(index_master)
    for df in [etf_assets, index_assets]:
        for col in cols:
            if col not in df.columns:
                df[col] = ""
    combined = pd.concat([stage2_assets, etf_assets[cols], index_assets[cols]], ignore_index=True)
    conflicts = combined[combined.duplicated("asset_id", keep=False)].copy()
    if not conflicts.empty:
        combined = combined.drop_duplicates("asset_id", keep="first")
    save_csv(conflicts, PROCESSED_DIR / "asset_id_conflicts_etf_index.csv")
    save_csv(combined, FINAL_DIR / "asset_master_with_etf_index.csv")
    return combined, conflicts


def safe_get(row: Any, candidates: str | list[str], default: str = "") -> str:
    if isinstance(candidates, str):
        candidates = [candidates]
    for candidate in candidates:
        try:
            if candidate in row.index:
                value = row.get(candidate)
            elif isinstance(row, dict) and candidate in row:
                value = row.get(candidate)
            else:
                continue
        except Exception:
            continue
        if value is not None and not pd.isna(value) and clean_text(value) != "":
            return clean_text(value)
    return default


def normalize_asset_record(row: Any, role: str = "asset") -> dict[str, str]:
    exchange = safe_get(row, ["exchange_code", "exchange"], "")
    underlying_group = safe_get(row, "underlying_group", "")
    etf_code = safe_get(row, ["etf_code", "symbol", "underlying_symbol"], "")
    option_symbol = safe_get(row, ["option_symbol", "symbol"], "")
    asset_id = safe_get(row, ["asset_id", "option_asset_id", "asset_id_x", "asset_id_y"], "")
    if not asset_id and role == "etf":
        code = safe_get(row, ["etf_code", "symbol", "ts_code"], "")
        code = etf_code_from_ts(code) or code
        asset_id = f"ETF_{exchange}_{code}" if exchange and code else ""
    elif not asset_id and role == "index":
        asset_id = f"INDEX_{underlying_group}" if underlying_group else ""
    elif not asset_id and role == "option":
        asset_id = f"OPT_{exchange}_{option_symbol}" if exchange and option_symbol else ""
    asset_type = safe_get(row, "asset_type", "")
    if not asset_type:
        asset_type = {"etf": "ETF", "index": "INDEX", "option": "OPTION"}.get(role, "")
    symbol = safe_get(row, ["symbol", "option_symbol", "etf_code", "index_code", "ts_code"], "")
    if role == "etf":
        symbol = etf_code_from_ts(symbol) or symbol
    return {
        "asset_id": asset_id,
        "symbol": symbol,
        "name_cn": safe_get(row, ["name_cn", "option_name_cn", "etf_name", "name", "tracking_index_name"], ""),
        "asset_type": asset_type,
        "subtype": safe_get(row, "subtype", ""),
        "exchange_code": exchange,
        "underlying_group": underlying_group,
        "tradable": safe_get(row, "tradable", ""),
        "can_long": safe_get(row, "can_long", ""),
        "can_short": safe_get(row, "can_short", ""),
    }


def columns_of(row: Any) -> str:
    try:
        return "|".join(map(str, row.index.tolist()))
    except Exception:
        if isinstance(row, dict):
            return "|".join(map(str, row.keys()))
    return ""


def make_mapping(
    asset_a: pd.Series,
    asset_b: pd.Series,
    relation: str,
    strategy: str,
    confidence: str,
    role_a: str = "asset",
    role_b: str = "asset",
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    a = normalize_asset_record(asset_a, role=role_a)
    b = normalize_asset_record(asset_b, role=role_b)
    if not a["asset_id"] or not b["asset_id"]:
        return None, {
            "relation_type": relation,
            "asset_a_raw_id": safe_get(asset_a, ["asset_id", "option_asset_id", "asset_id_x", "asset_id_y"], ""),
            "asset_b_raw_id": safe_get(asset_b, ["asset_id", "option_asset_id", "asset_id_x", "asset_id_y"], ""),
            "reason": "missing normalized asset_id",
            "asset_a_columns": columns_of(asset_a),
            "asset_b_columns": columns_of(asset_b),
            "notes": f"role_a={role_a}; role_b={role_b}",
        }
    mapping_id = f"{relation}_{a['asset_id']}_{b['asset_id']}"
    return {
        "mapping_id": mapping_id,
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
        "underlying_group": a["underlying_group"] or b["underlying_group"],
        "relation_type": relation,
        "strategy_type": strategy,
        "market_relation": "CN_DOMESTIC",
        "tradable_check": "Y",
        "direction_supported": "CHECK",
        "mapping_confidence": confidence,
        "data_source_1": "Stage4",
        "data_source_2": "Stage1/Stage2",
        "source_status": "MAPPING_CANDIDATE_NEED_REVIEW",
        "notes": "根据 underlying_group 自动生成映射候选，后续按套利场景复核。",
    }, None


def add_mapping(
    rows: list[dict[str, Any]],
    skips: list[dict[str, Any]],
    asset_a: pd.Series,
    asset_b: pd.Series,
    relation: str,
    strategy: str,
    confidence: str,
    role_a: str,
    role_b: str,
) -> None:
    mapping, skip = make_mapping(asset_a, asset_b, relation, strategy, confidence, role_a=role_a, role_b=role_b)
    if mapping is not None:
        rows.append(mapping)
    if skip is not None:
        skips.append(skip)


def build_mapping(asset_master: pd.DataFrame, options_with_etf: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    assets = asset_master.copy()
    by_id = assets.set_index("asset_id") if "asset_id" in assets.columns else pd.DataFrame()
    indexes = assets[assets["asset_type"] == "INDEX"]
    etfs = assets[assets["asset_type"] == "ETF"]
    futures = assets[(assets["asset_type"] == "FUTURE") & (assets["subtype"] == "INDEX_FUTURE")]
    index_options = assets[(assets["asset_type"] == "OPTION") & (assets["subtype"] == "INDEX_OPTION")]
    rows = []
    skips = []
    index_by_group = {row["underlying_group"]: row for _, row in indexes.iterrows() if not is_missing(row["underlying_group"])}

    for _, etf in etfs.iterrows():
        idx = index_by_group.get(etf["underlying_group"])
        if idx is not None:
            add_mapping(rows, skips, etf, idx, "ETF_INDEX", "ETF_PREMIUM_DISCOUNT_ARBITRAGE", "HIGH", "etf", "index")

    for _, opt in options_with_etf[options_with_etf["subtype"] == "ETF_OPTION"].iterrows():
        underlying_id = opt.get("underlying_asset_id", "")
        if underlying_id in by_id.index:
            opt_asset = opt if "asset_id" in opt.index else by_id.loc[underlying_id]
            if "asset_id" in opt.index and opt.get("asset_id") in by_id.index:
                opt_asset = by_id.loc[opt["asset_id"]]
            etf_asset = by_id.loc[underlying_id]
            add_mapping(rows, skips, opt_asset, etf_asset, "OPTION_UNDERLYING", "OPTION_ARBITRAGE", "HIGH", "option", "etf")

    for _, fut in futures.iterrows():
        idx = index_by_group.get(fut["underlying_group"])
        if idx is not None:
            add_mapping(rows, skips, fut, idx, "FUTURE_INDEX", "BASIS_ARBITRAGE", "HIGH", "asset", "index")
        for _, etf in etfs[etfs["underlying_group"] == fut["underlying_group"]].iterrows():
            add_mapping(rows, skips, fut, etf, "FUTURE_ETF", "BASIS_ARBITRAGE", "HIGH", "asset", "etf")

    for _, opt in index_options.iterrows():
        idx = index_by_group.get(opt["underlying_group"])
        if idx is not None:
            add_mapping(rows, skips, opt, idx, "OPTION_INDEX", "OPTION_ARBITRAGE", "HIGH", "option", "index")

    mapping = pd.DataFrame(rows).drop_duplicates("mapping_id") if rows else pd.DataFrame()
    skip_log = pd.DataFrame(
        skips,
        columns=["relation_type", "asset_a_raw_id", "asset_b_raw_id", "reason", "asset_a_columns", "asset_b_columns", "notes"],
    )
    save_csv(mapping, FINAL_DIR / "etf_index_mapping.csv")
    save_csv(skip_log, PROCESSED_DIR / "etf_index_mapping_skip_log.csv")
    print(f"etf_index_mapping 行数: {len(mapping)}")
    print(f"mapping_skip_log 行数: {len(skip_log)}")
    if not mapping.empty and "relation_type" in mapping.columns:
        print("各 relation_type 数量:")
        for relation_type, count in mapping["relation_type"].value_counts().sort_index().items():
            print(f"  {relation_type}: {count}")
    return mapping, skip_log


def build_quality_report(
    etf_master: pd.DataFrame,
    index_master: pd.DataFrame,
    asset_master: pd.DataFrame,
    options_with_etf: pd.DataFrame,
    conflicts: pd.DataFrame,
    mapping: pd.DataFrame,
) -> pd.DataFrame:
    rows = [
        {
            "section": "overall",
            "exchange_code": "ALL",
            "asset_type": "ALL",
            "count": "",
            "underlying_group_missing_count": "",
            "total_etf_count": len(etf_master),
            "total_index_count": len(index_master),
            "asset_master_total_count": len(asset_master),
            "etf_option_count": int((options_with_etf["subtype"] == "ETF_OPTION").sum()) if not options_with_etf.empty else 0,
            "etf_option_underlying_asset_id_missing_count": int(options_with_etf[(options_with_etf["subtype"] == "ETF_OPTION") & (options_with_etf["underlying_asset_id"].apply(is_missing))].shape[0]) if not options_with_etf.empty else 0,
            "etf_underlying_group_missing_count": int(etf_master["underlying_group"].apply(is_missing).sum()) if not etf_master.empty else 0,
            "index_underlying_group_missing_count": int(index_master["underlying_group"].apply(is_missing).sum()) if not index_master.empty else 0,
            "asset_id_duplicate_count": len(conflicts),
            "etf_index_mapping_count": len(mapping),
            "etf_index_mapping_high_confidence_count": int((mapping["mapping_confidence"] == "HIGH").sum()) if not mapping.empty else 0,
        }
    ]
    for (exchange, asset_type), group in asset_master.groupby(["exchange_code", "asset_type"], dropna=False):
        rows.append(
            {
                "section": "by_exchange_asset_type",
                "exchange_code": exchange,
                "asset_type": asset_type,
                "count": len(group),
                "underlying_group_missing_count": int(group["underlying_group"].apply(is_missing).sum()) if "underlying_group" in group else "",
                "total_etf_count": "",
                "total_index_count": "",
                "asset_master_total_count": "",
                "etf_option_count": "",
                "etf_option_underlying_asset_id_missing_count": "",
                "etf_underlying_group_missing_count": "",
                "index_underlying_group_missing_count": "",
                "asset_id_duplicate_count": "",
                "etf_index_mapping_count": "",
                "etf_index_mapping_high_confidence_count": "",
            }
        )
    report = pd.DataFrame(rows)
    save_csv(report, PROCESSED_DIR / "etf_index_quality_report.csv")
    return report


def readme_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ("项目", "Step 4 ETF现货和指数基础表"),
            ("数据源", "Tushare fund_basic、Tushare index_basic、AKShare fund/ETF接口。"),
            ("完成内容", "已构建 ETF 现货资产和 INDEX 指数资产。"),
            ("ETF期权", "已回填 ETF期权 underlying_asset_id。"),
            ("映射候选", "已生成 ETF-INDEX、ETF_OPTION-ETF、INDEX_FUTURE-INDEX、INDEX_OPTION-INDEX 等映射候选。"),
            ("待补字段", "ETF折溢价套利相关字段如申赎清单、折溢价率、申赎单位仍待后续基金公司/交易所数据源补充。"),
        ],
        columns=["item", "description"],
    )


def write_excel(
    asset_master: pd.DataFrame,
    etf_master: pd.DataFrame,
    index_master: pd.DataFrame,
    options_with_etf: pd.DataFrame,
    mapping: pd.DataFrame,
    quality: pd.DataFrame,
    review: pd.DataFrame,
    fund: pd.DataFrame,
    index_summary: pd.DataFrame,
    ak_log: pd.DataFrame,
) -> None:
    sheets = {
        "README": readme_df(),
        "asset_master_with_etf_index": asset_master,
        "etf_master": etf_master,
        "index_master": index_master,
        "options_master_with_etf_underlying": options_with_etf,
        "etf_index_mapping": mapping,
        "etf_index_quality_report": quality,
        "options_etf_underlying_review_list": review,
        "tushare_fund_basic": fund,
        "tushare_index_basic_summary": index_summary,
        "akshare_etf_index_call_log": ak_log,
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
    stage2_assets = read_csv(STAGE2_ASSET_PATH)
    stage2_options = read_csv(STAGE2_OPTIONS_PATH)
    fund, index_all, index_summary = fetch_tushare()
    fund_name_em, fund_etf_spot_em, _, ak_log = fetch_akshare()
    etf_master = build_etf_master(fund, fund_name_em, fund_etf_spot_em, stage2_options)
    index_master = build_index_master(index_all)
    options_with_etf, option_review, before_missing, after_missing = update_options_underlying(stage2_options, etf_master)
    asset_master, conflicts = build_asset_master(stage2_assets, etf_master, index_master)
    print(f"asset_master columns: {asset_master.columns.tolist()}")
    print(f"options_with_etf columns: {options_with_etf.columns.tolist()}")
    print(f"etf_master columns: {etf_master.columns.tolist()}")
    print(f"index_master columns: {index_master.columns.tolist()}")
    mapping, mapping_skip_log = build_mapping(asset_master, options_with_etf)
    quality = build_quality_report(etf_master, index_master, asset_master, options_with_etf, conflicts, mapping)
    write_excel(asset_master, etf_master, index_master, options_with_etf, mapping, quality, option_review, fund, index_summary, ak_log)

    print(f"Tushare fund_basic 行数: {len(fund)}")
    print(f"Tushare index_basic 总行数: {len(index_all)}")
    print(f"AKShare 成功函数数量: {int(ak_log['success'].sum()) if not ak_log.empty and 'success' in ak_log else 0}")
    print(f"etf_master 行数: {len(etf_master)}")
    print(f"index_master 行数: {len(index_master)}")
    print(f"asset_master_with_etf_index 行数: {len(asset_master)}")
    print(f"ETF期权 underlying_asset_id 缺失修复前数量: {before_missing}")
    print(f"ETF期权 underlying_asset_id 缺失修复后数量: {after_missing}")
    print(f"etf_index_mapping 行数: {len(mapping)}")
    print(f"ETF underlying_group 缺失数量: {int(etf_master['underlying_group'].apply(is_missing).sum()) if not etf_master.empty else 0}")
    print(f"index underlying_group 缺失数量: {int(index_master['underlying_group'].apply(is_missing).sum()) if not index_master.empty else 0}")
    print(f"Excel 输出路径: {EXCEL_PATH}")


if __name__ == "__main__":
    main()
