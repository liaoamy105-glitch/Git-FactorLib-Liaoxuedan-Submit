import re
from pathlib import Path

import pandas as pd


FUTURES_INPUT = Path("data/raw/tushare_futures_basic.csv")
OPTIONS_INPUT = Path("data/raw/tushare_options_basic.csv")
ETF_OPTIONS_INPUT = Path("data/raw/tushare_etf_options_basic.csv")
FUND_BASIC_INPUT = Path("data/raw/tushare_fund_basic.csv")
INDEX_BASIC_INPUT = Path("data/raw/tushare_index_basic.csv")
FOREIGN_MAPPING_INPUT = Path("data/processed/effective_foreign_mapping.csv")
MAPPING_SHEET_INPUT = Path("data/processed/mapping_sheet.csv")
DOMESTIC_MASTER_INPUT = Path("data/processed/domestic_master_with_etf.csv")

DETAIL_SHEET_OUTPUT = Path("data/processed/detail_sheet.csv")
CHECK_OUTPUT = Path("data/processed/detail_sheet_check.csv")
SUMMARY_OUTPUT = Path("data/processed/detail_sheet_summary.csv")

EXCHANGE_NAMES = {
    "SHFE": "上海期货交易所",
    "INE": "上海国际能源交易中心",
    "DCE": "大连商品交易所",
    "CZCE": "郑州商品交易所",
    "GFEX": "广州期货交易所",
    "CFFEX": "中国金融期货交易所",
    "SSE": "上海证券交易所",
    "SZSE": "深圳证券交易所",
}

CFFEX_OPTION_UNDERLYING = {
    "HO": "IH",
    "IO": "IF",
    "MO": "IM",
}

DETAIL_COLUMNS = [
    "品种类型",
    "完整合约代码",
    "合约短代码",
    "交易所代码",
    "交易所名称",
    "合约名称",
    "品种名称",
    "品种代码",
    "标的资产代码",
    "标的资产名称",
    "看涨看跌",
    "行权方式",
    "行权价格",
    "到期月份",
    "合约单位",
    "合约乘数",
    "报价单位",
    "最小变动价位",
    "交割方式",
    "合约月份规则",
    "合约数量规则",
    "合约数量规则标准化代码",
    "到期日规则",
    "最后交易日规则",
    "开始交易日/上市日",
    "最后交易日",
    "到期日/最后交割日",
    "数据来源",
    "核验状态",
    "备注",
]

RULES = {
    "FUTURE": {
        "合约月份规则": "Tushare期货基础信息自动获取，具体规则待核验",
        "合约数量规则": "待交易所核验",
        "合约数量规则标准化代码": "FUTURE_RULE_PENDING",
        "到期日规则": "Tushare日期字段自动获取，具体规则待核验",
        "最后交易日规则": "Tushare日期字段自动获取，具体规则待核验",
    },
    "OPTION": {
        "合约月份规则": "期货期权合约月份规则，待交易所核验",
        "合约数量规则": "待交易所核验",
        "合约数量规则标准化代码": "FUTURE_OPTION_RULE_PENDING",
        "到期日规则": "Tushare日期字段自动获取，具体规则待核验",
        "最后交易日规则": "Tushare日期字段自动获取，具体规则待核验",
    },
    "STOCK_INDEX_OPTION": {
        "合约月份规则": "当月、下2个月、随后3个季月",
        "合约数量规则": "三近三季",
        "合约数量规则标准化代码": "INDEX_OPTION_3_NEAR_3_QUARTER",
        "到期日规则": "到期月份第三个星期五，遇节假日顺延",
        "最后交易日规则": "到期月份第三个星期五，遇节假日顺延",
    },
    "ETF_OPTION": {
        "合约月份规则": "当月、下月及随后两个季月",
        "合约数量规则": "二近二季",
        "合约数量规则标准化代码": "ETF_OPTION_2_NEAR_2_QUARTER",
        "到期日规则": "到期月份第四个星期三，遇节假日顺延",
        "最后交易日规则": "到期月份第四个星期三，遇节假日顺延",
    },
    "SPOT": {
        "合约月份规则": "不适用",
        "合约数量规则": "不适用",
        "合约数量规则标准化代码": "不适用",
        "到期日规则": "不适用",
        "最后交易日规则": "不适用",
    },
    "FOREIGN_FUTURE": {
        "合约月份规则": "待核验",
        "合约数量规则": "待核验",
        "合约数量规则标准化代码": "待核验",
        "到期日规则": "待核验",
        "最后交易日规则": "待核验",
    },
}


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


def before_dot(value):
    return clean_text(value).split(".")[0]


def exchange_name(exchange_code):
    return EXCHANGE_NAMES.get(normalize_code(exchange_code), "")


def futures_product_name(name):
    return re.sub(r"\d+", "", clean_text(name)).strip()


def option_product_name(name):
    text = clean_text(name)
    if "期权" in text:
        return text.split("期权", 1)[0].strip() + "期权"
    text = re.sub(r"\d+(\.\d+)?", "", text)
    text = re.sub(r"[CP]", "", text, flags=re.IGNORECASE).strip()
    return f"{text}期权" if text else ""


def etf_option_product_name(name, underlying_code):
    text = clean_text(name)
    if "期权" in text:
        return text.split("期权", 1)[0].strip() + "期权"
    return f"{underlying_code}ETF期权"


def index_product_name(name):
    text = clean_text(name)
    if not text:
        return ""
    return text if text.endswith("指数") else f"{text}指数"


def extract_option_product_code(row):
    for column in ["opt_code", "ts_code"]:
        raw_code = clean_text(row.get(column, ""))
        if not raw_code:
            continue
        code = raw_code.split(".")[0].upper()
        code = re.sub(r"^OP", "", code, flags=re.IGNORECASE)
        match = re.match(r"^[A-Z]+", code)
        if match:
            return match.group(0)
    return ""


def extract_underlying_etf_code(row):
    opt_code = normalize_code(row.get("opt_code", ""))
    if opt_code.startswith("OP") and re.match(r"^OP\d{6}\.(SH|SZ)$", opt_code):
        return opt_code[2:]

    for column in ["ts_code", "name"]:
        value = normalize_code(row.get(column, ""))
        match = re.search(r"(?:OP)?(\d{6}\.(?:SH|SZ))", value)
        if match:
            return match.group(1)
    return ""


def exchange_from_ts_code(ts_code):
    code = normalize_code(ts_code)
    if code.endswith(".SH"):
        return "SSE"
    if code.endswith(".SZ"):
        return "SZSE"
    return ""


def row_base():
    return {column: "" for column in DETAIL_COLUMNS}


def apply_rules(row, instrument_type):
    if instrument_type in ["ETF_SPOT", "INDEX_SPOT"]:
        rules = RULES["SPOT"]
    else:
        rules = RULES[instrument_type]
    row.update(rules)
    return row


def build_product_name_lookup(domestic_master):
    lookup = {}
    if domestic_master.empty:
        return lookup
    for _, row in domestic_master.iterrows():
        key = (
            normalize_code(row.get("instrument_type", "")),
            normalize_code(row.get("product_code", "")),
            normalize_code(row.get("exchange_code", "")),
        )
        lookup[key] = clean_text(row.get("product_name", ""))
    return lookup


def lookup_underlying_name(product_lookup, product_code, exchange_code):
    key = ("FUTURE", normalize_code(product_code), normalize_code(exchange_code))
    return product_lookup.get(key, "")


def collect_mapping_sheet_codes(mapping_sheet):
    etf_codes = set()
    index_codes = set()
    if mapping_sheet.empty:
        return etf_codes, index_codes

    pattern = re.compile(r"\(([^()]+)\)")
    for _, row in mapping_sheet.iterrows():
        for column in ["ETF现货", "ETF期权", "现货/指数", "国内跨市场映射标的"]:
            for code in pattern.findall(clean_text(row.get(column, ""))):
                code = clean_text(code)
                if re.match(r"^\d{6}\.(SH|SZ)$", code, flags=re.IGNORECASE):
                    if column == "现货/指数":
                        index_codes.add(normalize_code(code))
                    else:
                        etf_codes.add(normalize_code(code))
                if code.upper().startswith("OP"):
                    etf_codes.add(normalize_code(code[2:]))
    return etf_codes, index_codes


def build_futures_rows(futures_raw):
    rows = []
    if futures_raw.empty:
        return rows
    for _, item in futures_raw.iterrows():
        exchange_code = normalize_code(item.get("exchange", ""))
        row = row_base()
        row.update(
            {
                "品种类型": "FUTURE",
                "完整合约代码": item.get("ts_code", ""),
                "合约短代码": item.get("symbol", ""),
                "交易所代码": exchange_code,
                "交易所名称": exchange_name(exchange_code),
                "合约名称": item.get("name", ""),
                "品种名称": futures_product_name(item.get("name", "")),
                "品种代码": normalize_code(item.get("fut_code", "")),
                "标的资产代码": "",
                "标的资产名称": "",
                "看涨看跌": "不适用",
                "行权方式": "不适用",
                "行权价格": "不适用",
                "到期月份": item.get("d_month", ""),
                "合约单位": item.get("trade_unit", ""),
                "合约乘数": item.get("per_unit", ""),
                "报价单位": item.get("quote_unit", ""),
                "最小变动价位": item.get("quote_unit_desc", ""),
                "交割方式": item.get("d_mode_desc", ""),
                "开始交易日/上市日": item.get("list_date", ""),
                "最后交易日": item.get("delist_date", ""),
                "到期日/最后交割日": item.get("last_ddate", ""),
                "数据来源": "Tushare期货基础信息",
                "核验状态": "已自动获取",
                "备注": "期货合约级明细",
            }
        )
        rows.append(apply_rules(row, "FUTURE"))
    return rows


def build_option_rows(options_raw, product_lookup):
    rows = []
    if options_raw.empty:
        return rows
    for _, item in options_raw.iterrows():
        exchange_code = normalize_code(item.get("exchange", ""))
        product_code = extract_option_product_code(item)
        instrument_type = "STOCK_INDEX_OPTION" if exchange_code == "CFFEX" else "OPTION"
        underlying_code = CFFEX_OPTION_UNDERLYING.get(product_code, product_code)
        row = row_base()
        row.update(
            {
                "品种类型": instrument_type,
                "完整合约代码": item.get("ts_code", ""),
                "合约短代码": before_dot(item.get("ts_code", "")),
                "交易所代码": exchange_code,
                "交易所名称": exchange_name(exchange_code),
                "合约名称": item.get("name", ""),
                "品种名称": option_product_name(item.get("name", "")),
                "品种代码": product_code,
                "标的资产代码": underlying_code,
                "标的资产名称": lookup_underlying_name(product_lookup, underlying_code, exchange_code),
                "看涨看跌": item.get("call_put", ""),
                "行权方式": item.get("exercise_type", ""),
                "行权价格": item.get("exercise_price", ""),
                "到期月份": item.get("s_month", ""),
                "合约单位": item.get("per_unit", ""),
                "合约乘数": item.get("per_unit", ""),
                "报价单位": item.get("quote_unit", ""),
                "最小变动价位": item.get("min_price_chg", ""),
                "交割方式": "现金交割" if exchange_code == "CFFEX" else "待交易所核验",
                "开始交易日/上市日": item.get("list_date", ""),
                "最后交易日": item.get("last_ddate", ""),
                "到期日/最后交割日": item.get("maturity_date", ""),
                "数据来源": "Tushare期权基础信息",
                "核验状态": "已自动获取",
                "备注": "期权合约级明细",
            }
        )
        rows.append(apply_rules(row, instrument_type))
    return rows


def build_fund_name_lookup(fund_basic):
    lookup = {}
    if fund_basic.empty or "ts_code" not in fund_basic.columns:
        return lookup
    for _, row in fund_basic.iterrows():
        lookup[normalize_code(row.get("ts_code", ""))] = clean_text(row.get("name", ""))
    return lookup


def build_etf_option_rows(etf_options_raw, fund_name_lookup):
    rows = []
    if etf_options_raw.empty:
        return rows
    for _, item in etf_options_raw.iterrows():
        underlying_code = extract_underlying_etf_code(item)
        exchange_code = normalize_code(item.get("exchange", "")) or normalize_code(
            item.get("fetch_exchange", "")
        )
        row = row_base()
        row.update(
            {
                "品种类型": "ETF_OPTION",
                "完整合约代码": item.get("ts_code", ""),
                "合约短代码": before_dot(item.get("ts_code", "")),
                "交易所代码": exchange_code,
                "交易所名称": exchange_name(exchange_code),
                "合约名称": item.get("name", ""),
                "品种名称": etf_option_product_name(item.get("name", ""), underlying_code),
                "品种代码": f"OP{underlying_code}" if underlying_code else "",
                "标的资产代码": underlying_code,
                "标的资产名称": fund_name_lookup.get(normalize_code(underlying_code), ""),
                "看涨看跌": item.get("call_put", ""),
                "行权方式": item.get("exercise_type", ""),
                "行权价格": item.get("exercise_price", ""),
                "到期月份": item.get("s_month", ""),
                "合约单位": item.get("per_unit", ""),
                "合约乘数": item.get("per_unit", ""),
                "报价单位": item.get("quote_unit", ""),
                "最小变动价位": item.get("min_price_chg", ""),
                "交割方式": "实物交割/证券给付，待交易所规则核验",
                "开始交易日/上市日": item.get("list_date", ""),
                "最后交易日": item.get("last_ddate", ""),
                "到期日/最后交割日": item.get("maturity_date", ""),
                "数据来源": "Tushare ETF期权基础信息",
                "核验状态": "已自动获取",
                "备注": "ETF期权合约级明细",
            }
        )
        rows.append(apply_rules(row, "ETF_OPTION"))
    return rows


def build_etf_spot_rows(fund_basic, referenced_etf_codes):
    rows = []
    if fund_basic.empty or "ts_code" not in fund_basic.columns:
        return rows
    fund_basic = fund_basic.copy()
    fund_basic["ts_code"] = fund_basic["ts_code"].apply(normalize_code)
    fund_basic = fund_basic[fund_basic["ts_code"].isin(referenced_etf_codes)]
    for _, item in fund_basic.iterrows():
        exchange_code = exchange_from_ts_code(item.get("ts_code", ""))
        row = row_base()
        row.update(
            {
                "品种类型": "ETF_SPOT",
                "完整合约代码": item.get("ts_code", ""),
                "合约短代码": item.get("ts_code", ""),
                "交易所代码": exchange_code,
                "交易所名称": exchange_name(exchange_code),
                "合约名称": item.get("name", ""),
                "品种名称": item.get("name", ""),
                "品种代码": item.get("ts_code", ""),
                "标的资产代码": item.get("benchmark", ""),
                "标的资产名称": "",
                "看涨看跌": "不适用",
                "行权方式": "不适用",
                "行权价格": "不适用",
                "到期月份": "不适用",
                "合约单位": "不适用",
                "合约乘数": "不适用",
                "报价单位": "不适用",
                "最小变动价位": "不适用",
                "交割方式": "不适用",
                "开始交易日/上市日": item.get("list_date", ""),
                "最后交易日": item.get("delist_date", ""),
                "到期日/最后交割日": "不适用",
                "数据来源": "Tushare基金基础信息",
                "核验状态": "已自动获取",
                "备注": "ETF现货补充行",
            }
        )
        rows.append(apply_rules(row, "ETF_SPOT"))
    return rows


def build_index_spot_rows(index_basic, referenced_index_codes):
    rows = []
    if index_basic.empty or "ts_code" not in index_basic.columns:
        return rows
    index_basic = index_basic.copy()
    index_basic["ts_code"] = index_basic["ts_code"].apply(normalize_code)
    if referenced_index_codes:
        index_basic = index_basic[index_basic["ts_code"].isin(referenced_index_codes)]
    index_basic = index_basic.drop_duplicates(subset=["ts_code"], keep="first")
    for _, item in index_basic.iterrows():
        exchange_code = normalize_code(item.get("market", "")) or exchange_from_ts_code(
            item.get("ts_code", "")
        )
        name = index_product_name(item.get("name", ""))
        row = row_base()
        row.update(
            {
                "品种类型": "INDEX_SPOT",
                "完整合约代码": item.get("ts_code", ""),
                "合约短代码": item.get("ts_code", ""),
                "交易所代码": exchange_code,
                "交易所名称": exchange_name(exchange_code),
                "合约名称": name,
                "品种名称": name,
                "品种代码": item.get("ts_code", ""),
                "标的资产代码": "",
                "标的资产名称": "",
                "看涨看跌": "不适用",
                "行权方式": "不适用",
                "行权价格": "不适用",
                "到期月份": "不适用",
                "合约单位": "不适用",
                "合约乘数": "不适用",
                "报价单位": "不适用",
                "最小变动价位": "不适用",
                "交割方式": "不适用",
                "开始交易日/上市日": item.get("list_date", ""),
                "最后交易日": "不适用",
                "到期日/最后交割日": "不适用",
                "数据来源": "Tushare指数基础信息",
                "核验状态": "已自动获取",
                "备注": "指数现货补充行",
            }
        )
        rows.append(apply_rules(row, "INDEX_SPOT"))
    return rows


def build_foreign_future_rows(foreign_mapping):
    rows = []
    if foreign_mapping.empty:
        return rows
    deduped = foreign_mapping.drop_duplicates(
        subset=["foreign_product_code", "foreign_exchange_code"], keep="first"
    )
    for _, item in deduped.iterrows():
        product_code = normalize_code(item.get("foreign_product_code", ""))
        exchange_code = normalize_code(item.get("foreign_exchange_code", ""))
        row = row_base()
        row.update(
            {
                "品种类型": "FOREIGN_FUTURE",
                "完整合约代码": f"{product_code}_{exchange_code}",
                "合约短代码": product_code,
                "交易所代码": exchange_code,
                "交易所名称": item.get("foreign_exchange_name", ""),
                "合约名称": item.get("foreign_product_name", ""),
                "品种名称": item.get("foreign_product_name", ""),
                "品种代码": product_code,
                "标的资产代码": "待核验",
                "标的资产名称": "待核验",
                "看涨看跌": "不适用",
                "行权方式": "不适用",
                "行权价格": "不适用",
                "到期月份": "待核验",
                "合约单位": "待核验",
                "合约乘数": "待核验",
                "报价单位": "待核验",
                "最小变动价位": "待核验",
                "交割方式": "待核验",
                "开始交易日/上市日": "",
                "最后交易日": "待核验",
                "到期日/最后交割日": "待核验",
                "数据来源": "人工国外映射规则",
                "核验状态": "人工映射待核验",
                "备注": f"国家地区={item.get('foreign_country', '')}; {item.get('mapping_note', '')}",
            }
        )
        rows.append(apply_rules(row, "FOREIGN_FUTURE"))
    return rows


def build_detail_sheet(
    futures_raw,
    options_raw,
    etf_options_raw,
    fund_basic,
    index_basic,
    foreign_mapping,
    mapping_sheet,
    domestic_master,
):
    product_lookup = build_product_name_lookup(domestic_master)
    fund_name_lookup = build_fund_name_lookup(fund_basic)
    referenced_etf_codes, referenced_index_codes = collect_mapping_sheet_codes(mapping_sheet)

    rows = []
    rows.extend(build_futures_rows(futures_raw))
    rows.extend(build_option_rows(options_raw, product_lookup))
    rows.extend(build_etf_option_rows(etf_options_raw, fund_name_lookup))
    rows.extend(build_etf_spot_rows(fund_basic, referenced_etf_codes))
    rows.extend(build_index_spot_rows(index_basic, referenced_index_codes))
    rows.extend(build_foreign_future_rows(foreign_mapping))

    detail_sheet = pd.DataFrame(rows, columns=DETAIL_COLUMNS).fillna("")
    detail_sheet = detail_sheet.drop_duplicates(
        subset=["品种类型", "完整合约代码", "交易所代码"], keep="first"
    )
    return detail_sheet


def add_check(checks, check_item, passed, message):
    checks.append(
        {
            "check_item": check_item,
            "status": "passed" if passed else "failed",
            "message": message,
        }
    )


def contains_row(detail_sheet, instrument_type, product_code, exchange_code=None):
    matched = detail_sheet[
        (detail_sheet["品种类型"] == instrument_type)
        & (detail_sheet["品种代码"].apply(normalize_code) == normalize_code(product_code))
    ]
    if exchange_code:
        matched = matched[
            matched["交易所代码"].apply(normalize_code) == normalize_code(exchange_code)
        ]
    return not matched.empty


def build_checks(detail_sheet):
    checks = []
    type_counts = detail_sheet["品种类型"].value_counts()

    add_check(
        checks,
        "future_contract_rows_gt_1000",
        int(type_counts.get("FUTURE", 0)) > 1000,
        f"FUTURE rows: {int(type_counts.get('FUTURE', 0))}",
    )
    add_check(
        checks,
        "option_contract_rows_gt_1000",
        int(type_counts.get("OPTION", 0)) > 1000,
        f"OPTION rows: {int(type_counts.get('OPTION', 0))}",
    )
    add_check(
        checks,
        "etf_option_contract_rows_gt_1000",
        int(type_counts.get("ETF_OPTION", 0)) > 1000,
        f"ETF_OPTION rows: {int(type_counts.get('ETF_OPTION', 0))}",
    )

    required_checks = [
        ("contains_PB_LME_foreign_lead", "FOREIGN_FUTURE", "PB", "LME"),
        ("contains_510050_SH_etf_spot", "ETF_SPOT", "510050.SH", "SSE"),
        ("contains_OP510050_SH_etf_option", "ETF_OPTION", "OP510050.SH", "SSE"),
        ("contains_000016_SH_index_spot", "INDEX_SPOT", "000016.SH", "SSE"),
    ]
    for check_item, instrument_type, product_code, exchange_code in required_checks:
        passed = contains_row(detail_sheet, instrument_type, product_code, exchange_code)
        add_check(
            checks,
            check_item,
            passed,
            "required item found"
            if passed
            else f"missing item: {instrument_type}/{product_code}/{exchange_code}",
        )

    stock_index_missing = []
    for product_code in ["HO", "IO", "MO"]:
        if not contains_row(detail_sheet, "STOCK_INDEX_OPTION", product_code, "CFFEX"):
            stock_index_missing.append(product_code)
    add_check(
        checks,
        "contains_HO_IO_MO_stock_index_options",
        not stock_index_missing,
        "HO/IO/MO stock index options found"
        if not stock_index_missing
        else "missing stock index option codes: " + "；".join(stock_index_missing),
    )

    duplicated = detail_sheet[
        detail_sheet.duplicated(
            subset=["品种类型", "完整合约代码", "交易所代码"], keep=False
        )
    ]
    add_check(
        checks,
        "detail_sheet_no_duplicate_type_contract_exchange",
        duplicated.empty,
        "no duplicated 品种类型 + 完整合约代码 + 交易所代码"
        if duplicated.empty
        else f"duplicated rows: {len(duplicated)}",
    )

    for column in ["完整合约代码", "品种代码", "交易所代码"]:
        empty_rows = detail_sheet[detail_sheet[column].apply(clean_text) == ""]
        add_check(
            checks,
            f"detail_sheet_no_empty_{column}",
            empty_rows.empty,
            f"no empty {column}" if empty_rows.empty else f"empty {column} rows: {len(empty_rows)}",
        )

    return pd.DataFrame(checks, columns=["check_item", "status", "message"])


def build_summary(detail_sheet, check_report):
    type_counts = detail_sheet["品种类型"].value_counts()
    rows = [
        {"metric": "detail_sheet_rows", "value": len(detail_sheet)},
        {"metric": "futures_contract_rows", "value": int(type_counts.get("FUTURE", 0))},
        {"metric": "options_contract_rows", "value": int(type_counts.get("OPTION", 0))},
        {
            "metric": "stock_index_option_contract_rows",
            "value": int(type_counts.get("STOCK_INDEX_OPTION", 0)),
        },
        {
            "metric": "etf_option_contract_rows",
            "value": int(type_counts.get("ETF_OPTION", 0)),
        },
        {"metric": "etf_spot_rows", "value": int(type_counts.get("ETF_SPOT", 0))},
        {"metric": "index_spot_rows", "value": int(type_counts.get("INDEX_SPOT", 0))},
        {
            "metric": "foreign_future_rows",
            "value": int(type_counts.get("FOREIGN_FUTURE", 0)),
        },
        {
            "metric": "check_failed_count",
            "value": int((check_report["status"] == "failed").sum()),
        },
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def main():
    futures_raw = read_csv(FUTURES_INPUT)
    options_raw = read_csv(OPTIONS_INPUT)
    etf_options_raw = read_csv(ETF_OPTIONS_INPUT)
    fund_basic = read_csv(FUND_BASIC_INPUT)
    index_basic = read_csv(INDEX_BASIC_INPUT)
    foreign_mapping = read_csv(FOREIGN_MAPPING_INPUT, required=True)
    mapping_sheet = read_csv(MAPPING_SHEET_INPUT, required=True)
    domestic_master = read_csv(DOMESTIC_MASTER_INPUT, required=True)

    detail_sheet = build_detail_sheet(
        futures_raw,
        options_raw,
        etf_options_raw,
        fund_basic,
        index_basic,
        foreign_mapping,
        mapping_sheet,
        domestic_master,
    )
    check_report = build_checks(detail_sheet)
    summary = build_summary(detail_sheet, check_report)

    DETAIL_SHEET_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    detail_sheet.to_csv(DETAIL_SHEET_OUTPUT, index=False, encoding="utf-8-sig")
    check_report.to_csv(CHECK_OUTPUT, index=False, encoding="utf-8-sig")
    summary.to_csv(SUMMARY_OUTPUT, index=False, encoding="utf-8-sig")

    check_failed_count = int((check_report["status"] == "failed").sum())

    print("detail sheet built successfully")
    print(f"detail_sheet rows: {len(detail_sheet)}")
    print(f"check failed count: {check_failed_count}")
    print(f"detail sheet saved to {DETAIL_SHEET_OUTPUT}")
    print(f"check report saved to {CHECK_OUTPUT}")
    print(f"summary saved to {SUMMARY_OUTPUT}")


if __name__ == "__main__":
    main()
