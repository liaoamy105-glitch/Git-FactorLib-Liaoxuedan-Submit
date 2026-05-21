import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import tushare as ts


ETF_OPTION_EXCHANGES = ["SSE", "SZSE"]
INDEX_CODES = ["000016.SH", "000300.SH", "000905.SH", "000852.SH"]

ETF_OPTIONS_OUTPUT = Path("data/raw/tushare_etf_options_basic.csv")
FUND_BASIC_OUTPUT = Path("data/raw/tushare_fund_basic.csv")
INDEX_BASIC_OUTPUT = Path("data/raw/tushare_index_basic.csv")
ERROR_LOG = Path("data/raw/tushare_etf_fetch_errors.log")


def write_error(message):
    ERROR_LOG.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with ERROR_LOG.open("a", encoding="utf-8") as file:
        file.write(f"[{timestamp}] {message}\n")


def save_if_not_empty(data, output_path):
    if data is None or data.empty:
        return False

    output_path.parent.mkdir(parents=True, exist_ok=True)
    data.to_csv(output_path, index=False, encoding="utf-8-sig")
    return True


def fetch_etf_options_basic(pro):
    frames = []

    for exchange in ETF_OPTION_EXCHANGES:
        try:
            data = pro.opt_basic(exchange=exchange)
            if data is not None and not data.empty:
                data = data.copy()
                data["fetch_exchange"] = exchange
                frames.append(data)
        except Exception as error:
            write_error(f"Failed to fetch ETF options basic for {exchange}: {error}")

    if not frames:
        write_error("ETF options basic fetch returned no data.")
        return False

    return save_if_not_empty(
        pd.concat(frames, ignore_index=True),
        ETF_OPTIONS_OUTPUT,
    )


def fetch_fund_basic(pro):
    try:
        data = pro.fund_basic(market="E")
    except Exception as error:
        write_error(f"Failed to fetch fund basic with market='E': {error}")
        return False

    if data is None or data.empty:
        write_error("Fund basic fetch with market='E' returned no data.")
        return False

    return save_if_not_empty(data, FUND_BASIC_OUTPUT)


def fetch_index_basic(pro):
    frames = []

    try:
        index_basic = pro.index_basic()
        if index_basic is not None and not index_basic.empty:
            if "ts_code" in index_basic.columns:
                index_basic = index_basic[index_basic["ts_code"].isin(INDEX_CODES)]
            if not index_basic.empty:
                frames.append(index_basic)
    except Exception as error:
        write_error(f"Failed to fetch index basic: {error}")

    if frames:
        return save_if_not_empty(
            pd.concat(frames, ignore_index=True).drop_duplicates(),
            INDEX_BASIC_OUTPUT,
        )

    fallback_frames = []
    for ts_code in INDEX_CODES:
        try:
            data = pro.index_dailybasic(ts_code=ts_code)
            if data is not None and not data.empty:
                fallback_frames.append(data)
        except Exception as error:
            write_error(f"Failed to fetch index dailybasic for {ts_code}: {error}")

    if fallback_frames:
        return save_if_not_empty(
            pd.concat(fallback_frames, ignore_index=True),
            INDEX_BASIC_OUTPUT,
        )

    for ts_code in INDEX_CODES:
        try:
            data = pro.index_daily(ts_code=ts_code)
            if data is not None and not data.empty:
                fallback_frames.append(data)
        except Exception as error:
            write_error(f"Failed to fetch index daily for {ts_code}: {error}")

    if fallback_frames:
        return save_if_not_empty(
            pd.concat(fallback_frames, ignore_index=True),
            INDEX_BASIC_OUTPUT,
        )

    write_error("Index basic fetch returned no data from available interfaces.")
    return False


def print_result(success, success_message, failure_message):
    if success:
        print(success_message)
    else:
        print(failure_message)


def main():
    token = os.getenv("TUSHARE_TOKEN")
    if not token:
        print("TUSHARE_TOKEN environment variable is not set.")
        print("Please set it before running this script.")
        return

    pro = ts.pro_api(token)

    etf_options_success = fetch_etf_options_basic(pro)
    fund_basic_success = fetch_fund_basic(pro)
    index_basic_success = fetch_index_basic(pro)

    print_result(
        etf_options_success,
        f"ETF options basic saved to {ETF_OPTIONS_OUTPUT}",
        f"ETF options fetch failed, see {ERROR_LOG}",
    )
    print_result(
        fund_basic_success,
        f"fund basic saved to {FUND_BASIC_OUTPUT}",
        f"fund basic fetch failed, see {ERROR_LOG}",
    )
    print_result(
        index_basic_success,
        f"index basic saved to {INDEX_BASIC_OUTPUT}",
        f"index basic fetch failed, see {ERROR_LOG}",
    )


if __name__ == "__main__":
    main()
