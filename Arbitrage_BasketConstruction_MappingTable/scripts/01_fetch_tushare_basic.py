import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import tushare as ts
import yaml


CONFIG_PATH = Path("config/config.yaml")
DEFAULT_FUTURES_EXCHANGES = ["SHFE", "INE", "DCE", "CZCE", "GFEX", "CFFEX"]
DEFAULT_FUTURES_OUTPUT = Path("data/raw/tushare_futures_basic.csv")
DEFAULT_OPTIONS_OUTPUT = Path("data/raw/tushare_options_basic.csv")
DEFAULT_ERROR_LOG = Path("data/raw/tushare_fetch_errors.log")


def load_config():
    if not CONFIG_PATH.exists():
        return {}

    with CONFIG_PATH.open("r", encoding="utf-8") as file:
        return yaml.safe_load(file) or {}


def write_error(log_path, message):
    log_path.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as file:
        file.write(f"[{timestamp}] {message}\n")


def fetch_futures_basic(pro, exchanges, output_path, error_log_path):
    frames = []

    for exchange in exchanges:
        try:
            data = pro.fut_basic(exchange=exchange)
            if data is not None and not data.empty:
                frames.append(data)
            print(f"futures basic fetched: {exchange}")
        except Exception as error:
            write_error(
                error_log_path,
                f"Failed to fetch futures basic for {exchange}: {error}",
            )
            print(f"futures basic fetch failed: {exchange}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if frames:
        pd.concat(frames, ignore_index=True).to_csv(output_path, index=False)
    else:
        pd.DataFrame().to_csv(output_path, index=False)


def fetch_options_basic_if_available(pro, exchanges, output_path, error_log_path):
    option_api = getattr(pro, "opt_basic", None)
    if option_api is None:
        write_error(error_log_path, "Options basic API is not available: opt_basic")
        print("options basic API is not available")
        return

    frames = []
    for exchange in exchanges:
        try:
            data = option_api(exchange=exchange)
            if data is not None and not data.empty:
                frames.append(data)
            print(f"options basic fetched: {exchange}")
        except Exception as error:
            write_error(
                error_log_path,
                f"Failed to fetch options basic for {exchange}: {error}",
            )
            print(f"options basic fetch failed: {exchange}")

    if frames:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        pd.concat(frames, ignore_index=True).to_csv(output_path, index=False)


def main():
    token = os.getenv("TUSHARE_TOKEN")
    if not token:
        print("TUSHARE_TOKEN environment variable is not set.")
        print("Please set it before running this script.")
        return

    config = load_config()
    tushare_config = config.get("tushare", {})

    exchanges = tushare_config.get("futures_exchanges", DEFAULT_FUTURES_EXCHANGES)
    futures_output = Path(
        tushare_config.get("futures_basic_output", DEFAULT_FUTURES_OUTPUT)
    )
    options_output = Path(
        tushare_config.get("options_basic_output", DEFAULT_OPTIONS_OUTPUT)
    )
    error_log = Path(tushare_config.get("error_log", DEFAULT_ERROR_LOG))

    pro = ts.pro_api(token)
    fetch_futures_basic(pro, exchanges, futures_output, error_log)
    fetch_options_basic_if_available(pro, exchanges, options_output, error_log)


if __name__ == "__main__":
    main()
