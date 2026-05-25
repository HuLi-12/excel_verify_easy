from __future__ import annotations

from datetime import datetime
from pathlib import Path

from config_loader import load_app_config


BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"
APP_CONFIG_PATH = CONFIG_DIR / "app_config.json"
AUTO_MATCH_CONFIG_PATH = CONFIG_DIR / "auto_match_rules.json"

APP_CONFIG = load_app_config(APP_CONFIG_PATH, BASE_DIR)

TEMPLATE_PATH = APP_CONFIG["template_path"]
DATA_PATH = APP_CONFIG["data_path"]
HEADER_ROWS = APP_CONFIG["header_rows"]
SAMPLE_ROW = APP_CONFIG["sample_row"]
SHEET_NAME = APP_CONFIG["sheet_name"]
OUTPUT_CONFIG = APP_CONFIG["output"]


def build_timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S_%f")


def build_data_output_path() -> Path:
    data_name = DATA_PATH.stem
    prefix = OUTPUT_CONFIG["data_prefix"]
    return OUTPUT_CONFIG["data_dir"] / f"{prefix}_{data_name}_{build_timestamp()}.xlsx"


def build_template_output_path() -> Path:
    stem = TEMPLATE_PATH.stem.split("(")[0].strip()
    prefix = OUTPUT_CONFIG["template_prefix"]
    return OUTPUT_CONFIG["template_dir"] / f"{prefix}_{stem}_{build_timestamp()}.xlsx"
