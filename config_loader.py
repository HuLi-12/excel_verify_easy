from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from template_reader import get_first_sheet_name


DEFAULT_OUTPUT = {
    "data_dir": "output/data",
    "template_dir": "output/template",
    "data_prefix": "output",
    "template_prefix": "填写模板",
}


def load_app_config(config_path: str | Path, base_dir: str | Path) -> dict[str, Any]:
    path = Path(config_path)
    root = Path(base_dir)

    with path.open("r", encoding="utf-8") as file:
        raw_config = json.load(file)

    if not isinstance(raw_config, dict):
        raise ValueError("app_config.json 必须是 JSON object")

    template_path = resolve_path(raw_config["template_path"], root)
    data_path = resolve_path(raw_config["data_path"], root)
    output_config = build_output_config(raw_config.get("output", {}), root)
    sheet_name = normalize_sheet_name(raw_config.get("sheet_name"), template_path)

    return {
        "template_path": template_path,
        "data_path": data_path,
        "sheet_name": sheet_name,
        "header_rows": normalize_int_list(raw_config.get("header_rows", [1])),
        "sample_row": int(raw_config.get("sample_row", 2)),
        "output": output_config,
    }


def resolve_path(value: str | Path, base_dir: Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return base_dir / path


def build_output_config(raw_output: dict[str, Any], base_dir: Path) -> dict[str, Any]:
    if not isinstance(raw_output, dict):
        raw_output = {}

    merged = {**DEFAULT_OUTPUT, **raw_output}
    return {
        "data_dir": resolve_path(merged["data_dir"], base_dir),
        "template_dir": resolve_path(merged["template_dir"], base_dir),
        "data_prefix": str(merged["data_prefix"]),
        "template_prefix": str(merged["template_prefix"]),
    }


def normalize_sheet_name(sheet_name: Any, template_path: Path) -> str:
    if sheet_name is None:
        return get_first_sheet_name(template_path)

    text = str(sheet_name).strip()
    if not text or text.lower() == "auto":
        return get_first_sheet_name(template_path)

    return text


def normalize_int_list(values: Any) -> list[int]:
    if not isinstance(values, list) or not values:
        raise ValueError("header_rows 必须是非空数组")

    return [int(value) for value in values]
