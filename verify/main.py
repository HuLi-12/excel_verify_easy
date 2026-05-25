from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from output_writer import validate_excel_and_write_output
from settings import DATA_PATH, HEADER_ROWS, SAMPLE_ROW, SHEET_NAME, TEMPLATE_PATH, build_data_output_path
from template_reader import read_template_rules


def build_output_path() -> Path:
    return build_data_output_path()


def build_rules() -> list[dict]:
    return read_template_rules(
        template_path=TEMPLATE_PATH,
        sheet_name=SHEET_NAME,
        header_rows=HEADER_ROWS,
        sample_row=SAMPLE_ROW,
    )


def verify_data_file() -> tuple[Path, dict[str, int]]:
    rules = build_rules()
    output_path = build_output_path()
    summary = validate_excel_and_write_output(
        data_path=DATA_PATH,
        output_path=output_path,
        rules=rules,
        sheet_name=SHEET_NAME,
    )
    print(
        "校验完成："
        f"错误 {summary['errors']} 个，"
        f"警告 {summary['warnings']} 个，"
        f"自动修正 {summary['normalized']} 个"
    )
    print(f"校验结果已生成：{output_path}")
    return output_path, summary


def main() -> None:
    verify_data_file()


if __name__ == "__main__":
    main()
