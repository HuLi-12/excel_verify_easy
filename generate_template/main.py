from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from settings import AUTO_MATCH_CONFIG_PATH, HEADER_ROWS, SAMPLE_ROW, SHEET_NAME, TEMPLATE_PATH, build_template_output_path
from template_generator import generate_fill_template
from template_reader import read_template_rules


def build_rules() -> list[dict]:
    return read_template_rules(
        template_path=TEMPLATE_PATH,
        sheet_name=SHEET_NAME,
        header_rows=HEADER_ROWS,
        sample_row=SAMPLE_ROW,
    )


def build_output_path() -> Path:
    return build_template_output_path()


def generate_template_file() -> Path:
    rules = build_rules()
    output_path = build_output_path()
    generate_fill_template(
        template_path=TEMPLATE_PATH,
        output_path=output_path,
        rules=rules,
        sheet_name=SHEET_NAME,
        header_rows=HEADER_ROWS,
        sample_row=SAMPLE_ROW,
        auto_match_config_path=AUTO_MATCH_CONFIG_PATH,
    )
    print(f"填写模板已生成：{output_path}")
    print(f"规则数量：{len(rules)} 列")
    return output_path


def main() -> None:
    generate_template_file()


if __name__ == "__main__":
    main()
