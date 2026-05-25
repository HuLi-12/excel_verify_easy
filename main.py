from __future__ import annotations

from generate_template.main import build_output_path as build_template_output_path
from generate_template.main import generate_template_file
from verify.main import build_output_path
from verify.main import verify_data_file


def main() -> None:
    """兼容入口：仍然顺序执行生成模板和数据校验。"""
    generate_template_file()
    print()
    verify_data_file()


if __name__ == "__main__":
    main()
