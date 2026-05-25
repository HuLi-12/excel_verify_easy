import json

from openpyxl import Workbook

from config_loader import load_app_config


def test_load_app_config_resolves_relative_paths(tmp_path):
    config_path = tmp_path / "app_config.json"
    config_path.write_text(
        json.dumps(
            {
                "template_path": "input/template/template.xlsx",
                "data_path": "input/data/data.xlsx",
                "sheet_name": "Sheet1",
                "header_rows": [1, 2],
                "sample_row": 3,
                "output": {
                    "data_dir": "output/data",
                    "template_dir": "output/template",
                    "data_prefix": "output",
                    "template_prefix": "\u586b\u5199\u6a21\u677f",
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    config = load_app_config(config_path=config_path, base_dir=tmp_path)

    assert config["template_path"] == tmp_path / "input" / "template" / "template.xlsx"
    assert config["data_path"] == tmp_path / "input" / "data" / "data.xlsx"
    assert config["sheet_name"] == "Sheet1"
    assert config["header_rows"] == [1, 2]
    assert config["sample_row"] == 3
    assert config["output"]["data_dir"] == tmp_path / "output" / "data"
    assert config["output"]["template_prefix"] == "\u586b\u5199\u6a21\u677f"


def test_load_app_config_uses_first_sheet_when_sheet_name_is_auto(tmp_path):
    template_path = tmp_path / "template.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "FirstSheet"
    wb.create_sheet("SecondSheet")
    wb.save(template_path)

    config_path = tmp_path / "app_config.json"
    config_path.write_text(
        json.dumps(
            {
                "template_path": "template.xlsx",
                "data_path": "data.xlsx",
                "sheet_name": "auto",
                "header_rows": [1],
                "sample_row": 2,
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    config = load_app_config(config_path=config_path, base_dir=tmp_path)

    assert config["sheet_name"] == "FirstSheet"
