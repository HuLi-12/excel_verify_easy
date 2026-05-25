import re

from generate_template.main import build_template_output_path
from settings import DATA_PATH, TEMPLATE_PATH
from verify.main import build_output_path


def test_output_path_uses_high_resolution_timestamp():
    output_path = build_output_path()

    assert re.fullmatch(rf"output_{re.escape(DATA_PATH.stem)}_\d{{8}}_\d{{6}}_\d{{6}}\.xlsx", output_path.name)
    assert output_path.parent.name == "data"


def test_template_output_path_uses_high_resolution_timestamp():
    output_path = build_template_output_path()

    template_stem = TEMPLATE_PATH.stem.split("(")[0].strip()
    assert re.fullmatch(rf"填写模板_{re.escape(template_stem)}_\d{{8}}_\d{{6}}_\d{{6}}\.xlsx", output_path.name)
    assert output_path.parent.name == "template"
