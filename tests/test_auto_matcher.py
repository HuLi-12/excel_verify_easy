import json

from auto_matcher import apply_auto_match, load_auto_match_config


def test_auto_match_is_disabled_when_config_file_is_missing(tmp_path):
    config = load_auto_match_config(tmp_path / "missing.json")

    result = apply_auto_match("\u5c0f\u5c0f\u7c73", "\u516c\u53f8", config)

    assert result["status"] == "PASS"
    assert result["output_value"] == "\u5c0f\u5c0f\u7c73"


def test_auto_match_normalizes_similar_company_value(tmp_path):
    config_path = tmp_path / "auto_match_rules.json"
    config_path.write_text(
        json.dumps(
            {
                "enabled": True,
                "rules": [
                    {
                        "field_keywords": ["\u516c\u53f8"],
                        "candidates": ["\u5c0f\u7c73\u516c\u53f8", "\u534e\u4e3a\u516c\u53f8"],
                        "auto_replace_threshold": 0.5,
                        "warning_threshold": 0.35,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    config = load_auto_match_config(config_path)

    assert apply_auto_match("\u5c0f\u5c0f\u7c73", "\u516c\u53f8", config)["output_value"] == "\u5c0f\u7c73\u516c\u53f8"
    assert apply_auto_match("\u534e\u4e3a\u6280", "\u516c\u53f8", config)["output_value"] == "\u534e\u4e3a\u516c\u53f8"


def test_auto_match_does_not_run_for_unconfigured_field():
    config = {
        "enabled": True,
        "rules": [
            {
                "field_keywords": ["\u516c\u53f8"],
                "candidates": ["\u5c0f\u7c73\u516c\u53f8"],
                "auto_replace_threshold": 0.5,
            }
        ],
    }

    result = apply_auto_match("\u5c0f\u5c0f\u7c73", "\u59d3\u540d", config)

    assert result["status"] == "PASS"
    assert result["output_value"] == "\u5c0f\u5c0f\u7c73"


def test_auto_match_keeps_standard_candidate_value_unchanged():
    config = {
        "enabled": True,
        "rules": [
            {
                "field_keywords": ["\u516c\u53f8"],
                "candidates": ["\u5c0f\u7c73\u516c\u53f8", "\u534e\u4e3a\u516c\u53f8"],
            }
        ],
    }

    result = apply_auto_match("\u5c0f\u7c73\u516c\u53f8", "\u6240\u5c5e\u516c\u53f8", config)

    assert result["status"] == "PASS"
    assert result["output_value"] == "\u5c0f\u7c73\u516c\u53f8"


def test_auto_match_uses_multi_score_for_typo_correction():
    config = {
        "enabled": True,
        "rules": [
            {
                "field_keywords": ["\u5355\u4f4d"],
                "candidates": ["\u57ce\u5efa\u96c6\u56e2", "\u4ea4\u6295\u96c6\u56e2"],
                "auto_replace_threshold": 0.70,
                "warning_threshold": 0.55,
                "min_score_gap": 0.12,
            }
        ],
    }

    result = apply_auto_match("\u6210\u5efa\u96c6\u56e2", "\u7ba1\u7406\u5355\u4f4d", config)

    assert result["status"] == "NORMALIZED"
    assert result["output_value"] == "\u57ce\u5efa\u96c6\u56e2"
    assert "score" not in result["message"]
    assert "gap" not in result["message"]


def test_auto_match_errors_when_top_candidates_are_ambiguous():
    config = {
        "enabled": True,
        "rules": [
            {
                "field_keywords": ["\u516c\u53f8"],
                "candidates": ["\u5c0f\u7c73\u516c\u53f8", "\u5c0f\u7c73\u96c6\u56e2"],
                "auto_replace_threshold": 0.5,
                "warning_threshold": 0.4,
                "min_score_gap": 0.2,
            }
        ],
    }

    result = apply_auto_match("\u5c0f\u7c73", "\u516c\u53f8", config)

    assert result["status"] == "ERROR"
    assert result["output_value"] == "/"


def test_auto_match_errors_when_no_candidate_is_confident():
    config = {
        "enabled": True,
        "rules": [
            {
                "field_keywords": ["\u516c\u53f8"],
                "candidates": ["\u5c0f\u7c73\u516c\u53f8", "\u534e\u4e3a\u516c\u53f8"],
                "auto_replace_threshold": 0.88,
                "min_score_gap": 0.12,
            }
        ],
    }

    result = apply_auto_match("\u5b8c\u5168\u4e0d\u76f8\u5173", "\u516c\u53f8", config)

    assert result["status"] == "ERROR"
    assert result["output_value"] == "/"
    assert "score" not in result["message"]
