from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
import json
from pathlib import Path
from typing import Any


DEFAULT_AUTO_REPLACE_THRESHOLD = 0.88
DEFAULT_MIN_SCORE_GAP = 0.12

SKIP_VALUES = {"", "/", "无", "暂无", "未提供", "不涉及", "至今"}


@dataclass(frozen=True)
class CandidateScore:
    value: str
    score: float


def load_auto_match_config(config_path: str | Path | None) -> dict[str, Any]:
    if config_path is None:
        return {"enabled": False, "rules": []}

    path = Path(config_path)
    if not path.exists():
        return {"enabled": False, "rules": []}

    with path.open("r", encoding="utf-8") as file:
        config = json.load(file)

    if not isinstance(config, dict):
        return {"enabled": False, "rules": []}

    config.setdefault("enabled", False)
    config.setdefault("rules", [])
    return config


def apply_auto_match(value: Any, field_name: str, config: dict[str, Any]) -> dict[str, Any]:
    original = "" if value is None else str(value).strip()
    if not config.get("enabled", False) or original in SKIP_VALUES:
        return _result("PASS", original, original, "未启用候选清单匹配")

    rule = _find_rule(field_name, config.get("rules", []))
    if rule is None:
        return _result("PASS", original, original, "字段未启用候选清单匹配")

    candidates = parse_candidates(rule.get("candidates", []))
    if not candidates:
        return _result("PASS", original, original, "未配置候选清单")

    if original in candidates:
        return _result("PASS", original, original, "已在候选清单中")

    normalized_matches = find_normalized_exact_matches(original, candidates)
    if len(normalized_matches) == 1:
        return _result("NORMALIZED", original, normalized_matches[0], "已按候选清单修正")
    if len(normalized_matches) > 1:
        return _result("ERROR", original, "/", "不在候选清单中，且候选项存在歧义")

    ranked = rank_candidates(original, candidates)
    if not ranked:
        return _result("ERROR", original, "/", "不在候选清单中，且未找到可信匹配")

    best = ranked[0]
    second_score = ranked[1].score if len(ranked) > 1 else 0.0
    score_gap = best.score - second_score
    auto_threshold = float(rule.get("auto_replace_threshold", rule.get("threshold", DEFAULT_AUTO_REPLACE_THRESHOLD)))
    min_score_gap = float(rule.get("min_score_gap", DEFAULT_MIN_SCORE_GAP))

    if best.score >= auto_threshold and score_gap >= min_score_gap:
        return _result("NORMALIZED", original, best.value, "已按候选清单修正")

    return _result("ERROR", original, "/", "不在候选清单中，且未找到可信匹配")


def parse_candidates(raw_candidates: list[Any]) -> list[str]:
    candidates = []
    for item in raw_candidates:
        if isinstance(item, dict):
            value = str(item.get("standard", "")).strip()
        else:
            value = str(item).strip()

        if value and value not in candidates:
            candidates.append(value)

    return candidates


def find_normalized_exact_matches(value: str, candidates: list[str]) -> list[str]:
    cleaned_value = clean_match_text(value)
    if not cleaned_value:
        return []

    return [candidate for candidate in candidates if clean_match_text(candidate) == cleaned_value]


def rank_candidates(value: str, candidates: list[str]) -> list[CandidateScore]:
    scored = [CandidateScore(value=candidate, score=similarity_score(value, candidate)) for candidate in candidates]
    scored.sort(key=lambda item: item.score, reverse=True)
    return scored


def similarity_score(value: str, candidate: str) -> float:
    return pair_similarity_score(clean_match_text(value), clean_match_text(candidate))


def pair_similarity_score(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0

    # 字符顺序相似度：关注字符排列顺序是否接近。
    sequence = SequenceMatcher(None, left, right).ratio()
    # 单字 Jaccard：关注两个文本的单字重合比例。
    char_jaccard = jaccard_similarity(char_ngrams(left, 1), char_ngrams(right, 1))
    # 双字片段 Jaccard：关注连续两个字符片段的重合比例。
    bigram_jaccard = jaccard_similarity(char_ngrams(left, 2), char_ngrams(right, 2))
    # 编辑距离相似度：关注需要多少次增删改才能互相转换。
    edit = edit_distance_similarity(left, right)
    # 包含关系相似度：处理简称和全称互相包含的情况。
    contains = contains_similarity(left, right)

    weighted = (
        0.35 * sequence
        + 0.25 * char_jaccard
        + 0.20 * bigram_jaccard
        + 0.20 * edit
    )

    return max(weighted, sequence, edit, contains)


def clean_match_text(value: str) -> str:
    text = str(value).strip().lower()
    for char in [" ", "　", "\t", "\n", "\r", "（", "）", "(", ")", "-", "_"]:
        text = text.replace(char, "")
    return text


def char_ngrams(text: str, n: int) -> set[str]:
    if not text:
        return set()
    if len(text) <= n:
        return {text}
    return {text[index : index + n] for index in range(0, len(text) - n + 1)}


def jaccard_similarity(left: set[str] | str, right: set[str] | str) -> float:
    left_set = set(left)
    right_set = set(right)
    union = left_set | right_set
    if not union:
        return 0.0
    return len(left_set & right_set) / len(union)


def edit_distance_similarity(left: str, right: str) -> float:
    max_len = max(len(left), len(right))
    if max_len == 0:
        return 1.0
    return max(0.0, 1.0 - levenshtein_distance(left, right) / max_len)


def levenshtein_distance(left: str, right: str) -> int:
    if left == right:
        return 0
    if not left:
        return len(right)
    if not right:
        return len(left)

    previous = list(range(len(right) + 1))
    for left_index, left_char in enumerate(left, 1):
        current = [left_index]
        for right_index, right_char in enumerate(right, 1):
            insert_cost = current[right_index - 1] + 1
            delete_cost = previous[right_index] + 1
            replace_cost = previous[right_index - 1] + (left_char != right_char)
            current.append(min(insert_cost, delete_cost, replace_cost))
        previous = current

    return previous[-1]


def contains_similarity(left: str, right: str) -> float:
    if left in right or right in left:
        return min(len(left), len(right)) / max(len(left), len(right))
    return 0.0


def _find_rule(field_name: str, rules: list[dict[str, Any]]) -> dict[str, Any] | None:
    text = str(field_name)
    for rule in rules:
        keywords = [str(item) for item in rule.get("field_keywords", [])]
        if any(keyword and keyword in text for keyword in keywords):
            return rule
    return None


def _result(
    status: str,
    original: str,
    output: str,
    message: str,
    score: float | None = None,
    score_gap: float | None = None,
) -> dict[str, Any]:
    return {
        "status": status,
        "original_value": original,
        "output_value": output,
        "message": message,
        "score": score,
        "score_gap": score_gap,
    }
