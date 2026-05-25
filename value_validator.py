from __future__ import annotations

from datetime import date, datetime
import re
from typing import Any


WARNING_EMPTY_VALUES = {"无"}
SPECIAL_EMPTY_VALUES = {"", "/", "暂无", "未提供", "不涉及"}
STRUCTURED_TYPES = {"positive_integer", "number", "number_with_unit", "list", "date", "id_card", "phone", "age"}
DATE_ALLOWED_TEXT_VALUES = {"\u81f3\u4eca"}


def clean_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        text = value.strftime("%Y-%m-%d")
    elif isinstance(value, date):
        text = value.strftime("%Y-%m-%d")
    else:
        text = str(value)

    for char in [" ", "　", "\t", "\n", "\r"]:
        text = text.replace(char, "")
    while text.endswith("。") or text.endswith("."):
        text = text[:-1]
    return text.replace(";", "；").strip()


def result(status: str, original: str, output: str, message: str) -> dict[str, str]:
    return {
        "status": status,
        "original_value": original,
        "output_value": output,
        "message": message,
    }


def parse_date_value(value: Any, granularity: str = "ymd") -> datetime | None:
    if isinstance(value, datetime):
        if granularity == "y" and (value.month != 1 or value.day != 1):
            return None
        if granularity == "ym" and value.day != 1:
            return None
        return value
    if isinstance(value, date):
        if granularity == "y" and (value.month != 1 or value.day != 1):
            return None
        if granularity == "ym" and value.day != 1:
            return None
        return datetime(value.year, value.month, value.day)

    text = str(value).strip()
    for pattern, fmt in _date_parse_formats(granularity):
        if not re.fullmatch(pattern, text):
            continue
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def validate_positive_integer(value: str) -> bool:
    return re.fullmatch(r"[1-9]\d*", value) is not None


def validate_number(value: str) -> bool:
    try:
        return float(value) >= 0
    except ValueError:
        return False


def validate_number_with_unit(value: str, unit: str) -> bool:
    return re.fullmatch(rf"[0-9]+(?:\.[0-9]+)?{re.escape(unit)}", value) is not None


def validate_id_card(value: str) -> bool:
    return re.fullmatch(r"\d{17}[\dXx]", value) is not None


def validate_phone(value: str) -> bool:
    return re.fullmatch(r"1[3-9]\d{9}", value) is not None


def normalize_age_value(value: str) -> str:
    if value.endswith("\u5c81"):
        return value[:-1]
    return value


def validate_by_rule(value: Any, rule: dict[str, Any]) -> dict[str, str]:
    original = "" if value is None else str(value).strip()
    cleaned = clean_value(value)
    rule_type = rule.get("type", "string")
    changed_by_clean = original != cleaned

    if rule_type == "any":
        if changed_by_clean:
            return result("NORMALIZED", original, cleaned, "基础清洗后写回")
        return result("PASS", original, cleaned, "不校验字段")

    if cleaned in SPECIAL_EMPTY_VALUES:
        return result("PASS", original, cleaned, "特殊空值，通过")

    if cleaned in WARNING_EMPTY_VALUES:
        if rule_type in STRUCTURED_TYPES:
            return result("WARNING", original, cleaned, f"{rule_type} 字段填写“无”，请人工确认")
        return result("PASS", original, cleaned, "普通字段填写“无”")

    if rule_type == "id_card":
        normalized = cleaned[:-1] + cleaned[-1].upper() if cleaned else cleaned
        if validate_id_card(normalized):
            if original != normalized:
                return result("NORMALIZED", original, normalized, "身份证号格式规范化")
            return result("PASS", original, normalized, "身份证号通过")
        return result("ERROR", original, "/", "身份证号应为18位，前17位数字，最后一位数字或X")

    if rule_type == "phone":
        if validate_phone(cleaned):
            if changed_by_clean:
                return result("NORMALIZED", original, cleaned, "手机号删除空格后通过")
            return result("PASS", original, cleaned, "手机号通过")
        return result("ERROR", original, "/", "手机号应为11位，并符合手机号格式")

    if rule_type == "age":
        normalized = normalize_age_value(cleaned)
        if validate_positive_integer(normalized):
            if original != normalized:
                return result("NORMALIZED", original, normalized, "年龄字段已去除“岁”并统一为正整数")
            return result("PASS", original, normalized, "年龄通过")
        return result("ERROR", original, "/", "年龄应为正整数，可填写如 20 或 20岁")

    if rule_type == "enum":
        allowed_values = rule.get("allowed_values", [])
        if cleaned in allowed_values:
            if changed_by_clean:
                return result("NORMALIZED", original, cleaned, "枚举值清洗后通过")
            return result("PASS", original, cleaned, "枚举值通过")
        return result("ERROR", original, "/", f"枚举值错误，允许值：{'/'.join(allowed_values)}")

    if rule_type == "positive_integer":
        if validate_positive_integer(cleaned):
            if changed_by_clean:
                return result("NORMALIZED", original, cleaned, "正整数清洗后通过")
            return result("PASS", original, cleaned, "正整数通过")
        return result("ERROR", original, "/", "应为正整数")

    if rule_type == "number":
        if validate_number(cleaned):
            if changed_by_clean:
                return result("NORMALIZED", original, cleaned, "数值清洗后通过")
            return result("PASS", original, cleaned, "数值通过")
        return result("ERROR", original, "/", "应为非负数字")

    if rule_type == "number_with_unit":
        unit = rule.get("unit", "")
        if validate_number_with_unit(cleaned, unit):
            if changed_by_clean:
                return result("NORMALIZED", original, cleaned, "带单位数值清洗后通过")
            return result("PASS", original, cleaned, "带单位数值通过")
        return result("ERROR", original, "/", f"应为数字+单位，单位必须为：{unit}")

    if rule_type == "date":
        granularity = rule.get("date_granularity", "ymd")
        if cleaned in DATE_ALLOWED_TEXT_VALUES and granularity in {"ymd", "ym"}:
            if original != cleaned:
                return result("NORMALIZED", original, cleaned, "日期字段文本值已清洗后通过")
            return result("PASS", original, cleaned, "日期字段允许填写“至今”")
        parsed = parse_date_value(cleaned, granularity)
        if parsed:
            default_format = "yyyy" if granularity == "y" else "yyyy/mm" if granularity == "ym" else "yyyy/mm/dd"
            output = format_date_value(parsed, rule.get("date_format", default_format))
            if original != output:
                return result("NORMALIZED", original, output, "日期格式统一为模板格式")
            return result("PASS", original, output, "日期通过")
        if granularity == "y":
            return result("ERROR", original, "/", "年份字段只能填写年份")
        if granularity == "ym":
            return result("ERROR", original, "/", "年月字段只能填写年月")
        return result("ERROR", original, "/", "日期字段必须填写完整年月日")

    if rule_type == "list":
        if cleaned in WARNING_EMPTY_VALUES:
            return result("WARNING", original, cleaned, "列表字段填写“无”，请人工确认")
        if "；" not in cleaned:
            item_rule = rule.get("item_rule", {"type": "string"})
            item_result = validate_by_rule(cleaned, item_rule)
            if item_result["status"] == "ERROR":
                return result("ERROR", original, "/", f"列表元素错误：{cleaned}，{item_result['message']}")
            if changed_by_clean:
                return result("NORMALIZED", original, cleaned, "列表单项值清洗后通过")
            return result("PASS", original, cleaned, "列表单项值通过")

        items = [item for item in cleaned.split("；") if item]
        if not items:
            return result("ERROR", original, "/", "列表不能为空")
        if any(item in WARNING_EMPTY_VALUES for item in items):
            return result("ERROR", original, "/", "列表中不能混入“无”")

        item_rule = rule.get("item_rule", {"type": "string"})
        for item in items:
            item_result = validate_by_rule(item, item_rule)
            if item_result["status"] == "ERROR":
                return result("ERROR", original, "/", f"列表元素错误：{item}，{item_result['message']}")

        output = "；".join(items)
        if original != output:
            return result("NORMALIZED", original, output, "列表分隔符或空格已规范化")
        return result("PASS", original, output, "列表通过")

    if rule_type == "string":
        if changed_by_clean:
            return result("NORMALIZED", original, cleaned, "字符串清洗后写回")
        return result("PASS", original, cleaned, "字符串通过")

    return result("PASS", original, cleaned, "默认通过")


def format_date_value(value: datetime, target_format: str) -> str:
    replacements = {
        "yyyy": f"{value.year:04d}",
        "mm": f"{value.month:02d}",
        "dd": f"{value.day:02d}",
        "m": str(value.month),
        "d": str(value.day),
    }
    output = target_format
    for token in ["yyyy", "mm", "dd", "m", "d"]:
        output = output.replace(token, replacements[token])
    return output


def _date_parse_formats(granularity: str) -> list[tuple[str, str]]:
    if granularity == "y":
        return [
            (r"\d{4}", "%Y"),
        ]
    if granularity == "ym":
        return [
            (r"\d{4}-\d{1,2}", "%Y-%m"),
            (r"\d{4}\.\d{1,2}", "%Y.%m"),
            (r"\d{4}/\d{1,2}", "%Y/%m"),
            (r"\d{6}", "%Y%m"),
            (r"\d{4}年\d{1,2}月", "%Y年%m月"),
        ]
    return [
        (r"\d{4}-\d{1,2}-\d{1,2}", "%Y-%m-%d"),
        (r"\d{4}\.\d{1,2}\.\d{1,2}", "%Y.%m.%d"),
        (r"\d{4}/\d{1,2}/\d{1,2}", "%Y/%m/%d"),
        (r"\d{8}", "%Y%m%d"),
        (r"\d{4}年\d{1,2}月\d{1,2}日", "%Y年%m月%d日"),
        (r"\d{4}-\d{1,2}-\d{1,2}\d{1,2}:\d{1,2}:\d{1,2}", "%Y-%m-%d%H:%M:%S"),
    ]
