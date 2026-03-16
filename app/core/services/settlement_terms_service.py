from __future__ import annotations

from datetime import datetime
from typing import Any

import re

SETTLEMENT_TYPES = ["费比", "固定金额", "不计费", "按实际"]
SETTLEMENT_TERM_TEMPLATE_FIELDS = [
    "是否启用",
    "客户编码",
    "客户名称",
    "结算类型",
    "费比",
    "固定金额",
    "不计费原因",
    "备注",
    "生效月份",
    "失效月份",
]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    return text


def normalize_month(value: Any) -> str:
    if value is None:
        return ""
    if hasattr(value, "year") and hasattr(value, "month"):
        try:
            year = int(value.year)
            month = int(value.month)
            if 1 <= month <= 12:
                return f"{year:04d}-{month:02d}"
            return ""
        except Exception:
            return ""

    text = normalize_text(value)
    if text == "":
        return ""

    m = re.match(r"^(\d{4})(\d{2})$", text)
    if m:
        year = int(m.group(1))
        month = int(m.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
        return ""

    m = re.match(r"^(\d{4})[-/](\d{1,2})$", text)
    if m:
        year = int(m.group(1))
        month = int(m.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
        return ""

    try:
        dt = datetime.fromisoformat(text.replace("/", "-"))
        if 1 <= dt.month <= 12:
            return f"{dt.year:04d}-{dt.month:02d}"
        return ""
    except Exception:
        return ""


def to_optional_float(value: Any) -> float | None:
    text = normalize_text(value).replace(",", "")
    if text.endswith("%"):
        text = text[:-1].strip()
    if text == "":
        return None
    try:
        return float(text)
    except Exception:
        return None


def normalize_ratio_value(value: Any) -> str:
    numeric = to_optional_float(value)
    if numeric is None:
        return ""
    return f"{numeric:g}%"


def to_template_row(payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "是否启用": normalize_text(payload.get("enabled", "Y")) or "Y",
        "客户编码": normalize_text(payload.get("customer_code", "")),
        "客户名称": normalize_text(payload.get("customer_name", "")),
        "结算类型": normalize_text(payload.get("settle_type", "")),
        "费比": normalize_text(payload.get("ratio", "")),
        "固定金额": normalize_text(payload.get("fixed_amount", "")),
        "不计费原因": normalize_text(payload.get("no_charge_reason", "")),
        "备注": normalize_text(payload.get("remark", "")),
        "生效月份": normalize_month(payload.get("start_month", "")),
        "失效月份": normalize_month(payload.get("end_month", "")),
    }


def build_form_defaults(task: dict[str, Any]) -> dict[str, Any]:
    saved = task.get("maintain_payload") if isinstance(task.get("maintain_payload"), dict) else {}

    def get_saved(en_key: str, zh_key: str, default: Any = "") -> Any:
        if en_key in saved:
            return saved.get(en_key)
        if zh_key in saved:
            return saved.get(zh_key)
        return default

    return {
        "enabled": normalize_text(get_saved("enabled", "是否启用", "Y")) or "Y",
        "customer_code": normalize_text(get_saved("customer_code", "客户编码", task.get("customer_code", ""))),
        "customer_name": normalize_text(get_saved("customer_name", "客户名称", task.get("customer_name", ""))),
        "settle_type": normalize_text(get_saved("settle_type", "结算类型", "按实际")) or "按实际",
        "ratio": normalize_text(get_saved("ratio", "费比", "")),
        "fixed_amount": normalize_text(get_saved("fixed_amount", "固定金额", "")),
        "no_charge_reason": normalize_text(get_saved("no_charge_reason", "不计费原因", "")),
        "remark": normalize_text(get_saved("remark", "备注", "")),
        "start_month": normalize_month(get_saved("start_month", "生效月份", "")),
        "end_month": normalize_month(get_saved("end_month", "失效月份", "")),
    }


def build_payload(form_values: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    payload = {
        "enabled": normalize_text(form_values.get("enabled", "Y")) or "Y",
        "customer_code": normalize_text(form_values.get("customer_code", "")),
        "customer_name": normalize_text(form_values.get("customer_name", "")),
        "settle_type": normalize_text(form_values.get("settle_type", "")),
        "ratio": normalize_text(form_values.get("ratio", "")),
        "fixed_amount": normalize_text(form_values.get("fixed_amount", "")),
        "no_charge_reason": normalize_text(form_values.get("no_charge_reason", "")),
        "remark": normalize_text(form_values.get("remark", "")),
        "start_month": normalize_month(form_values.get("start_month", "")),
        "end_month": normalize_month(form_values.get("end_month", "")),
    }

    errors: list[str] = []
    if payload["customer_code"] == "" or payload["customer_code"] == "[缺失客户编码]":
        errors.append("客户编码")
    if payload["settle_type"] not in SETTLEMENT_TYPES:
        errors.append("结算类型")
    raw_start_month = normalize_text(form_values.get("start_month", ""))
    raw_end_month = normalize_text(form_values.get("end_month", ""))
    if payload["start_month"] == "":
        if raw_start_month == "":
            errors.append("生效月份")
        else:
            errors.append("生效月份格式错误")
    if raw_end_month != "" and payload["end_month"] == "":
        errors.append("失效月份格式错误")

    settle_type = payload["settle_type"]
    if settle_type == "费比":
        if to_optional_float(payload["ratio"]) is None:
            errors.append("费比")
        else:
            payload["ratio"] = normalize_ratio_value(payload["ratio"])
        payload["fixed_amount"] = ""
        payload["no_charge_reason"] = ""
    elif settle_type == "固定金额":
        if to_optional_float(payload["fixed_amount"]) is None:
            errors.append("固定金额")
        payload["ratio"] = ""
        payload["no_charge_reason"] = ""
    elif settle_type == "不计费":
        if payload["no_charge_reason"] == "":
            errors.append("不计费原因")
        payload["ratio"] = ""
        payload["fixed_amount"] = ""
    elif settle_type == "按实际":
        payload["ratio"] = ""
        payload["fixed_amount"] = ""

    return payload, errors
