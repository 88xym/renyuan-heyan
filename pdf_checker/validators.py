# -*- coding: utf-8 -*-
"""字段校验规则实现。

规则类型（rules[].type）：
- regex      正则匹配（pattern 用完整匹配）
- in_set     取值必须属于给定集合（values）
- date       合法日期（自动识别 YYYY-MM-DD / YYYYMMDD / YYYY年MM月DD日）
- id_card    中国大陆 18 位居民身份证号（含出生日期与校验位）

跨字段一致性（cross_checks[].type）：
- id_birth_match   身份证号中的出生日期 == 出生日期字段
- id_gender_match  身份证号第 17 位奇偶 与 性别字段一致
"""

from __future__ import annotations

import datetime
import re
from typing import Any, Dict, List, Optional, Tuple

ID_WEIGHTS = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
ID_CHECK_MAP = "10X98765432"


def normalize(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def check_regex(value: Any, pattern: str) -> bool:
    v = normalize(value)
    if not v:
        return False
    try:
        return re.fullmatch(pattern, v) is not None
    except re.error:
        return False


def check_in_set(value: Any, values: List[str]) -> bool:
    return normalize(value) in values


def check_id_card(value: Any) -> bool:
    """中国大陆居民身份证号校验：长度 18、出生日期合法、校验位正确。"""
    v = normalize(value).upper()
    if not re.fullmatch(r"\d{17}[\dX]", v):
        return False
    # 出生日期合法性
    try:
        datetime.date(int(v[6:10]), int(v[10:12]), int(v[12:14]))
    except ValueError:
        return False
    # 校验位
    total = sum(int(v[i]) * ID_WEIGHTS[i] for i in range(17))
    return ID_CHECK_MAP[total % 11] == v[17]


def id_card_birth(value: Any) -> Optional[str]:
    """从身份证号提取出生日期 YYYY-MM-DD；格式不合法返回 None。"""
    v = normalize(value).upper()
    if not re.fullmatch(r"\d{17}[\dX]", v):
        return None
    try:
        return f"{v[6:10]}-{v[10:12]}-{v[12:14]}"
    except Exception:
        return None


def id_card_gender(value: Any) -> Optional[str]:
    """从身份证号第 17 位推断性别；格式不合法返回 None。"""
    v = normalize(value).upper()
    if not re.fullmatch(r"\d{17}[\dX]", v):
        return None
    return "男" if int(v[16]) % 2 == 1 else "女"


_DATE_PATTERNS = [
    (re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$"), "%Y-%m-%d"),
    (re.compile(r"^(\d{4})/(\d{1,2})/(\d{1,2})$"), "%Y-%m-%d"),
    (re.compile(r"^(\d{4})年(\d{1,2})月(\d{1,2})日?$"), "%Y-%m-%d"),
    (re.compile(r"^(\d{8})$"), "%Y%m%d"),
]


def normalize_date(value: Any) -> Optional[str]:
    """把常见日期写法归一为 YYYY-MM-DD；无法解析返回 None。"""
    v = normalize(value)
    for pattern, _ in _DATE_PATTERNS:
        m = pattern.match(v)
        if m:
            try:
                dt = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                return dt.strftime("%Y-%m-%d")
            except ValueError:
                return None
    return None


def check_date(value: Any) -> bool:
    return normalize_date(value) is not None


def run_single_rule(rule: Dict[str, Any], value: Any) -> Tuple[bool, str]:
    """执行单条规则，返回 (是否通过, 说明)。"""
    rtype = rule.get("type", "")
    message = rule.get("message", "未通过校验")

    if rtype == "regex":
        ok = check_regex(value, rule.get("pattern", ""))
    elif rtype == "in_set":
        ok = check_in_set(value, rule.get("values", []))
    elif rtype == "date":
        ok = check_date(value)
    elif rtype == "id_card":
        ok = check_id_card(value)
    else:
        ok = False
        message = f"未知规则类型: {rtype}"
    return ok, message


# ---------------- 跨字段一致性 ----------------

def run_cross_check(check: Dict[str, Any], values: Dict[str, str]) -> Tuple[bool, str]:
    """执行一条跨字段检查；任一字段缺失则跳过（返回 True + 说明）。"""
    ctype = check.get("type", "")
    message = check.get("message", "跨字段一致性未通过")
    name = check.get("name", ctype)

    id_field = check.get("id_field", "id_card")
    birth_field = check.get("birth_field", "birth_date")
    gender_field = check.get("gender_field", "gender")

    if ctype == "id_birth_match":
        idv = values.get(id_field, "")
        bv = values.get(birth_field, "")
        if not idv or not bv:
            return True, f"[跳过] {name}：缺少身份证号或出生日期字段"
        id_birth = id_card_birth(idv)
        b_norm = normalize_date(bv)
        if id_birth is None or b_norm is None:
            return False, f"{name}：身份证号或出生日期格式无法解析"
        return id_birth == b_norm, f"{name}：身份证出生日期 {id_birth} vs 出生日期 {b_norm}"

    if ctype == "id_gender_match":
        idv = values.get(id_field, "")
        gv = values.get(gender_field, "")
        if not idv or not gv:
            return True, f"[跳过] {name}：缺少身份证号或性别字段"
        id_gender = id_card_gender(idv)
        if id_gender is None:
            return False, f"{name}：身份证号格式无法解析"
        return id_gender == normalize(gv), f"{name}：身份证推断 {id_gender} vs 填写 {gv}"

    return False, f"未知跨字段检查类型: {ctype}"
