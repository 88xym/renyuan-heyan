# -*- coding: utf-8 -*-
"""字段定位模块：在提取出的文本行里，按“标签”找到字段值。

支持两种定位方式：
- right（默认）：值在标签同一行的右侧（常见于“姓名：张三”或表格
  中“标签单元格 | 值单元格”的布局）；
- below：值在标签的下一行（常见于上标签下值的布局）。

同时做了三项容错：
1. 若标签和值在同一 span（如“姓名：张三”），按字符串切分；
2. 若值单元格后续又出现其它标签（如“姓名 张三 性别 男”同一行），
   会在下一个已知标签处截断；
3. 找不到精确标签时，用整页文本正则兜底。
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional

from .extract import Line


# 值两侧常见的噪音字符（冒号、空格、标点、竖线等）
_LEADING_NOISE = re.compile(r"^[\s:：,，。;；、|｜_\-—]+")
_TRAILING_NOISE = re.compile(r"[\s,，。;；、|｜_\-—]+$")


def _clean(value: str) -> str:
    return _TRAILING_NOISE.sub("", _LEADING_NOISE.sub("", value))


def _line_contains_label(line: Line, label: str) -> bool:
    return label in line.text


def _value_from_span_split(line: Line, label: str) -> Optional[str]:
    """情形1：标签和值在同一个 span / 同一行文本里（如“姓名：张三”）。

    返回标签之后、下一个已知标签之前的部分；若标签后没有内容返回 None。
    """
    text = line.text
    idx = text.find(label)
    if idx < 0:
        return None
    rest = text[idx + len(label):]
    return _clean(rest) or None


def _value_from_right_spans(line: Line, label: str, all_labels: List[str]) -> Optional[str]:
    """情形2：标签是一个独立 span，值在同行的右侧 span 中（表格布局）。

    取标签 span 之后的所有同行 span 拼接，遇到其它已知标签即截断。
    """
    label_span_end: Optional[float] = None
    for span in line.spans:
        if label in span.text:
            label_span_end = span.x1
            break
    if label_span_end is None:
        return None

    parts: List[str] = []
    for span in line.spans:
        if span.x0 < label_span_end:
            continue
        # 该 span 若本身是其它标签，截断
        if any(other in span.text for other in all_labels if other != label):
            break
        parts.append(span.text)
    return _clean("".join(parts)) or None


def _value_from_same_baseline_right(
    lines: List[Line], line: Line, label: str, all_labels: List[str], tolerance: float = 10.0
) -> Optional[str]:
    """情形3：值在“同一基线、标签行右侧”的其它文本行里（跨 block 拆行）。

    按 x 排序拼接；遇其它已知标签截断。
    """
    yc = line.y_center
    candidates = [
        ln
        for ln in lines
        if ln.page == line.page
        and abs(ln.y_center - yc) <= tolerance
        and ln.x0 >= line.x1 - 2
        and ln is not line
    ]
    candidates.sort(key=lambda ln: ln.x0)
    parts: List[str] = []
    for ln in candidates:
        if any(other in ln.text for other in all_labels if other != label):
            break
        parts.append(ln.text)
    return _clean("".join(parts)) or None


def _value_from_below(
    lines: List[Line], line: Line, label: str, all_labels: List[str], x_tolerance: float = 80.0
) -> Optional[str]:
    """情形4：值在标签下方（上标签下值布局）。取标签下第一条非空、非标签的行。"""
    below = [
        ln
        for ln in lines
        if ln.page == line.page
        and ln.y0 > line.y1
        and abs(ln.x_center - line.x_center) <= x_tolerance
    ]
    below.sort(key=lambda ln: (ln.y0, ln.x0))
    for ln in below:
        text = _clean(ln.text)
        if not text:
            continue
        if any(other in text for other in all_labels):
            continue
        return text
    return None


def locate_field_values(
    lines: List[Line],
    label: str,
    all_labels: List[str],
    locate: str = "right",
) -> List[Optional[str]]:
    """按标签定位字段值，返回全部命中结果（通常取第一个）。

    locate: right（同行右侧，默认）| below（下一行）
    """
    label_lines = [ln for ln in lines if _line_contains_label(ln, label)]
    if not label_lines:
        return []

    results: List[Optional[str]] = []
    for ln in label_lines:
        # 1. 同行 span 内 / 同行文本内切分
        value = _value_from_span_split(ln, label)
        if value:
            results.append(value)
            continue
        # 2. 同行右侧 span（表格布局，标签独立 span）
        value = _value_from_right_spans(ln, label, all_labels)
        if value:
            results.append(value)
            continue
        # 3. 同一基线右侧其它行
        value = _value_from_same_baseline_right(lines, ln, label, all_labels)
        if value:
            results.append(value)
            continue
        # 4. 下方
        if locate == "below":
            value = _value_from_below(lines, ln, label, all_labels)
            if value:
                results.append(value)
                continue
        results.append(None)
    return results


def regex_fallback(full_text: str, label: str) -> Optional[str]:
    """兜底：直接在整页文本里用正则找“标签后跟值”。

    注意不跨行匹配（[ \\t:：]*），避免标签后空值时误取下一行内容。
    """
    pattern = re.compile(re.escape(label) + r"[ \t:：]*([^\s:：,，。;；、|｜]{1,60})")
    m = pattern.search(full_text)
    if not m:
        return None
    return _clean(m.group(1)) or None
