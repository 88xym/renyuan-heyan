# -*- coding: utf-8 -*-
"""校验编排：单个 PDF 的完整校验流程。

流程：提取文本行 → 按标签定位各字段值 → 逐字段执行规则 →
执行跨字段一致性检查 → 汇总结果。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import yaml

from .extract import Line, extract_full_text, extract_lines
from .locate import locate_field_values, regex_fallback
from .validators import run_cross_check, run_single_rule


@dataclass
class RuleResult:
    rule_type: str
    passed: bool
    message: str


@dataclass
class FieldResult:
    label: str
    alias: str
    value: Optional[str]
    found: bool
    required: bool
    rules: List[RuleResult] = field(default_factory=list)
    passed: bool = True


@dataclass
class CrossCheckResult:
    name: str
    passed: bool
    message: str


@dataclass
class FileResult:
    file_name: str
    overall: bool
    fields: List[FieldResult] = field(default_factory=list)
    cross_checks: List[CrossCheckResult] = field(default_factory=list)

    @property
    def issues(self) -> List[str]:
        out: List[str] = []
        for f in self.fields:
            if not f.passed:
                for r in f.rules:
                    if not r.passed:
                        out.append(f"[{f.label}] {r.message}")
                if f.found and not f.value:
                    out.append(f"[{f.label}] 未填写")
        for c in self.cross_checks:
            if not c.passed:
                out.append(f"[跨字段] {c.message}")
        return out


def load_rules(rules_path: str) -> Dict[str, Any]:
    with open(rules_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def validate_pdf(
    pdf_path: str,
    rules: Dict[str, Any],
    ocr_lines: Optional[List[Line]] = None,
    ocr_text: str = "",
) -> FileResult:
    """校验单个 PDF。

    ocr_lines / ocr_text：扫描件 OCR 结果（仅当 PDF 文本层为空时使用，
    作为 locate 定位与正则兜底的文本源），普通 PDF 不受影响。
    """
    lines = extract_lines(pdf_path)
    full_text = extract_full_text(pdf_path)
    if ocr_lines and not lines:
        lines = ocr_lines
    if ocr_text and not full_text.strip():
        full_text = ocr_text

    fields_cfg: List[Dict[str, Any]] = rules.get("fields", [])
    all_labels = [str(f.get("label", "")) for f in fields_cfg if f.get("label")]

    values: Dict[str, str] = {}
    field_results: List[FieldResult] = []

    for cfg in fields_cfg:
        label = str(cfg.get("label", ""))
        alias = str(cfg.get("alias", label))
        required = bool(cfg.get("required", False))
        locate = str(cfg.get("locate", "right"))

        candidates = locate_field_values(lines, label, all_labels, locate)
        value: Optional[str] = None
        found = bool(candidates)
        if candidates:
            value = candidates[0] or None
        if value is None:
            # 兜底正则
            value = regex_fallback(full_text, label)

        values[alias] = value or ""

        fr = FieldResult(
            label=label,
            alias=alias,
            value=value,
            found=found,
            required=required,
        )

        # 必填检查
        if required and not value:
            fr.rules.append(RuleResult("required", False, f"必填字段[{label}]未填写或未找到"))
            fr.passed = False

        # 格式规则
        for rule in cfg.get("rules", []):
            if not value:
                fr.rules.append(RuleResult(rule.get("type", ""), False, f"[{label}] 未填写，跳过格式校验"))
                fr.passed = False
                break
            ok, message = run_single_rule(rule, value)
            fr.rules.append(RuleResult(rule.get("type", ""), ok, message))
            if not ok:
                fr.passed = False

        field_results.append(fr)

    # 跨字段一致性
    cross_results: List[CrossCheckResult] = []
    for check in rules.get("cross_checks", []):
        ok, message = run_cross_check(check, values)
        cross_results.append(CrossCheckResult(check.get("name", ""), ok, message))

    overall = all(f.passed for f in field_results) and all(c.passed for c in cross_results)

    return FileResult(
        file_name=pdf_path.split("\\")[-1].split("/")[-1],
        overall=overall,
        fields=field_results,
        cross_checks=cross_results,
    )
