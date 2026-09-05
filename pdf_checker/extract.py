# -*- coding: utf-8 -*-
"""PDF 打印内容提取模块。

原理：只读取 PDF 的文本层（打印/印刷内容），手写内容通常是扫描图像，
不会被提取出来 —— 天然符合“手写的位置不重要，主要看打印内容”的需求。

使用 PyMuPDF 按 span（文本片段）粒度提取，保留坐标，便于后续按
“标签右侧 / 标签下方”定位字段值。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List

import pymupdf  # PyMuPDF（>=1.26 推荐 import pymupdf）


@dataclass
class Span:
    """一个文本片段（同字体、同样式的一段连续文字）。"""

    page: int          # 页码（0 起）
    x0: float
    y0: float
    x1: float
    y1: float
    text: str

    @property
    def x_center(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def y_center(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class Line:
    """同一基线的一行文字，由多个 Span 组成。"""

    page: int
    x0: float
    y0: float
    x1: float
    y1: float
    spans: List[Span] = field(default_factory=list)

    @property
    def text(self) -> str:
        """整行拼接文本（span 之间无分隔符）。"""
        return "".join(s.text for s in self.spans)

    @property
    def y_center(self) -> float:
        return (self.y0 + self.y1) / 2


def extract_lines(pdf_path: str) -> List[Line]:
    """提取 PDF 全部文本行（含坐标）。"""
    doc = pymupdf.open(pdf_path)
    lines: List[Line] = []
    try:
        for pno in range(doc.page_count):
            page = doc[pno]
            data = page.get_text("dict")
            for block in data.get("blocks", []):
                if block.get("type") != 0:
                    continue  # 只处理文本块，跳过图片
                for raw_line in block.get("lines", []):
                    spans: List[Span] = []
                    for raw_span in raw_line.get("spans", []):
                        text = raw_span.get("text", "")
                        if not text:
                            continue
                        bbox = raw_span["bbox"]
                        spans.append(
                            Span(
                                page=pno,
                                x0=bbox[0],
                                y0=bbox[1],
                                x1=bbox[2],
                                y1=bbox[3],
                                text=text,
                            )
                        )
                    if not spans:
                        continue
                    bbox = raw_line["bbox"]
                    lines.append(
                        Line(
                            page=pno,
                            x0=bbox[0],
                            y0=bbox[1],
                            x1=bbox[2],
                            y1=bbox[3],
                            spans=spans,
                        )
                    )
    finally:
        doc.close()
    return lines


def extract_full_text(pdf_path: str) -> str:
    """提取 PDF 全部文本（顺序拼接，用于正则兜底检索）。"""
    doc = pymupdf.open(pdf_path)
    try:
        return "\n".join(page.get_text() for page in doc)
    finally:
        doc.close()
