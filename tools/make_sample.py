# -*- coding: utf-8 -*-
"""生成两个样例 PDF（一份填写正确、一份有错误）到 inputfile/，用于体验与测试。

用法：python tools/make_sample.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pymupdf  # PyMuPDF（>=1.26 推荐 import pymupdf）

PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = PROJECT_ROOT / "inputfile"

ID_WEIGHTS = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
ID_CHECK_MAP = "10X98765432"


def make_id_card(area: str, birth: str, seq: str) -> str:
    """按规则生成一个校验位正确的 18 位身份证号。"""
    body = area + birth + seq
    total = sum(int(body[i]) * ID_WEIGHTS[i] for i in range(17))
    return body + ID_CHECK_MAP[total % 11]


def build_form(path: Path, rows: list[tuple[str, str]]) -> None:
    """绘制一张简单的“人员核验表”（打印字段），内容按 rows 填写。"""
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)  # A4 纵向

    # 标题
    page.insert_text(
        (210, 60), "人员核验表", fontsize=18, fontname="china-s", color=(0, 0, 0)
    )

    left_x, mid_x, right_x = 80, 230, 480
    row_h = 50
    y = 120
    for label, value in rows:
        # 单元格边框
        page.draw_rect(pymupdf.Rect(left_x, y, right_x, y + row_h), color=(0, 0, 0), width=1)
        page.draw_rect(
            pymupdf.Rect(mid_x, y, right_x, y + row_h), color=(0, 0, 0), width=1
        )
        # 标签（左单元格，居中）
        page.insert_text(
            (left_x + (mid_x - left_x) / 2 - len(label) * 3.5, y + row_h / 2 + 4),
            label,
            fontsize=12,
            fontname="china-s",
            color=(0, 0, 0),
        )
        # 值（右单元格，左侧对齐，留白模拟填写位置）
        if value:
            page.insert_text(
                (mid_x + 12, y + row_h / 2 + 4),
                value,
                fontsize=12,
                fontname="china-s",
                color=(0, 0, 0),
            )
        y += row_h

    # 底部说明
    page.insert_text(
        (80, y + 40),
        "（样例文档，用于演示自动核验，数据均为虚构）",
        fontsize=10,
        fontname="china-s",
        color=(0.4, 0.4, 0.4),
    )

    doc.save(path)
    doc.close()
    print(f"已生成: {path}")


def main() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    valid_id = make_id_card("440402", "19900101", "123")
    wrong_id = valid_id[:-1] + ("0" if valid_id[-1] != "0" else "1")  # 破坏校验位

    correct_rows = [
        ("姓名", "张三"),
        ("身份证号", valid_id),
        ("性别", "男"),
        ("联系电话", "13800138000"),
        ("出生日期", "1990-01-01"),
        ("户籍地址", "广东省珠海市香洲区梅华东路1号"),
    ]
    wrong_rows = [
        ("姓名", "张"),                      # 姓名只有1个字
        ("身份证号", wrong_id),               # 校验位错误
        ("性别", "无"),                       # 性别非法
        ("联系电话", "1380013800"),           # 手机号少1位
        ("出生日期", "1989-12-31"),           # 与身份证出生日期不一致
        ("户籍地址", "广东省珠海市香洲区梅华东路1号"),
    ]

    build_form(INPUT_DIR / "样例_填写正确.pdf", correct_rows)
    build_form(INPUT_DIR / "样例_填写错误.pdf", wrong_rows)


if __name__ == "__main__":
    sys.exit(main())
