# -*- coding: utf-8 -*-
"""人员入场资料核验报告 PDF 生成器（跨平台版）。

每人一页 A4，包含：基本信息、支撑页面齐全性、名字/公司/日期一致性、
问题清单、结论。输出到 output/ 目录，文件名：<公司简称>-<日期>.pdf

跨平台说明：
- 字体自动探测：Windows(微软雅黑) → Linux(文鼎/文泉驿) → macOS(PingFang/黑体)
- 路径全部用 pathlib，无硬编码绝对路径
- 日期默认取当天，可通过 --date 覆盖
- 核验数据支持从外部 JSON 读取（--data），不依赖硬编码

依赖：reportlab（pip install reportlab）
用法：
    python tools/generate_report_pdf.py
    python tools/generate_report_pdf.py --data report_data.json
    python tools/generate_report_pdf.py --company 中交天航 --date 2026-09-06
"""

from __future__ import annotations

import argparse
import datetime
import json
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# ============================================================================
# 跨平台中文字体探测
# ============================================================================
# 注意：reportlab 只支持 TrueType outlines，不支持 OTF/CFF。
# 因此 Linux 上的 Noto Sans CJK（OTF）不可用，需用文鼎/文泉驿（TrueType）。
_FONT_CANDIDATES = [
    # Windows
    (r"C:\Windows\Fonts\msyh.ttc", 0, 1),       # 微软雅黑 常规/粗体
    (r"C:\Windows\Fonts\simsun.ttc", 0, 0),      # 宋体（无独立粗体，用同一索引）
    (r"C:\Windows\Fonts\simhei.ttf", None, None), # 黑体
    # Linux (Debian/Ubuntu)
    ("/usr/share/fonts/truetype/arphic/uming.ttc", 0, 0),   # 文鼎明体
    ("/usr/share/fonts/truetype/arphic/ukai.ttc", 0, 0),    # 文鼎楷体
    ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0, 0), # 文泉驿正黑
    ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 0, 0),
    # Linux (Fedora/RHEL)
    ("/usr/share/fonts/wqy-zenhei/wqy-zenhei.ttc", 0, 0),
    # macOS
    ("/System/Library/Fonts/PingFang.ttc", 0, 0),
    ("/System/Library/Fonts/STHeiti Light.ttc", 0, 0),
    ("/Library/Fonts/Arial Unicode.ttf", None, None),
]


def register_chinese_font() -> tuple[str, str]:
    """自动探测并注册可用的中文字体，返回 (常规字体名, 粗体字体名)。"""
    for path, regular_idx, bold_idx in _FONT_CANDIDATES:
        if not Path(path).exists():
            continue
        try:
            if regular_idx is not None:
                pdfmetrics.registerFont(TTFont("CJK", path, subfontIndex=regular_idx))
            else:
                pdfmetrics.registerFont(TTFont("CJK", path))
            # 粗体：优先用独立粗体索引，否则回退到常规字体
            if bold_idx is not None and bold_idx != regular_idx:
                try:
                    pdfmetrics.registerFont(TTFont("CJK-Bold", path, subfontIndex=bold_idx))
                except Exception:
                    pdfmetrics.registerFont(TTFont("CJK-Bold", path, subfontIndex=regular_idx))
            else:
                pdfmetrics.registerFont(TTFont("CJK-Bold", path, subfontIndex=regular_idx or 0))
            return ("CJK", "CJK-Bold")
        except Exception:
            continue
    raise RuntimeError(
        "未找到可用的中文字体。请安装以下任一字体：\n"
        "  Windows: 微软雅黑(msyh.ttc)、宋体(simsun.ttc)\n"
        "  Linux:   sudo apt install fonts-arphic-uming 或 fonts-wqy-zenhei\n"
        "  macOS:   系统自带 PingFang / 黑体"
    )


FONT_REGULAR, FONT_BOLD = register_chinese_font()

# ============================================================================
# 样式
# ============================================================================
styles = getSampleStyleSheet()
TITLE = ParagraphStyle("title", parent=styles["Title"], fontName=FONT_BOLD,
                       fontSize=18, leading=24, alignment=1, spaceAfter=6)
SUBTITLE = ParagraphStyle("subtitle", parent=styles["Normal"], fontName=FONT_REGULAR,
                          fontSize=10, leading=14, alignment=1, textColor=colors.grey, spaceAfter=10)
H2 = ParagraphStyle("h2", parent=styles["Heading2"], fontName=FONT_BOLD,
                    fontSize=12, leading=16, spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#1a5276"))
BODY = ParagraphStyle("body", parent=styles["Normal"], fontName=FONT_REGULAR,
                      fontSize=10, leading=15)
BODY_SMALL = ParagraphStyle("body_small", parent=styles["Normal"], fontName=FONT_REGULAR,
                            fontSize=9, leading=13)
ISSUE = ParagraphStyle("issue", parent=styles["Normal"], fontName=FONT_REGULAR,
                       fontSize=10, leading=15, leftIndent=12, bulletIndent=0)

# ============================================================================
# 默认核验数据（当不指定 --data 时使用）
# ============================================================================
DEFAULT_DATA = {
    "company_short": "中交天航",
    "source_pdf": "20260905153024.pdf",
    "people": [
        {
            "name": "陈昌洪", "id_type": "澳门永久性居民身份证", "id_no": "1668447(4)",
            "pages": "第2-7页", "page_count": 6,
            "pages_check": "齐全（个人资料/申请人声明/承判商声明/合同/证件复印件/身份证真伪记录）",
            "consistency": "名字、公司、身份证号、职安卡编号及有效期全部与证件复印件一致",
            "issues": [
                "承判商声明（第4页）：驻工地负责人签署、资料核对员签署，两处日期栏均为空",
                "个人资料页（第2页）：紧急联络人姓名、电话均未填写",
            ],
            "conclusion": "需补正",
        },
        {
            "name": "李永魁", "id_type": "特别逗留证", "id_no": "110520/2026",
            "pages": "第8-12页", "page_count": 5,
            "pages_check": "缺合同页（个人资料/申请人声明/承判商声明/证件复印件齐全；无身份证真伪记录，特别逗留证属正常）",
            "consistency": "名字、公司、特别逗留证号、职安卡编号及有效期全部与证件复印件一致",
            "issues": [
                "承判商声明（第10页）：两处签署日期栏均为空",
                "个人资料页（第8页）：紧急联络人姓名、电话均未填写",
                "缺少合同页：需确认特别逗留证人员是否免附合同",
            ],
            "conclusion": "需补正",
        },
        {
            "name": "梁智科", "id_type": "外地雇员身份识别证（蓝卡）", "id_no": "26188904",
            "pages": "第13-21页", "page_count": 9,
            "pages_check": "缺合同页（含劳工局批示、治安警察局申请表、逗留许可等官方文件；无身份证真伪记录，蓝卡属正常）",
            "consistency": "名字、公司、蓝卡号、职安卡编号及有效期、聘用许可有效期全部与证件复印件一致",
            "issues": [
                "承判商声明（第15页）：两处签署日期栏均为空",
                "缺少合同页：需确认外地雇员是否免附合同",
            ],
            "conclusion": "需补正",
        },
        {
            "name": "梁锦昌", "id_type": "澳门永久性居民身份证", "id_no": "1224375(4)",
            "pages": "第22-27页", "page_count": 6,
            "pages_check": "齐全（个人资料/申请人声明/承判商声明/合同/证件复印件/身份证真伪记录）",
            "consistency": "名字、公司、身份证号、职安卡编号及有效期、身份证真伪记录全部与证件复印件一致",
            "issues": [
                "承判商声明（第24页）：两处签署日期栏均为空",
                "个人资料页（第22页）：紧急联络人姓名、电话均未填写",
                "申请人声明（第23页）：签署日期栏完全空白",
            ],
            "conclusion": "需补正",
        },
        {
            "name": "韦将", "id_type": "外地雇员身份识别证（蓝卡）", "id_no": "26863514",
            "pages": "第28-37页", "page_count": 10,
            "pages_check": "齐全（含劳工局批示、治安警察局申请表、逗留许可等官方文件；无身份证真伪记录，蓝卡属正常）",
            "consistency": "名字、公司、蓝卡号、职安卡编号及有效期全部与证件复印件一致；申请人声明日期清晰（2026.7.30）",
            "issues": [
                "承判商声明（第30页）：两处签署日期栏均为空",
            ],
            "conclusion": "需补正",
        },
    ],
}


# ============================================================================
# 生成单页
# ============================================================================
def build_person_page(p: dict, idx: int, total: int, company_short: str,
                      source_pdf: str, report_date: str) -> list:
    story = []
    story.append(Paragraph("人员入场资料核验报告", TITLE))
    story.append(Paragraph(
        f"一包：{company_short}　|　来源文件：{source_pdf}　|　报告日期：{report_date}", SUBTITLE))

    info_data = [
        ["姓名", p["name"], "证件类型", p["id_type"]],
        ["证件号码", p["id_no"], "支撑页面", p["pages"]],
        ["资料页数", f"{p['page_count']}页", "核验结论", p["conclusion"]],
    ]
    info_table = Table(info_data, colWidths=[22*mm, 55*mm, 22*mm, 65*mm])
    info_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), FONT_REGULAR),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eaf2f8")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#eaf2f8")),
        ("FONTNAME", (0, 0), (0, -1), FONT_BOLD),
        ("FONTNAME", (2, 0), (2, -1), FONT_BOLD),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("TEXTCOLOR", (3, 2), (3, 2), colors.HexColor("#c0392b")),
        ("FONTNAME", (3, 2), (3, 2), FONT_BOLD),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 6))

    story.append(Paragraph("一、支撑页面齐全性", H2))
    story.append(Paragraph(p["pages_check"], BODY))
    story.append(Spacer(1, 4))

    story.append(Paragraph("二、名字 / 公司 / 日期与证件复印件一致性", H2))
    story.append(Paragraph(p["consistency"], BODY))
    story.append(Spacer(1, 4))

    story.append(Paragraph("三、需补正问题清单", H2))
    if p["issues"]:
        for i, issue in enumerate(p["issues"], 1):
            story.append(Paragraph(f"{i}. {issue}", ISSUE))
    else:
        story.append(Paragraph("无", BODY))
    story.append(Spacer(1, 6))

    story.append(Paragraph("四、核验结论", H2))
    concl_data = [[f"该人员资料存在 {len(p['issues'])} 项需补正问题，请退回分包商补正后重新提交。"]]
    concl_table = Table(concl_data, colWidths=[164*mm])
    concl_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), FONT_BOLD),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#fdebd0")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#7e5109")),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#d4ac0d")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(concl_table)
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"第 {idx} / {total} 页", ParagraphStyle(
        "page", parent=BODY_SMALL, alignment=2, textColor=colors.grey)))
    return story


# ============================================================================
# 主函数
# ============================================================================
def main() -> int:
    parser = argparse.ArgumentParser(description="人员入场资料核验报告 PDF 生成器（跨平台）")
    parser.add_argument("--data", default=None, help="核验数据 JSON 文件路径（不指定则用内置默认数据）")
    parser.add_argument("--company", default=None, help="一包公司简称（覆盖数据文件中的值）")
    parser.add_argument("--source", default=None, help="来源 PDF 文件名（覆盖数据文件中的值）")
    parser.add_argument("--date", default=None, help="报告日期 YYYY-MM-DD（默认当天）")
    parser.add_argument("--outdir", default=None, help="输出目录（默认项目下 output/）")
    args = parser.parse_args()

    # 加载数据
    if args.data:
        data_path = Path(args.data)
        if not data_path.exists():
            print(f"[错误] 数据文件不存在: {data_path}", file=sys.stderr)
            return 1
        with open(data_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = DEFAULT_DATA

    company_short = args.company or data.get("company_short", "未命名")
    source_pdf = args.source or data.get("source_pdf", "")
    report_date = args.date or datetime.date.today().strftime("%Y-%m-%d")
    people = data.get("people", [])

    if not people:
        print("[错误] 核验数据为空", file=sys.stderr)
        return 1

    outdir = Path(args.outdir) if args.outdir else PROJECT_ROOT / "output"
    outdir.mkdir(parents=True, exist_ok=True)
    date_compact = report_date.replace("-", "")
    out_path = outdir / f"{company_short}-{date_compact}.pdf"

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title=f"{company_short} 人员入场资料核验报告",
        author="人员核验系统",
    )

    story = []
    total = len(people)
    for idx, p in enumerate(people, 1):
        story.extend(build_person_page(p, idx, total, company_short, source_pdf, report_date))
        if idx < total:
            story.append(PageBreak())

    doc.build(story)
    print(f"PDF 已生成: {out_path}")
    print(f"共 {total} 人，每人一页。")
    print(f"字体: {FONT_REGULAR} (自动探测)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
