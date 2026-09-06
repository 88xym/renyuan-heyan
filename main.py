# -*- coding: utf-8 -*-
"""人员核验 PDF 自动校验 - 入口。

用法：
    python main.py                       # 处理 inputfile/ 下所有 PDF
    python main.py inputfile/xxx.pdf     # 处理指定 PDF（可多个）
    python main.py --rules config/rules.yaml --outdir output

产物：
    output/核验报告_<时间戳>.xlsx   逐字段明细
    output/核验结果_<时间戳>.json   结构化结果
"""

from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

from pdf_checker.report import print_console, write_excel, write_json
from pdf_checker.validate import load_rules, validate_pdf

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_RULES = PROJECT_ROOT / "config" / "rules.yaml"
DEFAULT_INPUT = PROJECT_ROOT / "inputfile"
DEFAULT_OUTPUT = PROJECT_ROOT / "output"


def is_scanned_pdf(pdf_path: str, sample_pages: int = 3) -> bool:
    """快速检测 PDF 是否为扫描件（前 sample_pages 页文本层均为空）。"""
    try:
        import pymupdf
        doc = pymupdf.open(pdf_path)
        try:
            total = min(sample_pages, doc.page_count)
            for i in range(total):
                if doc[i].get_text().strip():
                    return False
            return True
        finally:
            doc.close()
    except Exception:
        return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description="人员核验 PDF 自动校验：提取打印内容并对照规则检查",
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="PDF 文件路径；不填则处理 inputfile/ 下全部 PDF",
    )
    parser.add_argument(
        "--rules", default=str(DEFAULT_RULES), help="校验规则配置文件（YAML）"
    )
    parser.add_argument(
        "--outdir", default=str(DEFAULT_OUTPUT), help="报告输出目录"
    )
    args = parser.parse_args()

    rules_path = Path(args.rules)
    if not rules_path.exists():
        print(f"[错误] 规则文件不存在: {rules_path}", file=sys.stderr)
        return 1

    rules = load_rules(str(rules_path))

    # 确定待处理文件
    if args.files:
        pdf_files = [Path(f) for f in args.files if Path(f).exists()]
        missing = [f for f in args.files if not Path(f).exists()]
        for m in missing:
            print(f"[警告] 文件不存在，跳过: {m}", file=sys.stderr)
    else:
        # 同时匹配 .pdf 和 .PDF（Linux/macOS 大小写敏感），并用 set 去重（Windows 不区分大小写会重复）
        pdf_files = sorted(
            set(list(DEFAULT_INPUT.glob("*.pdf")) + list(DEFAULT_INPUT.glob("*.PDF")))
        )

    if not pdf_files:
        print(f"[提示] 未找到 PDF 文件，请放入: {DEFAULT_INPUT}")
        print(f"       可先运行 python tools/make_sample.py 生成样例体验。")
        return 1

    results = []
    for pdf in pdf_files:
        print(f"正在核验: {pdf.name} ...")
        if is_scanned_pdf(str(pdf)):
            print(f"  [警告] {pdf.name} 疑似扫描件（无文本层），直接核验可能得到空结果。")
            print(f"         建议先运行: python tools/ocr_pdf.py \"{pdf}\"")
        try:
            results.append(validate_pdf(str(pdf), rules))
        except Exception as exc:  # noqa: BLE001
            print(f"[错误] 解析 {pdf.name} 失败: {exc}", file=sys.stderr)
            continue

    if not results:
        return 1

    print_console(results)

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_path = outdir / f"核验报告_{stamp}.xlsx"
    json_path = outdir / f"核验结果_{stamp}.json"
    write_excel(results, str(xlsx_path))
    write_json(results, str(json_path))
    print(f"Excel 报告: {xlsx_path}")
    print(f"JSON 结果:  {json_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
