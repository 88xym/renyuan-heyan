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
import shutil
import subprocess
import sys
import time
from pathlib import Path

from pdf_checker.extract import Line, Span
from pdf_checker.report import print_console, write_excel, write_json
from pdf_checker.validate import load_rules, validate_pdf

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_RULES = PROJECT_ROOT / "config" / "rules.yaml"
DEFAULT_INPUT = PROJECT_ROOT / "inputfile"
DEFAULT_OUTPUT = PROJECT_ROOT / "output"
CACHE_DIR = PROJECT_ROOT / ".cache"
CACHE_WARN_MB = 500  # 缓存目录超过此大小（MB）时提示删除


def get_dir_size(path: Path) -> int:
    """递归计算目录总大小（字节），忽略无法读取的文件。"""
    total = 0
    if not path.exists():
        return 0
    for f in path.rglob("*"):
        if f.is_file():
            try:
                total += f.stat().st_size
            except OSError:
                pass
    return total


def cleanup_old_cache(days: int = 7) -> int:
    """删除指定天数前的 OCR 缓存子目录，返回删除的目录数。

    OCR 缓存按 PDF 文件名存放在 .cache/ocr/<pdf名>/ 下，
    按子目录的修改时间判断是否过期，过期则整个子目录删除。
    """
    ocr_dir = CACHE_DIR / "ocr"
    if not ocr_dir.exists():
        return 0
    cutoff = time.time() - days * 86400
    deleted = 0
    for subdir in sorted(ocr_dir.iterdir()):
        if not subdir.is_dir():
            continue
        try:
            if subdir.stat().st_mtime < cutoff:
                shutil.rmtree(subdir, ignore_errors=True)
                deleted += 1
                print(f"  已删除: {subdir.name}")
        except OSError as exc:
            print(f"  删除失败 {subdir.name}: {exc}")
    return deleted


def check_cache_size() -> None:
    """检查缓存目录大小，超过阈值时交互式询问是否删除 7 天前的缓存。"""
    if not CACHE_DIR.exists():
        return
    size_mb = get_dir_size(CACHE_DIR) / (1024 * 1024)
    if size_mb <= CACHE_WARN_MB:
        return
    print(f"[警告] 缓存目录 {CACHE_DIR} 已占用 {size_mb:.1f} MB，超过 {CACHE_WARN_MB} MB。")
    print(f"       OCR 缓存可安全删除，删除后下次运行对应 PDF 会重新 OCR。")
    try:
        choice = input("       是否删除 7 天前的缓存？(y/N): ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return
    if choice in ("y", "yes"):
        print("       正在清理 7 天前的缓存...")
        deleted = cleanup_old_cache(days=7)
        if deleted > 0:
            new_size = get_dir_size(CACHE_DIR) / (1024 * 1024)
            print(f"       清理完成：删除 {deleted} 个缓存目录，当前缓存 {new_size:.1f} MB。")
        else:
            print("       没有 7 天前的缓存可删除。")
    else:
        print("       已跳过缓存清理。")
    print()


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


def ocr_cache_complete(pdf_path: Path) -> bool:
    """检查 OCR 缓存是否完整（缓存页数 == PDF 页数）。"""
    ocr_dir = CACHE_DIR / "ocr" / pdf_path.stem
    if not ocr_dir.exists():
        return False
    pages = list(ocr_dir.glob("page_*.txt"))
    if not pages:
        return False
    try:
        import pymupdf
        doc = pymupdf.open(pdf_path)
        n = doc.page_count
        doc.close()
        return len(pages) >= n
    except Exception:
        return True  # 无法打开 PDF 时信任已有缓存


def run_ocr(pdf_path: Path) -> bool:
    """自动调用 tools/ocr_pdf.py 执行 OCR，成功返回 True。"""
    script = PROJECT_ROOT / "tools" / "ocr_pdf.py"
    print(f"  [OCR] 未找到完整缓存，自动执行 OCR: {pdf_path.name} ...")
    print(f"        请耐心等待，扫描件按页识别，耗时取决于页数...")
    try:
        result = subprocess.run(
            [sys.executable, str(script), str(pdf_path)],
            cwd=str(PROJECT_ROOT),
        )
        return result.returncode == 0
    except Exception as exc:  # noqa: BLE001
        print(f"  [OCR错误] {exc}", file=sys.stderr)
        return False


def load_ocr_text(pdf_path: Path) -> str:
    """读取 OCR 缓存全文（优先 full_text.txt，否则拼接 page_*.txt）。"""
    ocr_dir = CACHE_DIR / "ocr" / pdf_path.stem
    ft = ocr_dir / "full_text.txt"
    if ft.exists():
        return ft.read_text(encoding="utf-8")
    parts = []
    for f in sorted(ocr_dir.glob("page_*.txt")):
        parts.append(f.read_text(encoding="utf-8"))
    return "\n".join(parts)


def build_ocr_lines(pdf_path: Path) -> list[Line]:
    """把 OCR 缓存文本构造为伪坐标 Line 列表，供字段定位使用。

    OCR 文本没有 PDF 坐标，这里把每行作为一个独立 Line：
    - 同行"标签：值"由 span 内切分逻辑处理（标签右侧取内容）；
    - 行距拉大到 100，避免"同一基线右侧"误拼其它行；
    - below 定位对 OCR 行不适用（无真实上下坐标），依赖同行取值。
    """
    ocr_dir = CACHE_DIR / "ocr" / pdf_path.stem
    lines: list[Line] = []
    pno = 0
    y = 0.0
    for f in sorted(ocr_dir.glob("page_*.txt")):
        for text in f.read_text(encoding="utf-8").splitlines():
            if not text.strip():
                continue
            w = float(len(text))
            span = Span(page=pno, x0=0.0, y0=y, x1=w, y1=y, text=text)
            lines.append(Line(page=pno, x0=0.0, y0=y, x1=w, y1=y, spans=[span]))
            y += 100.0
        pno += 1
    return lines


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

    # 启动时检查缓存目录大小，超过阈值提示删除
    check_cache_size()

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
        ocr_lines: list[Line] | None = None
        ocr_text = ""
        if is_scanned_pdf(str(pdf)):
            # 扫描件：自动检查 OCR 缓存，无缓存/不完整则自动 OCR
            if not ocr_cache_complete(pdf):
                if not run_ocr(pdf):
                    print(f"  [警告] {pdf.name} OCR 失败，将按原始文本核验（结果可能为空）。")
                else:
                    print(f"  [OK] {pdf.name} OCR 完成，缓存: .cache/ocr/{pdf.stem}")
            else:
                print(f"  [缓存] {pdf.name} 已有完整 OCR 缓存，直接使用。")
            ocr_lines = build_ocr_lines(pdf)
            ocr_text = load_ocr_text(pdf)
        try:
            results.append(validate_pdf(str(pdf), rules, ocr_lines=ocr_lines, ocr_text=ocr_text))
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
