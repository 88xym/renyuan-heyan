# -*- coding: utf-8 -*-
"""PDF 扫描件 OCR 工具：把每页渲染成图片，用 RapidOCR 识别打印内容，
结果（纯文本 + 带坐标 JSON）缓存到本地目录，供后续校验分析使用。

用法：python tools/ocr_pdf.py inputfile/xxx.pdf [--outdir .cache/ocr] [--dpi 250]
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import numpy as np
import pymupdf
from rapidocr_onnxruntime import RapidOCR

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def page_to_array(page: pymupdf.Page, dpi: int) -> np.ndarray:
    """把 PyMuPDF 页面渲染成 RGB numpy 数组（供 RapidOCR 使用）。"""
    zoom = dpi / 72
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csRGB, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    return arr[:, :, :3]


def ocr_pdf(pdf_path: Path, outdir: Path, dpi: int = 250) -> None:
    engine = RapidOCR()
    outdir.mkdir(parents=True, exist_ok=True)

    doc = pymupdf.open(pdf_path)
    pages_meta = []
    full_lines: list[str] = []

    for i in range(doc.page_count):
        page = doc[i]
        img = page_to_array(page, dpi)
        result, elapse = engine(img)

        lines: list[dict] = []
        if result:
            # result: [[box(4点), text, score], ...]
            for box, text, score in result:
                # 归一化坐标（相对页面，0~1），方便与 PDF 坐标互转
                x0 = min(p[0] for p in box) / img.shape[1]
                y0 = min(p[1] for p in box) / img.shape[0]
                x1 = max(p[0] for p in box) / img.shape[1]
                y1 = max(p[1] for p in box) / img.shape[0]
                lines.append({"text": text, "score": float(score),
                              "bbox": [round(x0, 5), round(y0, 5), round(x1, 5), round(y1, 5)]})
            lines.sort(key=lambda l: (l["bbox"][1], l["bbox"][0]))

        pages_meta.append({"page": i + 1, "lines": lines})
        page_text = "\n".join(l["text"] for l in lines)
        full_lines.append(f"===== 第 {i + 1} 页 =====\n{page_text}")
        (outdir / f"page_{i + 1:03d}.txt").write_text(page_text, encoding="utf-8")
        elapsed = sum(elapse) if isinstance(elapse, (list, tuple)) else elapse
        print(f"第 {i + 1}/{doc.page_count} 页 完成，识别 {len(lines)} 块（{elapsed:.1f}s）")

    (outdir / "full_text.txt").write_text("\n\n".join(full_lines), encoding="utf-8")
    (outdir / "ocr_result.json").write_text(
        json.dumps(pages_meta, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    doc.close()
    print(f"\n完成。结果目录: {outdir}")
    print(f"  - full_text.txt   全部页文本")
    print(f"  - page_NNN.txt    每页文本")
    print(f"  - ocr_result.json 带坐标的结构化结果")


def main() -> int:
    parser = argparse.ArgumentParser(description="PDF 扫描件 OCR 到本地缓存")
    parser.add_argument("pdf", help="PDF 文件路径")
    parser.add_argument("--outdir", default=None, help="缓存输出目录（默认 .cache/ocr/<文件名>）")
    parser.add_argument("--dpi", type=int, default=250, help="渲染分辨率（默认250）")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"[错误] 文件不存在: {pdf_path}", file=sys.stderr)
        return 1

    outdir = Path(args.outdir) if args.outdir else PROJECT_ROOT / ".cache" / "ocr" / pdf_path.stem
    t0 = time.time()
    ocr_pdf(pdf_path, outdir, args.dpi)
    print(f"总耗时: {time.time() - t0:.1f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
