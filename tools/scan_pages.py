# -*- coding: utf-8 -*-
"""扫描每页 OCR 文本中的人名，定位每个人对应的支撑页面范围。"""
from pathlib import Path

names = ["陈昌洪", "陳昌洪", "李永魁", "梁智科", "梁锦昌", "韋將", "韦将"]
pages = sorted(Path(".cache/ocr/20260905153024").glob("page_*.txt"))
for p in pages:
    text = p.read_text(encoding="utf-8")
    found = list({n for n in names if n in text})
    if found:
        print(f"{p.stem}: {' '.join(found)}")
    else:
        print(f"{p.stem}: (无匹配人名)")
