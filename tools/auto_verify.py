# -*- coding: utf-8 -*-
"""自动核验引擎：按核验标准 v2.0 检查扫描件 PDF，生成每人一页 PDF 报告。

流程：
  1. 读取 OCR 缓存（.cache/ocr/<pdf_stem>/page_*.txt）
  2. 从第 1 页人员统计表自动提取人员名单
  3. 定位每人支撑页面范围
  4. 按页面标题特征分类（个人资料/申请人声明/承判商声明/合同/证件/真伪记录/劳工三表）
  5. 按核验标准检查字段问题（紧急联络人/签署日期/批示编号等）
  6. 生成 report data JSON
  7. 调用 tools/generate_report_pdf.py 输出每人一页 PDF

用法：
  python tools/auto_verify.py                          # 自动处理 inputfile 最新 PDF
  python tools/auto_verify.py --pdf inputfile/xxx.pdf  # 指定 PDF
  python tools/auto_verify.py --names 陈昌洪,李永魁    # 手动指定名单（可选）
  python tools/auto_verify.py --company 中交天航       # 覆盖公司简称

依赖：本脚本读取 OCR 缓存，未 OCR 的扫描件请先运行：
  python tools/ocr_pdf.py inputfile/xxx.pdf
"""

from __future__ import annotations

import argparse
import datetime
import json
import re
import subprocess
import sys
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CACHE_ROOT = PROJECT_ROOT / ".cache" / "ocr"
OUTPUT_DIR = PROJECT_ROOT / "output"
ORGANIZATION_CFG = PROJECT_ROOT / "config" / "organization.yaml"

# 常见姓氏（用于从第 1 页统计表提取人员姓名）
_SURNAMES = set(
    "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"
    "戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐"
    "费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄"
    "和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝董梁"
    "杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高夏蔡田胡凌霍虞"
    "万支柯昝管卢莫经房裘缪干解应宗丁宣贲邓郁单杭洪包诸左石崔吉钮龚程"
    "嵇邢滑裴陆荣翁荀羊於惠甄麹家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧"
    "隗山谷车侯宓蓬全郗班仰秋仲伊宫宁仇栾暴甘斜厉戎祖武符刘景詹束龙叶"
    "幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴胥能苍双闻莘"
    "党翟谭贡劳逄姬申扶堵冉宰郦雍郤璩桑桂濮牛寿通边扈燕冀郏浦尚农温别"
    "庄晏柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡国"
    "文寇广禄阙东欧殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空曾毋"
    "沙乜养鞠须丰巢关蒯相查后荆红游竺权逯盖益桓公"
)

# 常用繁体字 → 简体字（用于 OCR 文本归一化与姓名繁简匹配）
_TC2SC = {
    "發": "发", "華": "华", "曉": "晓", "龍": "龙", "鬆": "松", "國": "国",
    "門": "门", "號": "号", "證": "证", "張": "张", "陳": "陈", "楊": "杨",
    "黃": "黄", "劉": "刘", "吳": "吴", "鄭": "郑", "謝": "谢", "羅": "罗",
    "蘇": "苏", "葉": "叶", "馬": "马", "許": "许", "郭": "郭", "譚": "谭",
    "範": "范", "賴": "赖", "鍾": "钟", "萬": "万", "龔": "龚", "歐": "欧",
    "鄒": "邹", "蔣": "蒋", "湯": "汤", "馮": "冯", "藍": "蓝", "韓": "韩",
    "麥": "麦", "韋": "韦", "鄧": "邓", "聯": "联", "絡": "络", "資": "资",
    "號": "号", "聲": "声", "簽": "签", "稱": "称", "項": "项", "確": "确",
    "認": "认", "實": "实", "務": "务", "圖": "图", "關": "关", "動": "动",
    "購": "购", "單": "单", "員": "员", "勞": "劳", "處": "处", "後": "后",
    "專": "专", "業": "业", "興": "兴", "灣": "湾", "計": "计", "個": "个",
    "複": "复", "記": "记", "錄": "录", "書": "书", "員": "员", "頻": "频",
    "總": "总", "匯": "汇", "類": "类", "別": "别", "電": "电", "話": "话",
    "現": "现", "居": "居", "證": "证", "複": "复", "閱": "阅", "讀": "读",
    "爲": "为", "長": "长", "開": "开", "閘": "闸", "動": "动", "遞": "递",
    "縉": "缙", "鏟": "铲",
}


def norm_text(text: str) -> str:
    """繁转简归一化（覆盖本项目出现的常用繁体字），用于关键词匹配。"""
    for tc, sc in _TC2SC.items():
        if tc in text:
            text = text.replace(tc, sc)
    return text


def name_variants(name: str) -> set[str]:
    """生成姓名的繁简变体集合（如 翟發六 → {翟發六, 翟发六}）。"""
    variants = {""}
    for ch in name:
        sc = _TC2SC.get(ch, ch)
        variants = {v + ch for v in variants} | {v + sc for v in variants}
    return variants if len(name) == len(next(iter(variants))) else {name}

# 页面标题 → 类别（按顺序匹配，先命中的优先级更高）
# 注意：页面文本会先做繁简归一化（norm_text），关键词一律用规范简体。
_PAGE_CLASSIFIERS = [
    # 注意：表单页脚"开卡个人申请表 1/3"带页码，用负向断言排除
    ("personal_info", [r"申请人个人[资料餐料]", r"开卡个人申请表(?!\s*[0-9０-９]/\d)"]),
    ("applicant_decl", ["申请人声明", "申人声明"]),
    ("contractor_decl", ["承判商声明"]),
    ("id_authenticity", ["真伪记录", "查核智能身份证"]),
    ("contract", ["合约", "合同"]),
    ("labor_police_form", ["逗留许可.*申请表", "申请表编号", "申睛表"]),
    ("labor_receipt", ["收入收", "收据", "行街纸"]),
    ("labor_bureau", ["劳工事务局", "DSAL", "批示"]),
    ("id_copy", ["通行证", "职安", "外地雇员身份", "行街纸", "真伪记录"]),
]

_ID_TYPE_RULES = [
    ("macau_resident", ["永久性居民身份"]),
    ("special_stay", ["特别逗留"]),
    ("blue_card", ["外地雇员身份", "外地雇员身份"]),
]

# 六类资料名称（报告展示用）
_DOC_NAMES = {
    "personal_info": "申请人个人资料页",
    "applicant_decl": "申请人声明",
    "contractor_decl": "承判商声明",
    "contract": "合同页",
    "id_copy": "证件复印件页",
    "id_authenticity": "身份证真伪记录",
    "labor_bureau": "劳工局函件/批示",
    "labor_police_form": "治安警察局申请表",
    "labor_receipt": "行街纸收据+逗留许可(红印纸)",
}


# ============================================================================
# 工具函数
# ============================================================================
def load_ocr_pages(pdf_stem: str) -> list[tuple[int, str]]:
    """读取 OCR 缓存，返回 [(页码, 文本), ...] 按页号排序。"""
    ocr_dir = CACHE_ROOT / pdf_stem
    if not ocr_dir.exists():
        return []
    pages = []
    for f in sorted(ocr_dir.glob("page_*.txt")):
        m = re.search(r"page_(\d+)", f.stem)
        if m:
            pages.append((int(m.group(1)), f.read_text(encoding="utf-8")))
    return pages


def extract_names(page1_text: str) -> list[str]:
    """从第 1 页人员统计表提取人员姓名（宽松：含常见姓氏的 2-4 字中文行）。"""
    names = []
    for line in page1_text.splitlines():
        line = line.strip()
        if not line or len(line) < 2 or len(line) > 4:
            continue
        if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", line) and line[0] in _SURNAMES:
            if line not in names:
                names.append(line)
    return names


def person_name_variants(name: str) -> set[str]:
    """生成姓名的繁简变体集合（如 翟發六 → {翟發六, 翟发六}），供页面定位匹配。"""
    return name_variants(name)


def is_name_on_pages(name: str, pages: list[tuple[int, str]], skip_pages: set[int] = {1}) -> bool:
    """判断姓名是否出现在指定页（排除第 1 页统计表，避免表头/联络人/跟进人误判）。"""
    variants = name_variants(name)
    for n, t in pages:
        if n in skip_pages:
            continue
        norm = norm_text(t)
        if any(v in norm for v in variants):
            return True
    return False


def classify_page(text: str) -> set[str]:
    """按标题特征给页面打分类标签（文本先做繁简归一化）。

    只匹配页面顶部区域（前 800 字）：表单标题都在页面上部，
    避免底部页脚（如"开卡个人申请表 1/3"）造成误分类。
    """
    cats = set()
    norm = norm_text(text)[:800]
    for cat, keywords in _PAGE_CLASSIFIERS:
        for kw in keywords:
            if re.search(kw, norm):
                cats.add(cat)
                break
    return cats


def detect_id_type(text: str) -> str | None:
    """从页面文本判定证件类型（澳门本地/特别逗留/蓝卡）。"""
    for id_type, keywords in _ID_TYPE_RULES:
        if any(kw in text for kw in keywords):
            return id_type
    return None


def _clean(s: str) -> str:
    """去掉非字母数字，用于宽松比对。"""
    return re.sub(r"[^0-9A-Za-z]", "", s.upper())


def line_has_content(text: str, label: str) -> tuple[bool, str]:
    """检查 'label' 同行是否有填写内容。返回 (是否已填, 同行内容)。"""
    m = re.search(label + r"[:：]?\s*([^\n]*)", text)
    if not m:
        return False, ""
    content = m.group(1).strip()
    # 去掉常见模板残留
    content = re.sub(r"^[:：\s]*", "", content)
    return bool(content), content


def find_line_after(text: str, anchor: str, label: str) -> tuple[bool, str]:
    """在 anchor 之后查找 label 行，返回 (是否已填, 同行内容)。"""
    idx = text.find(anchor)
    if idx < 0:
        return False, ""
    seg = text[idx:]
    m = re.search(label + r"[:：]?\s*([^\n]*)", seg)
    if not m:
        return False, ""
    content = m.group(1).strip()
    return bool(content), content


# ============================================================================
# 单人核验
# ============================================================================
def check_person(name: str, person_pages: list[tuple[int, str]], names_all: list[str]) -> dict:
    """对一个人执行完整核验，返回报告数据结构。"""
    # 统一繁简归一化，后续全部关键词/正则匹配基于规范简体
    person_pages = [(n, norm_text(t)) for n, t in person_pages]
    pages_text = "".join(t for _, t in person_pages)
    page_nums = [n for n, _ in person_pages]

    # 页面分类统计
    cat_pages: dict[str, list[int]] = {}
    for n, t in person_pages:
        for cat in classify_page(t):
            cat_pages.setdefault(cat, []).append(n)
    # 全文档身份证类型判定（优先个人资料页）
    id_type = None
    for n, t in person_pages:
        dt = detect_id_type(t)
        if dt:
            id_type = dt
            break
    id_type = id_type or "unknown"
    # 类型推断：OCR 未识别出证件类型字样，但同时有劳工局文件 + 治安警察局
    # 逗留许可申请表（行街纸/红印纸）→ 判为外地雇员（蓝卡）类
    if id_type == "unknown" and cat_pages.get("labor_police_form") and cat_pages.get("labor_bureau"):
        id_type = "blue_card"

    # ---- 资料齐全性 ----
    has = {k: bool(cat_pages.get(k)) for k in _DOC_NAMES}
    missing_docs, notes = [], []
    if id_type == "macau_resident":
        required = ["personal_info", "applicant_decl", "contractor_decl",
                    "contract", "id_copy", "id_authenticity"]
    elif id_type == "special_stay":
        required = ["personal_info", "applicant_decl", "contractor_decl", "id_copy"]
    elif id_type == "blue_card":
        required = ["personal_info", "applicant_decl", "contractor_decl",
                    "id_copy", "labor_bureau", "labor_police_form", "labor_receipt"]
    else:
        required = ["personal_info", "applicant_decl", "contractor_decl", "id_copy"]

    for k in required:
        if not has.get(k):
            missing_docs.append(_DOC_NAMES[k])

    # 合同逻辑：承判商声明中是否出现二判/三判分判公司（富华等）
    contractor_text = "".join(t for n, t in person_pages if "contractor_decl" in classify_page(t))
    has_subcontractor = ("富華" in contractor_text) or ("富华" in contractor_text)
    if id_type in ("macau_resident", "special_stay", "blue_card"):
        if has.get("contractor_decl"):
            if has_subcontractor:
                if not has.get("contract"):
                    missing_docs.append("合同页（承判商声明填有二判/三判单位，需附一判-二判合同）")
                else:
                    notes.append("承判商声明填有二判/三判单位，合同页已附（一判-二判合约）")
            else:
                notes.append("承判商声明未填其他单位，属直属承判关系，免附合同")
        # 承判商声明缺失时，由 required 列表报告缺失，不附加"直属/免合同"说明

    # ---- 字段问题检查 ----
    issues: list[str] = []
    # 1. 紧急联络人（只报"标签存在但同行空白"，OCR 丢行/识别不清不误报）
    p_info_text = "".join(t for n, t in person_pages if "personal_info" in classify_page(t))
    if p_info_text:
        for label, hint in (("紧急联络人姓名", "紧急联络人姓名"), ("紧急联络电话", "紧急联络电话")):
            m = re.search(label, p_info_text)
            if m:
                line = p_info_text[m.start():].split("\n", 1)[0]
                rest = line[m.end():].strip(":： \t")
                if not rest:
                    issues.append(f"第{cat_pages['personal_info'][0]}页·个人资料页：{hint}空白")

    # 2. 申请人声明签署日期
    # 表单布局："签署日期：xxx" 在上、"申请人签署：xxx" 在下，直接搜"签署日期"字段
    decl_text = "".join(t for n, t in person_pages if "applicant_decl" in classify_page(t))
    if decl_text:
        m = re.search(r"签署日期[:：]?\s*([^\n]*)", decl_text)
        content = m.group(1).strip() if m else ""
        filled = bool(content)
        decl_page = cat_pages.get("applicant_decl", ["?"])[0]
        if not filled:
            issues.append(f"第{decl_page}页·申请人声明：签署日期完全空白（无任何书写痕迹）")
        # 手写日期做宽松识别：有内容即视为已签（不苛求 20xx 格式）

    # 3. 承判商声明两处签署日期
    if contractor_text:
        filled, _ = find_line_after(contractor_text, "承判商驻工地负责人", "日期")
        c_page = cat_pages.get("contractor_decl", ["?"])[0]
        if not filled:
            issues.append(f"第{c_page}页·承判商声明：承判商驻工地负责人签署日期空白")
        filled, _ = find_line_after(contractor_text, "资料核对员签署", "日期")
        if not filled:
            issues.append(f"第{c_page}页·承判商声明：资料核对员签署日期空白")

    # 4. 批示编号一致性（蓝卡：申请表 vs 劳工局批示）
    if id_type == "blue_card":
        form_text = "".join(t for n, t in person_pages if "labor_police_form" in classify_page(t))
        bureau_text = "".join(t for n, t in person_pages if "labor_bureau" in classify_page(t))
        m_form = re.search(r"(\d{4,5})\s*[/／]\s*IMO\s*[/／]\s*DSAL\s*[/／]\s*20\d{2}", form_text)
        m_bureau = re.search(r"第?\s*(\d{4,5})\s*[/／]\s*IMO\s*[/／]\s*DSAL\s*[/／]\s*20\d{2}", bureau_text)
        if m_form and m_bureau and m_form.group(1) != m_bureau.group(1):
            f_page = cat_pages.get("labor_police_form", ["?"])[0]
            b_page = cat_pages.get("labor_bureau", ["?"])[0]
            issues.append(
                f"第{f_page}页·治安警察局申请表：聘用许可批示编号({m_form.group(0)})"
                f"与劳工局批示编号({m_bureau.group(0)})不一致，需人工核实是否属不同批示"
            )

    # ---- 一致性（宽松比对）----
    id_no, card_no, company, chinese_name = "", "", "", ""
    if p_info_text:
        m = re.search(r"身份证号码[:：]*\s*([0-9A-Za-z（）()/／-]+)", p_info_text)
        id_no = m.group(1).strip() if m else ""
        m = re.search(r"职安[咭卡]编号（绿）[:：]*\s*([0-9A-Za-z/／-]+)", p_info_text)
        card_no = m.group(1).strip() if m else ""
        m = re.search(r"承判公司名称[:：]*\s*([^\n]+)", p_info_text)
        company = m.group(1).strip() if m else ""
        m = re.search(r"（中文）([\u4e00-\u9fa5]+)", p_info_text)
        chinese_name = m.group(1) if m else ""

    id_copy_text = "".join(t for n, t in person_pages if "id_copy" in classify_page(t))
    id_ok = (not id_no) or (_clean(id_no) in _clean(id_copy_text)) or (_clean(id_no) in _clean(pages_text))
    card_ok = (not card_no) or (_clean(card_no) in _clean(id_copy_text))
    name_ok = (not chinese_name) or (chinese_name in pages_text)
    if id_ok and card_ok and name_ok:
        consistency = "姓名、证件号码、职安卡编号已与证件复印件自动比对一致（OCR 宽松识别）"
    else:
        detail = []
        if not id_ok:
            detail.append("证件号码未在证件页找到")
        if not card_ok:
            detail.append("职安卡编号未在证件页找到")
        if not name_ok:
            detail.append("中文姓名比对异常")
        consistency = "需人工复核：" + "；".join(detail)

    # ---- 组装报告 ----
    if id_type == "macau_resident":
        type_name = "澳门本地人（澳门身份证）"
    elif id_type == "special_stay":
        type_name = "特别逗留证人士"
    elif id_type == "blue_card":
        type_name = "外地雇员身份识别证（蓝卡）"
    else:
        type_name = "未知"

    docs_desc = []
    for k in required:
        status = "✓" if has.get(k) else "✗"
        doc_name = _DOC_NAMES[k]
        pages = cat_pages.get(k)
        pg = f"第{'-'.join(map(str, pages))}页" if pages else ""
        docs_desc.append(f"{doc_name}({status}){pg}")
    if notes:
        docs_desc.append("；".join(notes))
    pages_check = "，".join(docs_desc)
    if missing_docs:
        pages_check += "。缺失：" + "、".join(missing_docs)

    conclusion = "需补正" if (issues or missing_docs) else "通过"
    return {
        "name": name,
        "id_type": type_name,
        "id_no": id_no or "未识别",
        "pages": f"第{page_nums[0]}-{page_nums[-1]}页" if page_nums else "",
        "page_count": len(person_pages),
        "pages_check": pages_check,
        "consistency": consistency,
        "issues": issues,
        "conclusion": conclusion,
    }


# ============================================================================
# 主流程
# ============================================================================
def main() -> int:
    parser = argparse.ArgumentParser(description="自动核验引擎：生成每人一页 PDF 核验报告")
    parser.add_argument("--pdf", default=None, help="PDF 路径（默认 inputfile 下第一个 PDF）")
    parser.add_argument("--names", default=None, help="人员名单（逗号分隔，可选；默认从第1页自动提取）")
    parser.add_argument("--company", default=None, help="一包公司简称（默认从 organization.yaml 读取）")
    parser.add_argument("--date", default=None, help="报告日期 YYYY-MM-DD（默认当天）")
    args = parser.parse_args()

    # 定位 PDF
    if args.pdf:
        pdf_path = Path(args.pdf)
    else:
        pdfs = sorted(PROJECT_ROOT.glob("inputfile/*.pdf")) + sorted(PROJECT_ROOT.glob("inputfile/*.PDF"))
        if not pdfs:
            print("[错误] inputfile 下未找到 PDF", file=sys.stderr)
            return 1
        pdf_path = pdfs[0]
    if not pdf_path.exists():
        print(f"[错误] PDF 不存在: {pdf_path}", file=sys.stderr)
        return 1
    pdf_stem = pdf_path.stem

    # 读取 OCR 缓存
    pages = load_ocr_pages(pdf_stem)
    if not pages:
        print(f"[提示] 未找到 OCR 缓存: {CACHE_ROOT / pdf_stem}")
        print(f"       请先运行: python tools/ocr_pdf.py \"{pdf_path}\"")
        return 1
    print(f"OCR 缓存已加载: {len(pages)} 页")

    # 提取名单
    if args.names:
        names = [n.strip() for n in args.names.split(",") if n.strip()]
    else:
        # 从第 1 页统计表提取名单，并剔除只在第 1 页出现（表头/联络人/跟进人等）的误提取
        names = [n for n in extract_names(pages[0][1]) if is_name_on_pages(n, pages)]
    if not names:
        print("[错误] 无法从第1页提取人员名单，请用 --names 指定", file=sys.stderr)
        return 1
    print(f"人员名单: {', '.join(names)}")

    # 逐人定位页面并核验
    # 方案1：按"个人资料页"位置切分（每个人以个人资料页开头，顺序对应总汇表名单）；
    # 方案2（回退）：按姓名首次出现页切分。
    people = []
    ranges = None
    person_starts = [n for n, t in pages[1:] if "personal_info" in classify_page(t)]
    if len(person_starts) == len(names):
        ranges = []
        for i, start in enumerate(person_starts):
            end = (person_starts[i + 1] - 1) if i + 1 < len(person_starts) else pages[-1][0]
            ranges.append((start, end))
    if ranges is None:
        first_occur = []
        for name in names:
            fp = next(
                (n for n, t in pages[1:] if any(v in norm_text(t) for v in name_variants(name))),
                None,
            )
            if fp is not None:
                first_occur.append((name, fp))
        if len(first_occur) == len(names):
            first_occur.sort(key=lambda x: x[1])
            ranges = []
            start = 2
            for i, (name, fp) in enumerate(first_occur):
                end = (first_occur[i + 1][1] - 1) if i + 1 < len(first_occur) else pages[-1][0]
                ranges.append((start, end))
                start = end + 1
    if ranges is None:
        print("[错误] 无法按名单切分页面，请检查第1页名单与页面结构", file=sys.stderr)
        return 1
    for name, (s, e) in zip(names, ranges):
        person_pages = [(n, t) for n, t in pages if s <= n <= e]
        if not person_pages:
            print(f"  [提示] 未找到 {name} 的页面，跳过")
            continue
        people.append(check_person(name, person_pages, names))
        print(f"  已核验 {name}: 第{s}-{e}页, {len(people[-1]['issues'])} 项问题")

    if not people:
        print("[错误] 无人通过核验", file=sys.stderr)
        return 1

    # 公司简称
    company = args.company
    if not company:
        try:
            org = yaml.safe_load(ORGANIZATION_CFG.read_text(encoding="utf-8"))
            company = org["hierarchy"][1]["alias"][0]
        except Exception:
            company = "未命名"
    report_date = args.date or datetime.date.today().strftime("%Y-%m-%d")
    date_compact = report_date.replace("-", "")

    # 输出数据 JSON
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    data = {
        "company_short": company,
        "source_pdf": pdf_path.name,
        "people": people,
    }
    data_json = OUTPUT_DIR / f"核验数据_{date_compact}.json"
    data_json.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"核验数据: {data_json}")

    # 调用 PDF 报告生成器
    script = PROJECT_ROOT / "tools" / "generate_report_pdf.py"
    cmd = [sys.executable, str(script), "--data", str(data_json),
           "--company", company, "--date", report_date]
    try:
        result = subprocess.run(cmd, cwd=str(PROJECT_ROOT))
        return result.returncode
    except Exception as exc:  # noqa: BLE001
        print(f"[错误] 调用报告生成器失败: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
