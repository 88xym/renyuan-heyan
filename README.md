# 人员核验 PDF 自动校验

把待核验的 PDF 文件放进 `inputfile/`，运行一条命令，程序提取 PDF 的**打印内容**
（手写内容为扫描图像，不在提取范围内，符合"手写的位置不重要，主要看打印内容"），
逐字段对照 `config/rules.yaml` 中的规则检查"填得对不对"，输出 Excel + JSON + PDF 报告。

## 环境要求

- **Python**: 3.10 ~ 3.14（跨平台：Windows / macOS / Linux）
- **中文字体**: 系统需自带任一可用中文字体（自动探测，无需配置）
  - Windows: 微软雅黑 / 宋体 / 黑体（系统自带）
  - macOS: PingFang / 黑体（系统自带）
  - Linux: 需安装 `fonts-arphic-uming` 或 `fonts-wqy-zenhei`（`sudo apt install fonts-arphic-uming`）
- **首次运行 OCR**: 需联网下载 RapidOCR 的 ONNX 模型（约 15MB），之后可离线使用

## 快速开始

### Windows

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 运行

```bash
# 生成样例 PDF 体验流程
python tools/make_sample.py

# 处理 inputfile/ 下所有 PDF
python main.py

# 处理指定文件
python main.py inputfile/样例_填写错误.pdf
```

## 扫描件 OCR

如果 PDF 是扫描件（无文本层），先做 OCR：

```bash
python tools/ocr_pdf.py inputfile/xxx.pdf
# 结果缓存到 .cache/ocr/<文件名>/，含每页文本 + 带坐标 JSON
```

## PDF 核验报告生成

把核验结果生成为每人一页的 PDF 报告：

```bash
# 使用内置默认数据
python tools/generate_report_pdf.py

# 从外部 JSON 读取数据
python tools/generate_report_pdf.py --data report_data.json

# 指定公司简称和日期
python tools/generate_report_pdf.py --company 中交天航 --date 2026-09-06
```

输出：`output/<公司简称>-<YYYYMMDD>.pdf`，每人一页 A4。

数据 JSON 格式：
```json
{
  "company_short": "中交天航",
  "source_pdf": "20260905153024.pdf",
  "people": [
    {
      "name": "张三",
      "id_type": "澳门永久性居民身份证",
      "id_no": "1234567(8)",
      "pages": "第2-7页",
      "page_count": 6,
      "pages_check": "齐全（...）",
      "consistency": "名字、公司、日期全部一致",
      "issues": ["问题1", "问题2"],
      "conclusion": "需补正"
    }
  ]
}
```

## 目录结构

```
人员核验/
├── inputfile/              # 放入待核验的 PDF
├── output/                 # 核验报告（xlsx/json/pdf，自动生成）
├── .cache/ocr/             # OCR 缓存（自动生成，含个人信息，已 gitignore）
├── config/
│   ├── rules.yaml          # 字段校验规则
│   ├── check_policy.yaml   # 核验策略（手写体宽松/0-O不区分等）
│   └── organization.yaml   # 分包层级（总包→一包→二包）
├── pdf_checker/            # 核心代码
│   ├── extract.py          #   提取打印文本（含坐标）
│   ├── locate.py           #   按标签定位字段值
│   ├── validators.py       #   校验规则实现
│   ├── validate.py         #   校验编排
│   └── report.py           #   报告输出（Excel/JSON）
├── tools/
│   ├── make_sample.py      # 样例 PDF 生成器
│   ├── ocr_pdf.py          # 扫描件 OCR 工具
│   ├── scan_pages.py       # 页面人名扫描（定位支撑页面范围）
│   └── generate_report_pdf.py  # PDF 核验报告生成（跨平台）
├── .vscode/                # VSCode 配置（launch/settings）
├── requirements.txt
└── main.py                 # 入口
```

## 如何定制规则

打开 `config/rules.yaml`，按实际表单修改：

1. `fields`：每个字段写清楚 PDF 中的**标签文字**（如"姓名""身份证号"）、
   是否必填、取值位置（`right` = 标签同行右侧 / `below` = 标签下一行）。
2. `rules`：支持 `regex`（正则）、`in_set`（枚举）、`date`（日期）、
   `id_card`（身份证号）四种校验。
3. `cross_checks`：跨字段一致性（身份证出生日期/性别 ↔ 填写字段）。

> 当前 rules.yaml 中是示例规则，覆盖姓名/身份证/性别/电话/出生日期，
> 请以真实表单为准调整或增删。

## 核验策略

`config/check_policy.yaml` 控制核验宽松度：
- `handwriting: lenient` — 手写体宽松检查（OCR 异常不直接判错）
- `zero_vs_oh: ignore` — 不区分数字 0 和字母 O
- `unclear_signature: pass` — 签名看不清视为已签署
- `empty_field: fail` — 字段完全空白仍判为缺失

`config/organization.yaml` 定义分包层级，雇佣关系不同（如受雇于一包但实际是二包员工）不视为矛盾。

## 报告

- `output/核验报告_<时间戳>.xlsx`：汇总 + 逐字段明细（通过/不通过标色）
- `output/核验结果_<时间戳>.json`：结构化结果，便于程序处理
- `output/<公司简称>-<日期>.pdf`：每人一页的 PDF 核验报告
