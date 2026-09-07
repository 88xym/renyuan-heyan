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

> **Windows 一键运行**：双击项目根目录 `开始核验.bat` 即可。
> 流程：自动 OCR（扫描件无缓存时）→ main.py 字段级核验（output 下 Excel/JSON）→ auto_verify.py 生成每人一页 PDF 核验报告（output/`<公司简称>-<日期>.pdf`）。
>
> **一包公司简称自动识别**：从申请人个人资料页"承判公司名称"字段提取（如"长安保安服务有限公司"→`长安保安`、"中交天航南方交通建设有限公司"→`中交天航`），同步用于输出文件名与 PDF 标题；OCR 识别为噪声（不含公司特征词）时回退 organization.yaml 默认一包别名。可用 `--company 名称` 覆盖。
>
> **两种文件结构均支持**：①第1页为人员统计表（按表定位每人，如土金批次）；②无统计表、第1页直接为个人资料页（如联合体文件，按"申请人个人资料"页切分，姓名取资料页"（中文）XXX"字段）。
> 该 bat 为 GBK(ANSI)+CRLF 编码，与中文 Windows 的 cmd 兼容。
> 如需修改提示文字，请改 `tools/bat_source.bat`（UTF-8）后运行
> `python tools/rebuild_bat.py` 重建，不要直接编辑 `开始核验.bat`。

## 扫描件 OCR

**自动识别（推荐）**：`main.py` / `开始核验.bat` 运行时会自动检测扫描件：

- PDF 有文本层 → 直接核验，不 OCR；
- 扫描件且 `.cache/ocr/<文件名>/` 已有完整缓存（页数一致）→ 直接使用缓存，不重新 OCR；
- 扫描件且无缓存/缓存不完整 → 自动调用 OCR，完成后继续核验（耗时取决于页数，约 1 分钟/10 页）。

手动 OCR（可选）：

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
│   ├── rules.yaml              # 字段校验规则（示例）
│   ├── check_policy.yaml       # 核验策略（手写体宽松/0-O不区分等）
│   ├── organization.yaml       # 分包层级（总包→一包→二包）
│   └── verification_rules.yaml # 三类人员核验标准规则 v2.0（资料清单/合同逻辑/蓝卡特殊）
├── docs/
│   └── 人员核验标准.md          # 核验标准人工对照版（三类人员资料清单+检查项）
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

`config/verification_rules.yaml` 定义三类人员核验标准（v2.0）：
- **澳门本地人**：6项资料，含身份证真伪记录（必附）
- **特别逗留证人士**：5项资料，无真伪记录要求
- **蓝卡**：6项资料，含劳工三表一致性核验；蓝卡未颁发时行街纸+红印纸可替代
- **合同逻辑**：承判商声明单位与第一页不一致 → 补二判/三判合同；直属总包 → 免附合同

`docs/人员核验标准.md` 为人工对照版核验标准。

`config/organization.yaml` 定义分包层级，雇佣关系不同（如受雇于一包但实际是二包员工）不视为矛盾。

## 报告

- `output/核验报告_<时间戳>.xlsx`：汇总 + 逐字段明细（通过/不通过标色）
- `output/核验结果_<时间戳>.json`：结构化结果，便于程序处理
- `output/<公司简称>-<日期>.pdf`：每人一页的 PDF 核验报告
