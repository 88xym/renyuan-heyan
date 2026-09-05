# 人员核验 PDF 自动校验

把待核验的 PDF 文件放进 `inputfile/`，运行一条命令，程序提取 PDF 的**打印内容**
（手写内容为扫描图像，不在提取范围内，符合“手写的位置不重要，主要看打印内容”），
逐字段对照 `config/rules.yaml` 中的规则检查“填得对不对”，输出 Excel + JSON 报告。

## 快速开始

```powershell
# 1. 安装依赖（建议用虚拟环境）
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# 2. 生成样例 PDF（一正确一错误）体验流程
python tools\make_sample.py

# 3. 运行核验（处理 inputfile/ 下所有 PDF）
python main.py

# 或指定文件
python main.py inputfile\样例_填写错误.pdf
```

## 目录结构

```
人员核验/
├── inputfile/           # 放入待核验的 PDF
├── output/              # 核验报告（xlsx/json，自动生成）
├── config/rules.yaml    # 校验规则（需按实际表单定制）
├── pdf_checker/         # 核心代码
│   ├── extract.py       #   提取打印文本（含坐标）
│   ├── locate.py        #   按标签定位字段值
│   ├── validators.py    #   校验规则实现
│   ├── validate.py      #   校验编排
│   └── report.py        #   报告输出
├── tools/make_sample.py # 样例 PDF 生成器
└── main.py              # 入口
```

## 如何定制规则（拿到真实表单后）

打开 `config/rules.yaml`，按实际表单修改：

1. `fields`：每个字段写清楚 PDF 中的**标签文字**（如“姓名”“身份证号”）、
   是否必填、取值位置（`right` = 标签同行右侧 / `below` = 标签下一行）。
2. `rules`：支持 `regex`（正则）、`in_set`（枚举）、`date`（日期）、
   `id_card`（身份证号）四种校验。
3. `cross_checks`：跨字段一致性（身份证出生日期/性别 ↔ 填写字段）。

> 当前 rules.yaml 中是示例规则，覆盖姓名/身份证/性别/电话/出生日期，
> 请以真实表单为准调整或增删。

## 报告

- `output/核验报告_<时间戳>.xlsx`：汇总 + 逐字段明细（通过/不通过标色）
- `output/核验结果_<时间戳>.json`：结构化结果，便于程序处理
