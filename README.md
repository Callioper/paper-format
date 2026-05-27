# paper-format

> 学术论文格式检测、自动修复与引文验证一站式工具
>
> Academic paper format checking, auto-fix, and citation verification

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 功能一览

| 功能 | 说明 |
|------|------|
| 格式检测 | 页边距、字体字号、行距、标题层级、三线表、页眉页脚、封面、摘要、目录、致谢 |
| 自动修复 | 一键修复格式问题，生成修复前后对比报告 |
| 引文格式化 | CNU《外国文学评论》/ CSL / GB/T 7714 标准 |
| 引文验证 | 本地 .bib 验证 + OpenAlex/CrossRef/知网/万方外部验证 |
| 引文逐字核对 | PDF 出处比对，支持扫描版 OCR + 截图高亮 |
| 模板生成 | 预配置中文学术论文模板（期刊/学位论文） |
| 样本解析 | 从格式正确的论文中提取格式规则 |
| 对比报告 | 6 列对比表 + 修复记录 + 统计摘要 + 引文核对报告 |

## 快速开始

```bash
# 克隆仓库
git clone https://github.com/Callioper/paper-format.git
cd paper-format

# 安装依赖
pip install -r requirements.txt

# 格式检测
python paper_format.py check 论文.docx

# 格式修复
python paper_format.py fix 论文.docx --output repaired.docx

# 生成模板
python paper_format.py template --mode thesis --output 模板.docx
```

## 安装方式

### 方式 1：Claude Code

下载 `paper-format.skill` 文件，拖入 Claude Desktop 或通过插件管理器安装。

或手动安装：
```bash
cp -r paper-format ~/.claude/skills/
```

触发方式：输入 `/paper-format` 或说"帮我检查论文格式"

### 方式 2：通用 Agent（Hermes / OpenClaw / Codex）

```bash
# 解压到 agent 的 skills 目录
unzip paper-format-universal.zip -d <agent-skills-dir>/

# 安装依赖
cd paper-format && pip install -r requirements.txt
```

Agent 读取 `SKILL.md` 获取指令，调用 `python paper_format.py <command>` 执行功能。

### 方式 3：命令行直接使用

```bash
pip install python-docx lxml
python paper_format.py check 论文.docx --mode journal
```

## 使用指南

### 文档模式

| 模式 | 适用 | 检测项 |
|------|------|--------|
| `journal`（默认） | 期刊论文/小论文 | 标题页、摘要、关键词、正文、参考文献、脚注 |
| `thesis` | 学位论文/大论文 | 封面、摘要、目录、正文、参考文献、致谢、附录 |

### 完整命令

```bash
# 格式检测
python paper_format.py check 论文.docx [--mode thesis|journal] [--spec spec.json]

# 格式修复（自动创建副本，不修改原文件）
python paper_format.py fix 论文.docx --output repaired.docx

# 生成对比报告
python paper_format.py report 原始.docx 修复后.docx 报告.docx

# 生成预格式化模板
python paper_format.py template --mode thesis --output 模板.docx

# 样本解析（从格式正确的论文提取规则）
python paper_format.py spec 样本.docx --output spec.json

# CSL 解析
python paper_format.py csl style.csl --output rules.json

# 本地引文验证（.bib 文件）
python paper_format.py verify 论文.docx --bib refs.bib --output result.json

# 外部引文验证（API）
python paper_format.py verify 论文.docx --sources openalex,crossref --output result.json

# 交叉校验（正文引用 ↔ 参考文献）
python paper_format.py cross 论文.docx --output cross.json

# 引文逐字核对（PDF 出处比对）
python paper_format.py quotes 论文.docx --sources "书.pdf:1-100" --output quote.json

# 引用格式修复
python paper_format.py cite 论文.docx --bib refs.bib --output report.docx
```

## 引文验证

### 本地数据库验证

从 Zotero + Better BibTeX 导出 .bib 文件：

1. 安装 [Better BibTeX](https://github.com/retorquere/zotero-better-bibtex) 插件
2. 在 Zotero 中选择文献库 → 右键 → 导出文献库 → Better BibTeX 格式
3. 勾选"保持更新"
4. 保存为 .bib 文件

```bash
python paper_format.py verify 论文.docx --bib refs.bib
```

### 外部 API 验证

| 数据源 | 类型 | 配置要求 |
|--------|------|---------|
| OpenAlex | 免费 API | 无需配置 |
| CrossRef | 免费 API | 无需配置 |
| Semantic Scholar | 免费 API | 可选 S2_API_KEY |
| CORE | 免费 API | 可选 CORE_API_KEY |
| 知网 CNKI | HTML 抓取 | 需要 CNKI_COOKIE（有封禁风险） |
| 万方 Wanfang | HTML 抓取 | 无需配置 |

### 引文逐字核对

从论文中提取直接引文，与出处 PDF 原文逐字比对：

```bash
python paper_format.py quotes 论文.docx --sources "书名.pdf:1-100" --output quote.json

# 扫描版 PDF 强制 OCR
python paper_format.py quotes 论文.docx --sources "扫描版.pdf" --ocr --output quote.json

# 导出 Excel（含截图嵌入）
python paper_format.py quotes 论文.docx --sources "书.pdf" --excel 核对报告.xlsx
```

## 依赖

### 必需
```bash
pip install python-docx lxml
```

### 可选
```bash
pip install pdfplumber Pillow openpyxl  # PDF + 截图 + Excel
pip install paddleocr                    # 扫描版 OCR（重型）
```

## 目录结构

```
paper-format/
├── paper_format.py          # 统一入口
├── SKILL.md                 # AI agent 指令（引导流程）
├── README.md                # 本文档
├── requirements.txt         # Python 依赖
├── package.py               # 通用打包脚本
├── scripts/
│   ├── check_format.py      # 格式检测（编排器）
│   ├── fix_format.py        # 格式修复
│   ├── generate_report.py   # 报告生成
│   ├── generate_template.py # 模板生成
│   ├── csl_parser.py        # CSL 解析
│   ├── parse_spec.py        # 样本解析
│   ├── verify_local.py      # 本地引文验证
│   ├── verify_external.py   # 外部引文验证
│   ├── verify_quotes.py     # 引文逐字核对 + OCR + 高亮 + Excel
│   ├── cross_check.py       # 交叉校验
│   ├── citation_repair.py   # 引用格式修复
│   ├── extract_references.py
│   ├── extract_footnotes.py
│   ├── checks/
│   │   └── format_checks.py # 检测函数库
│   ├── lib/
│   │   ├── bib_parser.py    # .bib 文件解析
│   │   ├── citation_models.py
│   │   ├── doc_helpers.py   # .docx/.md 文档工具
│   │   └── quote_models.py
│   └── sources/
│       ├── source_openalex.py
│       ├── source_crossref.py
│       ├── source_semantic_scholar.py
│       ├── source_core.py
│       ├── source_cnki.py
│       └── source_wanfang.py
└── references/
    ├── CNU_citation_rules.md
    ├── GBT7714-2015-numeric.csl
    └── common_issues.md
```

## 与其他 Agent 集成

### Python API

```python
from scripts.check_format import check_format
result = check_format("thesis.docx", mode="journal")

from scripts.fix_format import fix_format
result = fix_format("thesis.docx", output_path="repaired.docx")

from scripts.generate_template import generate_template
generate_template(mode="thesis", output_path="template.docx")
```

### SKILL.md

将 `SKILL.md` 内容复制到 agent 的 skill 指令中。SKILL.md 包含完整的引导决策树（Step 0-6），指导 agent 按正确顺序调用脚本。

## 下载

| 文件 | 说明 |
|------|------|
| [paper-format.skill](../../releases) | Claude Code 专用包 |
| [paper-format-universal.zip](../../releases) | 通用分发包（Hermes/OpenClaw/Codex） |

## 参考项目

本项目在开发过程中参考了以下开源项目和资源：

| 项目 | 参考内容 |
|------|---------|
| [sashavegal/paper-check](https://clawhub.ai/sashavegal/paper-check) | 原始 paper-check 项目的格式检测流程、三线表检测、修复记录结构、对比报告设计 |
| [ganzhi-black/humanities-thesis-skill](https://github.com/ganzhi-black/humanities-thesis-skill) | 引文验证的数据源架构（OpenAlex/CrossRef/知网/万方并行查询）、防幻觉规则引擎、知网/万方 HTML 抓取实现 |
| [Gostyan/docx-skill-4-cn-paper](https://github.com/Gostyan/docx-skill-4-cn-paper) | 中文学术论文 .docx 模板生成的预配置方案、样式定义参考 |
| [citation-checker](https://clawhub.ai/) (Zephyr Fang) | 引文逐字核对功能设计：PDF 出处比对、截图证据嵌入 Excel、OCR 扫描版支持 |
| [retorquere/zotero-better-bibtex](https://github.com/retorquere/zotero-better-bibtex) | Better BibTeX .bib 文件解析、file 字段 PDF 路径提取、引文元数据丰富化方案 |
| [citation-style-language/styles](https://github.com/citation-style-language/styles) | GB/T 7714-2015 CSL 文件，用于 CSL 解析器开发和测试 |
| [redleafnew/Chinese-STD-GB-T-7714-related-csl](https://github.com/redleafnew/Chinese-STD-GB-T-7714-related-csl) | 中文 CSL 样式参考 |

## License

MIT
