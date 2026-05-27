---
name: paper-format
description: >
  学术论文格式检测、自动修复与引文验证一站式工具。必须调用此技能：当用户说"检查论文格式"、"格式化论文"、
  "论文格式不对"、"格式审查"、"格式修改"、"/paper-format"、"/paper-check"，或提到"毕业论文"、"学位论文"、
  "thesis format"、"论文排版"、"格式规范"、"页边距"、"字体"、"行距"、"参考文献格式"时，立即调用此技能。
  即使用户只说"帮我看看论文有没有问题"、"这个论文格式对不对"、"论文被导师打回来了"、"论文要交了帮我查一下"，
  只要涉及学术论文的 .docx 文件，都应触发此技能而非手动处理。
  也适用于：引文完整性校验（用 OpenAlex/CrossRef/知网/万方验证引用真实性）、
  引文逐字核对（比对论文引文与出处 PDF 原文是否一致，支持扫描版 OCR，PDF 路径自动从 .bib 提取）、
  参考文献格式化（CNU 体例 / CSL 标准 / Better BibTeX .bib 丰富化）、
  样本文件解析（用格式正确的论文当参考模板）、生成预格式化论文模板、
  以及生成包含截图证据的 Excel 核对报告。
  支持学位论文（大论文）和期刊论文（小论文，默认）两种模式。
---

# paper-check — 学术论文格式检测与修复

一站式论文格式检测与修复。输入论文 .docx/.md，输出修复后的 .docx + 检测报告。

## 快速判断：用脚本还是手动？

| 条件 | 方式 |
|------|------|
| Python >= 3.9 可用 + `python-docx` 已安装 | ✅ 运行脚本（推荐，精确可靠） |
| Python 不可用或依赖缺失 | ⚠️ 手动模式（用 python-docx 逐段检查） |
| 用户只问"格式对不对"不提供文件 | 💬 引导用户上传 .docx/.md 文件 |

**安装依赖**（如缺失）：
```bash
pip install python-docx lxml pdfplumber Pillow openpyxl
```

---

## 输入

| 输入 | 类型 | 必填 | 说明 |
|------|------|------|------|
| 论文文档 | `.docx` | ✅ | 待检测/修复的论文 |
| 规范文档 | `.docx/.doc/.txt/.pdf` | ❌ | 学校格式规范，未提供时回退 GB/T 7713.1-2006 通用规则 |
| 参考样本 | `.docx` | ❌ | 格式正确的样本论文，脚本自动解析其格式规则 |
| CSL 文件 | `.csl` | ❌ | 引用样式标准文件，用于替代默认 CNU 引用规则 |
| 本地数据库 | `.bib` | ❌ | Better BibTeX 导出的参考文献库，用于引文验证 + 元数据丰富化 + PDF 路径提取 |

### 文档模式

支持两种模式，首次交互时询问用户（默认 `journal`）：

| 模式 | 适用 | 检测项 |
|------|------|--------|
| `journal`（默认） | 期刊论文/小论文 | 标题页、摘要、关键词、正文、参考文献、脚注 |
| `thesis` | 学位论文/大论文 | 封面、摘要、目录、正文、参考文献、致谢、附录 |

```bash
# 期刊论文模式（默认）
python scripts/check_format.py "论文.docx" --output result.json

# 学位论文模式
python scripts/check_format.py "论文.docx" --mode thesis --output result.json
```

### 样本文件解析

如果用户提供了格式正确的样本论文：

```bash
python scripts/parse_spec.py "样本.docx" --output spec.json
python scripts/check_format.py "论文.docx" --spec spec.json --output result.json
python scripts/fix_format.py "论文.docx" --spec spec.json --output repaired.docx
```

脚本从样本中提取：页面设置、字体字号、行距、标题样式、表格边框、页眉页脚格式，替代硬编码默认值。

### 模板生成

如果用户需要从零创建论文模板：

```bash
# 期刊论文模板（默认）
python scripts/generate_template.py --output "论文模板.docx"

# 学位论文模板（含封面、目录、致谢、附录）
python scripts/generate_template.py --mode thesis --output "毕业论文模板.docx"

# 用 spec.json 自定义模板格式
python scripts/generate_template.py --spec spec.json --output "自定义模板.docx"
```

模板包含：预设页面设置、样式定义（Normal/Heading 1-4）、三线表、占位文本、章节结构。

### CSL 引用标准

如果用户提供了 .csl 文件：

```bash
python scripts/csl_parser.py "style.csl" --output csl_rules.json
```

解析 CSL XML 提取：作者格式、标题格式（引号/斜体/书名号）、日期格式、页码格式、文献类型标识。
未提供 CSL 时默认使用 CNU《外国文学评论》2024修订版规则。

### 引文完整性校验（两步验证）

**Step 1: 本地数据库验证**
```bash
python scripts/verify_local.py "论文.docx" --bib refs.bib --output local_result.json
```

支持 .bib (Better BibTeX/BibLaTeX)、.ris (EndNote)、.xml (Zotero XML) 格式。
.bib 文件来自 Zotero + Better BibTeX 插件导出，自动提取 PDF 路径（file 字段）。
模糊匹配标题 + 对比作者/年份字段。

**Step 2: 外部数据源验证**
```bash
python scripts/verify_external.py "论文.docx" --sources openalex,crossref --output ext_result.json
```

对本地未匹配的引用，查询 OpenAlex/CrossRef/Semantic Scholar 等外部学术数据库。

**交叉校验**
```bash
python scripts/cross_check.py "论文.docx" --output cross_check.json
```

检查：孤立引用（正文有但参考文献无）、未使用文献（参考文献有但正文未引用）、重复引用、不完整条目。

---

## 工作流程（完整引导决策树）

### Step 0：信息收集（触发后第一步，全部必问）

用户触发技能后，**先收集全部信息再跑脚本**。以下所有问题都必须问，不能跳过：

**处理前必做**：
1. 提醒用户："处理前建议备份原文件。我会复制一份副本进行处理，不会修改您的原文档。"
2. 将用户文件复制到以论文名命名的文件夹中，后续所有输出都放在该文件夹内。

**全部必问**：
1. 论文文件路径在哪里？
2. 是期刊论文（小论文）还是学位论文（大论文）？→ 默认 `journal`
3. 有没有格式正确的参考文档（规范文档）？→ 可以是格式正确的样本论文或学校的格式规范文件，用 parse_spec.py 提取规则
4. 参考文献格式用什么标准？有没有 .csl 文件？
   - 有 .csl → 用 csl_parser.py 解析
   - 没有但有投稿指南链接或参考文献格式截图 → 用 journal-csl-builder 从链接/图片生成 .csl
   - 没有但想要 → 搜索 https://www.zotero.org/styles 或触发 journal-csl-builder 创建
   - 都不需要 → 默认 CNU《外国文学评论》2024修订版
5. 用 Zotero 管理参考文献吗？有没有 .bib 文件？
   - 有 .bib → 用于引文验证 + 元数据丰富化 + PDF 路径提取
   - 没有 → 跳过本地验证，但仍然做外部验证和交叉校验
6. 需不需要验证引用的真实性？（推荐：默认做外部验证 + 交叉校验）

**话术**：
> "收到！开始处理前确认几件事：
> 1. 这篇是期刊论文还是学位论文？
> 2. 有没有格式正确的参考文档？（可以是模板论文或学校格式规范）
> 3. 参考文献格式有 .csl 文件吗？没有的话我可以用默认的 CNU 格式，或者帮你从 Zotero Styles 搜索
> 4. 用 Zotero 管理文献吗？有 .bib 文件可以做引文验证
> 5. 需要验证引用的真实性吗？我可以用 OpenAlex/CrossRef 等学术数据库查询"

### Step 1：环境准备 + 输出目录 + 规范确定

**创建输出目录**：所有输出文件放入以论文名命名的文件夹。
```
D:/path/to/论文名/
├── 论文名_copy.docx          ← 副本（修复对象）
├── 论文名_repaired.docx      ← 修复后
├── spec.json                 ← 规范（如有）
├── csl_rules.json            ← CSL 规则（如有）
├── check_result.json         ← 检测结果
├── local_result.json         ← 本地验证结果
├── ext_result.json           ← 外部验证结果
├── cross_check.json          ← 交叉校验结果
└── 格式检测报告.docx          ← 最终报告
```

**规范文档统一处理**（样本论文和学校规范文档是同一个概念）：
```
决策树：
  用户提供规范文档（样本论文/学校规范文件）？
    ├─ 是 .docx → parse_spec.py 提取 spec.json → 用 --spec 参数
    ├─ 是 .txt/.pdf → Read 读取，提取数值覆盖默认值
    └─ 否 → 使用内置默认值
```

**CSL 处理**：
```
  用户有 .csl 文件？
    ├─ 是 → csl_parser.py 解析为 csl_rules.json
    └─ 否 → 有投稿指南链接或格式截图？
              ├─ 是 → 触发 journal-csl-builder 技能，从链接/图片生成 .csl
              │        生成后用 csl_parser.py 解析为规则
              └─ 否 → 需要自定义格式？
                        ├─ 是 → 触发 journal-csl-builder 技能（用户提供样例）
                        │        或搜索 https://www.zotero.org/styles
                        └─ 否 → 使用默认 CNU《外国文学评论》2024修订版格式
```

**journal-csl-builder 触发条件**：当用户说"帮我做一个期刊的引用格式"、"这个期刊的 CSL"、
"参考文献格式怎么做 Zotero 样式"、"从这个链接生成 CSL"、"根据这个截图生成引用格式"，
或 Step 0 第 4 问中用户选择"没有但有链接/截图"时，调用 `Skill("journal-csl-builder")`。
用户提供期刊投稿指南网页链接或参考文献格式截图，journal-csl-builder 从中提取规则并生成 .csl 文件。

```bash
# 如有规范文档（样本论文）
python scripts/parse_spec.py "规范文档.docx" --output "输出目录/spec.json"

# 如有 CSL 文件
python scripts/csl_parser.py "style.csl" --output "输出目录/csl_rules.json"
```

### Step 2：格式检测

```bash
python scripts/check_format.py "论文_copy.docx" --mode journal --output "输出目录/check_result.json"
# 或 --mode thesis

# 如有 spec.json
python scripts/check_format.py "论文_copy.docx" --spec "输出目录/spec.json" --output "输出目录/check_result.json"
```

### Step 3：问题诊断 + 用户确认

读取 `check_result.json` 的 `issues` 字段，向用户展示问题表格。**等待用户确认后再修复。**

### Step 4：自动修复

```bash
python scripts/fix_format.py "论文_copy.docx" --mode journal --output "输出目录/论文_repaired.docx"
```

**自动修复项**：页边距、正文字体字号、标题字体字号、行距、首行缩进、章节标题格式、参考文献字体格式

**脚注处理**：仅修改脚注字体样式（宋体 + TNR 10.5pt），不修改脚注段落间距和行距。

**无法自动修复（告知用户）**：页眉横线、目录更新、图片/公式位置

### Step 5：引文验证

**执行前必须询问用户**：Step 0 中已问过 .bib 文件，但如果当时没确认，此处再次询问：
> "你有 Zotero 导出的 .bib 文件吗？有的话可以用它验证引用完整性。
> 没有的话我直接用外部数据库（OpenAlex/CrossRef）验证。"

```bash
# Step 5a: 本地验证（如有 .bib，必做）
python scripts/verify_local.py "论文_copy.docx" --bib refs.bib --output "输出目录/local_result.json"

# Step 5b: 外部验证（始终执行，用免费数据源）
python scripts/verify_external.py "论文_copy.docx" --sources openalex,crossref --output "输出目录/ext_result.json"

# Step 5c: 交叉校验（始终执行）
python scripts/cross_check.py "论文_copy.docx" --output "输出目录/cross_check.json"
```

如果用户有 .bib 文件且有出处 PDF，询问是否需要引文逐字核对：
```bash
python scripts/verify_quotes.py "论文_copy.docx" --sources "书.pdf:1-100" --output "输出目录/quote_result.json"
```

### Step 6：报告生成

```bash
python scripts/generate_report.py "论文_copy.docx" "输出目录/论文_repaired.docx" "输出目录/格式检测报告.docx" \
    --format-issues "输出目录/check_result.json"
```

报告要求：中文用宋体，非中文用 Times New Roman。

---

## 参考文献格式规则

详见 `references/CNU_citation_rules.md`，核心格式：

| 类型 | 中文格式 | 英文格式 |
|------|---------|---------|
| 专著 | 作者：《书名》，出版地：出版社，年份，第X页。 | Author, *Title*, Place: Publisher, Year, p. X. |
| 期刊 | 作者：《文章题名》，《期刊名》年份第X期。 | Author, "Article," *Journal*, vol. X, no. X (Year), pp. X-X. |
| 析出 | 作者：《章节名》，编者编：《书名》，出版地：出版社，年份，第X—X页。 | Author, "Ch.," in *Book*, ed. Editor, Place: Publisher, Year, p. X. |

格式优先级：CNU 规范 > CSL 文件规则 > GB/T 7714-2015 通用格式

---

## 脚本参考

所有脚本位于 `scripts/` 目录下：

| 脚本 | 用途 | 关键参数 |
|------|------|---------|
| `check_format.py` | 格式检测（含引用+脚注） | `论文.docx --output result.json` |
| `fix_format.py` | 格式修复 | `论文_copy.docx --output repaired.docx` |
| `generate_report.py` | 生成对比报告 | `论文.docx repaired.docx report.docx` |
| `generate_template.py` | 生成预格式化论文模板 | `--mode thesis --output 模板.docx` |
| `parse_spec.py` | 从样本 .docx 提取格式规则 | `sample.docx --output spec.json` |
| `csl_parser.py` | 解析 CSL 引用样式文件 | `style.csl --output csl_rules.json` |
| `verify_local.py` | 本地数据库引文验证 | `论文.docx --bib refs.bib --output result.json` |
| `verify_external.py` | 外部 API 引文验证 | `论文.docx --sources openalex,crossref --output result.json` |
| `cross_check.py` | 正文引用<->参考文献交叉校验 | `论文.docx --output cross_check.json` |
| `verify_quotes.py` | 引文逐字核对（PDF 出处比对） | `论文.docx --sources "book.pdf:1-50" --output result.json` |
| `citation_repair.py` | 单独修复引用格式 | `论文.docx --bib refs.bib --output report.docx` |

**注意**：所有脚本始终在副本文件上操作，不会修改用户的原文档。
| `extract_references.py` | 提取参考文献节 | `thesis.docx` |
| `extract_footnotes.py` | 提取脚注 | `thesis.docx` |

### 引文验证数据源

| 数据源 | 类型 | 配置要求 |
|--------|------|---------|
| OpenAlex | 免费 API | 无需配置 |
| CrossRef | 免费 API | 无需配置 |
| Semantic Scholar | 免费 API | 可选 S2_API_KEY |
| CORE | 免费 API | 可选 CORE_API_KEY |
| 知网 CNKI | HTML 抓取 | 需要 CNKI_COOKIE 环境变量 |
| 万方 Wanfang | HTML 抓取 | 无需配置 |

### 参考文献格式优先级

| 场景 | 处理方式 |
|------|---------|
| 有 .bib 文件 + CNU 规范 | .bib 元数据 + CNU 格式 |
| 无 .bib + 有规范文档 | 解析原始字段 + CNU 格式 |
| 无规范文档 | GB/T 7714-2015 通用格式 |

---

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| `ModuleNotFoundError: python-docx` | `pip install python-docx lxml` |
| 脚本运行报错 | 查看错误消息，通常是文件路径含空格或中文，用引号包裹 |
| 检测不到参考文献节 | 检查节标题是否为"参考文献"/"References"等标准名称 |
| 参考文献修复不完整 | CNU 格式复杂，部分需人工核实（P3 标记） |
| 页眉横线未修复 | python-docx 不支持页眉边框操作，需在 Word 中手动处理 |
| 用户文件是 .doc 格式 | 告知用户在 Word 中"另存为 .docx"后重试 |

---

## 注意事项

- 支持 `.docx` 格式（.doc 需用户先转换为 .docx）
- **处理前必须备份**：始终先复制文件为副本，在副本上操作，不修改原文档
- 默认模式为 `journal`（期刊论文），学位论文需用户明确说明
- CNU 格式优先于 GB/T 7714 generic — 冲突时以 CNU 为准
- 参考文献节检测关键词：`参考文献`、`References`、`引用`、`Works Cited`、`Bibliography`
- 脚注检测基于 XML 解析，需要 `lxml` 依赖
- 知网 CNKI 数据源需要 Cookie，有封禁风险，引导用户使用小号
