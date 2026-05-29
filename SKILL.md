---
name: paper-format
description: >
  学术论文格式检测、自动修复与引文验证一站式工具。触发词：检查论文格式、格式化论文、格式审查、
  毕业论文、学位论文、thesis format、论文排版、格式规范、页边距、字体、行距、参考文献格式、
  /paper-format、/paper-check。即使用户只说"帮我看看论文有没有问题"、"论文被导师打回来了"，
  只要涉及 .docx 文件都应触发。也适用于引文验证（OpenAlex/CrossRef/知网）、引文逐字核对、
  CSL 解析、样本模板解析。支持学位论文和期刊论文两种模式。
---

# paper-format — 学术论文格式检测与修复

一站式论文格式检测与修复。输入论文 .docx/.md，输出修复后的 .docx + 检测报告。

## 工作边界

这些边界决定了本技能的职责范围，所有操作都应遵守：

**默认只改格式，不改内容观点。** 除非用户明确要求润色、改写或补写正文，否则只处理版式、结构层级、样式统一、编号、图表、脚注、参考文献和页面设置。如果发现正文逻辑、语言或引文内容本身有问题，可以单独提示，但不要混进格式调整里默默改写。

**符号审查属于默认处理范围。** 即使用户只说"排版"，也应同步排查并修正明显的中英文符号误用、全半角混乱和不符合中英文书写习惯的标点搭配。

**优先使用用户提供的本地材料。** 格式要求如果是本地 PDF、Word 模板、截图或院校说明文件，先直接读取。不要先把"读本地 PDF"混说成"联网查规范"。

**保留原稿，不在原文件上直接覆盖。** 始终先复制文件为副本，在副本上操作。最终说明里应明确写清来源：哪些规则来自用户提供的格式说明，哪些属于为了落地排版作出的合理补齐。补齐项要可解释，不能偷偷替用户定标准。

**图表按对象处理，不按截图处理。** 能保留可编辑表格、可编辑题注和可复用编号，就不要退化成贴图式处理。

---

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
| 论文文档 | `.docx` / `.doc` | ✅ | 待检测/修复的论文（.doc 会自动转换） |
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

### CSL 引用标准（完整流水线）

CSL 处理分两种情况：用户提供现成 .csl，或需要从零生成。

**情况 A：用户已有 .csl 文件**

执行三阶段校验流水线，只有全部通过才算可用：

```bash
# 阶段 1：XML 结构校验
python scripts/validate_csl.py "style.csl"

# 阶段 2：语义检查（中英文 locale 串用、URL 泄漏、中文误出 vol./no. 等）
python scripts/lint_csl_semantics.py "style.csl"

# 阶段 3：提取格式规则供 citation_repair.py 使用
python scripts/csl_parser.py "style.csl" --output csl_rules.json
```

如果阶段 3 环境可调用 Zotero 本地 API，追加冒烟测试：
```bash
python scripts/smoke_test_zotero_style.py "style.csl" \
    --item zh_journal=HXKEDRBQ --item en_journal=YYYYYYYY \
    --json --output smoke_result.json
```

**情况 B：需要从零生成 CSL**

当用户说"帮我做一个期刊的引用格式"、"这个期刊的 CSL"、"从这个链接生成 CSL"、"根据这个截图生成引用格式"，或 Step 0 第 4 问中用户选择"没有但有链接/截图"时：

1. 触发 `journal-csl-builder` 技能（如果已安装），或按以下流程内联处理
2. 读取用户材料（本地 PDF / 投稿指南网页 / 参考文献样例），整理成"格式规则摘要"
3. 整理"能力映射表"：哪些规则 CSL 可直接实现、哪些只能部分实现、哪些超出 CSL 能力边界
4. 如果属于中英混排脚注制，优先参考 `references/bilingual-note-base.csl` 的双 layout 结构
5. 生成 .csl 文件后，运行上述三阶段校验流水线
6. 用 `fixtures/bilingual-note-cases.json` 中的测试条目做回归检查

**CSL 能力边界**（必须在交付说明中写清）：

| 能力 | CSL 可做 | CSL 不可做（需 Zotero/Word/人工） |
|------|---------|-------------------------------|
| 作者格式（顺序、分隔符、截断） | ✅ | |
| 标题格式（引号/斜体/书名号） | ✅ | |
| 日期格式（年/月/日） | ✅ | |
| 卷期页码格式 | ✅ | |
| 文献类型标识 [J][M][D] | ✅ | |
| 中英文双 layout 分流 | ✅ | |
| 每页脚注重新编号 | | ❌ Zotero 设置 |
| 转引文献（cf. / qtd. in） | | ❌ 需手动处理 |
| Word 超链接/拼写波浪线 | | ❌ Word 设置 |
| 文献类型标识自动推断 | 部分 | 复杂情况需人工 |

未提供 CSL 时默认使用 CNU《外国文学评论》2024修订版规则。

### 引文完整性校验（两步验证）

**Step 1: 本地数据库验证**
```bash
python scripts/verify_local.py "论文.docx" --bib refs.bib --output local_result.json
```

支持 .bib (Better BibTeX/BibLaTeX)、.ris (EndNote)、.xml (Zotero XML) 格式。.bib 文件来自 Zotero + Better BibTeX 插件导出，自动提取 PDF 路径（file 字段）。模糊匹配标题 + 对比作者/年份字段。

**Step 2: 外部数据源验证**
```bash
python scripts/verify_external.py "论文.docx" --sources openalex,crossref --output ext_result.json
```

**交叉校验**
```bash
python scripts/cross_check.py "论文.docx" --output cross_check.json
```

---

## 工作流程

### 两种执行路径

收集完信息（Step 0）、确认执行方案（Step 0.5）之后，有两种把方案落地的方式：

**① 一键编排器 `fix_paper.py`（推荐用于标准流程）**

```bash
python scripts/fix_paper.py "论文.docx" --mode journal [--spec spec.json] [--bib refs.bib]
```

它一次性完成：建输出目录 → 复制副本 → **委托 `fix_format.fix_format()` 做实际修复**（楷体保留、run-in 摘要、参考文献、CNU 引文都在那里，单一事实来源）→ 脚注字体规整 → 写出 `check_result.json` + `repair_records.json` → **把修复记录传给 HTML 报告并生成 `论文名_report.html`**。它本身只做编排，所有修复能力来自 fix_format，因此 Step 5 的任何改进会自动惠及这条路径，before/after 报告也天然正确。

**② 细粒度脚本路径（需要逐步控制、插入符号审查/引文验证/视觉复核时）**

按 Step 1 → Step 8 逐步执行。当你需要在修复前后插入 Step 3 的用户确认、Step 4 符号审查、Step 6 引文验证、Step 7 视觉复核时，走这条路。注意：这条路上要自己负责把修复记录落盘后再喂给报告（见 Step 8）。

> 多数情况下用 `fix_paper.py` 起步，遇到需要人工确认或额外审查的环节再切换到细粒度脚本补做。两条路用的是同一批底层脚本，不冲突。

### Step 0：信息收集

用户触发技能后，**先收集全部信息再跑脚本**。以下所有问题都必须问，不能跳过：

**处理前必做**：
1. 提醒用户："处理前建议备份原文件。我会复制一份副本进行处理，不会修改您的原文档。"
2. 将用户文件复制到以论文名命名的文件夹中，后续所有输出都放在该文件夹内。

**.doc 文件自动处理**：如果用户提供的是 `.doc` 格式，自动使用 LibreOffice 或 python-docx 将其转换为 `.docx` 工作副本，再继续后续流程。无需让用户手动"另存为"。如果转换失败，再告知用户手动转换。

**全部必问**：
1. 论文文件路径在哪里？
2. 是期刊论文（小论文）还是学位论文（大论文）？→ 默认 `journal`
3. 有没有格式正确的参考文档（规范文档）？→ 可以是格式正确的样本论文或学校的格式规范文件
4. 参考文献格式用什么标准？有没有 .csl 文件？
5. 用 Zotero 管理参考文献吗？有没有 .bib 文件？
6. 需不需要验证引用的真实性？

**话术**：
> "收到！开始处理前确认几件事：
> 1. 这篇是期刊论文还是学位论文？
> 2. 有没有格式正确的参考文档？（可以是模板论文或学校格式规范）
> 3. 参考文献格式有 .csl 文件吗？没有的话我可以用默认的 CNU 格式，或者帮你从 Zotero Styles 搜索
> 4. 用 Zotero 管理文献吗？有 .bib 文件可以做引文验证
> 5. 需要验证引用的真实性吗？我可以用 OpenAlex/CrossRef 等学术数据库查询"

### Step 0.5：解析格式要求 → 形成执行方案 → 确认后再动手

这是整个流程的关键转折点。在跑任何修改脚本之前，必须先把格式要求拆清楚，形成明确执行方案，让用户确认后再执行。

**解析格式要求**：

读取用户提供的规范文档（PDF / Word 模板 / 截图 / 院校说明 / 投稿指南），至少梳理清楚以下项目：

| 类别 | 要点 |
|------|------|
| 页面与版心 | 纸张大小、页边距、装订线、页眉页脚、页码位置 |
| 字体与字号 | 标题层级、正文、摘要、关键词、注释、参考文献、图表题注 |
| 段落规则 | 首行缩进、行距、段前段后、对齐方式、分页控制 |
| 结构规则 | 中英文标题、作者信息、摘要、关键词、目录、正文层级、致谢、附录 |
| 编号规则 | 章节号、图号、表号、公式号、注释号、参考文献序号 |
| 特殊对象 | 引文块、脚注、表格、图片、公式、参考文献 |
| 符号规则 | 中文标点、英文标点、全角半角、空格、括号、引号、破折号、连接号 |

**形成执行方案**：

执行方案不是泛泛复述，而是要写成能直接落到 Word 的具体参数。例如：

```
- 一级标题：黑体、小二、居中、段前 24 磅、段后 18 磅
- 正文：宋体、小四、两端对齐、首行缩进 2 字符、固定值 24 磅
- 图题：宋体、五号、居中，置于图下
- 参考文献：宋体 + TNR 10.5pt，悬挂缩进，编号 [1] [2] ...
```

**来源标注**（透明度机制）：

在执行方案中对每条规则标注来源：

| 标记 | 含义 |
|------|------|
| 📄 用户提供 | 来自用户提供的格式说明文件 |
| 🔧 合理补齐 | 原始说明未覆盖，按中文学术写作惯例补齐 |
| 📐 默认规范 | 用户未提供规范文档时，回退到 GB/T 7713.1-2006 |

如果原始说明有冲突、缺项或表述模糊，先指出，再给出拟采用的处理办法。

**向用户展示方案，等待确认**：

> "我已经把格式要求整理成执行方案了，你先看看有没有问题：
> [执行方案表格]
>
> 其中标 🔧 的几条是原始说明没覆盖到的，我按常见学术规范补齐了，你看是否合适。
> 确认后我就开始处理。"

**只有用户确认后，才进入 Step 1。** 如果用户对方案有修改意见，先调整方案再确认。

### Step 1：环境准备 + 输出目录 + 规范确定

**创建输出目录**：所有输出文件放入以论文名命名的文件夹。
```
D:/path/to/论文名/
├── 论文名_copy.docx          ← 副本（修复对象）
├── 论文名_repaired.docx      ← 修复后
├── spec.json                 ← 规范（如有）
├── csl_rules.json            ← CSL 规则（如有）
├── execution_plan.md         ← 执行方案（Step 0.5 产出）
├── check_result.json         ← 格式检测结果
├── content_result.json       ← 内容结构检查结果
├── symbol_result.json        ← 符号审查结果（规则引擎）
├── local_result.json         ← 本地验证结果
├── ext_result.json           ← 外部验证结果
├── cross_check.json          ← 交叉校验结果
├── ref_enhanced.json         ← 增强引用检查结果（DOI/字段/URL）
├── 论文名_report.html        ← 交互式 HTML 报告（主报告）
└── 格式检测报告.docx          ← Word 报告（可选，用于打印/批注）
```

**报告格式**：默认生成 HTML 交互式报告（折叠/图表/颜色编码），Word 报告作为可选导出。

**规范文档统一处理**：
```
决策树：
  用户提供规范文档（样本论文/学校规范文件）？
    ├─ 是 .docx → parse_spec.py 提取 spec.json → 用 --spec 参数
    ├─ 是 .txt/.pdf → Read 读取，提取数值覆盖默认值
    └─ 否 → 使用内置默认值
```

**CSL 处理**（三阶段校验流水线）：
```
  用户有 .csl 文件？
    ├─ 是 → validate_csl.py 结构校验 → lint_csl_semantics.py 语义检查 → csl_parser.py 提取规则
    └─ 否 → 有投稿指南链接或格式截图？
              ├─ 是 → 触发 journal-csl-builder 技能（或内联生成）
              │        生成后同样走三阶段校验
              └─ 否 → 使用默认 CNU《外国文学评论》2024修订版格式
```

```bash
# 如有规范文档（样本论文）
python scripts/parse_spec.py "规范文档.docx" --output "输出目录/spec.json"

# 如有 CSL 文件（三阶段流水线）
python scripts/validate_csl.py "style.csl"                                    # 阶段 1：结构校验
python scripts/lint_csl_semantics.py "style.csl"                              # 阶段 2：语义检查
python scripts/csl_parser.py "style.csl" --output "输出目录/csl_rules.json"    # 阶段 3：提取规则
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

### Step 3.5：内容结构检查

在格式检测之后、符号审查之前，检查论文的内容结构是否完整：

```bash
python scripts/check_content.py "论文_copy.docx" --mode journal --output "输出目录/content_result.json"
# 或 --mode thesis
```

**检查项**：

| 模式 | 必需章节 | 可选章节 |
|------|---------|---------|
| `journal` | 标题、摘要、关键词、正文、参考文献 | 英文摘要、作者信息、基金项目 |
| `thesis` | 封面、中文摘要、英文摘要、关键词、目录、正文、参考文献、致谢 | 附录、作者简介、声明 |

**额外检查**：
- 摘要字数（建议中文 200-300 字）
- 关键词数量（建议 3-8 个）
- 缩略语候选识别（全文大写缩略语列表）

向用户展示缺失章节和关键词信息，如有问题在后续修复中补充。

### Step 4：中英文符号审查（规则引擎）

符号审查是默认步骤，即使用户只说"排版"也要做。审查范围覆盖全文——标题、摘要、关键词、脚注、图表题注、公式说明、参考文献、附录和作者信息。

**使用规则引擎执行**：

```bash
python scripts/check_language.py "论文_copy.docx" --rules references/language_rules.yaml --output "输出目录/symbol_result.json"
```

规则引擎读取 `references/language_rules.yaml` 中的 16 类规则族，逐段扫描并输出结构化报告。每条规则支持 `enabled`、`severity`（P1/P2/P3）、`autofix_safe` 配置。

**16 类规则族**：

| 规则族 | 说明 | 严重度 | 可自动修复 |
|--------|------|--------|-----------|
| `cjk_latin_spacing` | 中文与英文/数字之间缺少空格 | P2 | ✅ |
| `repeated_punctuation` | 连续重复标点（。。。。） | P2 | ❌ |
| `mixed_quote_style` | 弯引号""与直角引号「」混用 | P2 | ❌ |
| `bracket_mismatch` | 括号不配对或类型不一致 | P1 | ❌ |
| `quote_mismatch` | 引号不配对 | P1 | ❌ |
| `fullwidth_halfwidth_mix` | 同段落全角/半角标点混用 | P2 | ❌ |
| `zh_en_symbol_spacing` | 中英文标点间距 | P2 | ✅ |
| `dash_style` | 破折号/连接号风格（-- → ——） | P2 | ❌ |
| `ellipsis_style` | 省略号风格（... → ……） | P2 | ✅ |
| `unit_spacing` | 数字与单位间距（5cm → 5 cm） | P3 | ✅ |
| `number_range_style` | 数字范围风格（1-5 → 1–5） | P3 | ❌ |
| `weak_phrases` | 学术写作弱表达 | P3 | ❌ |
| `booktitle_mixed_style` | 书名标题风格混用 | P3 | ❌ |
| `enum_punctuation_style` | 列举标点风格 | P3 | ❌ |
| `connector_blacklist_simple` | 连接词黑名单（然后、所以） | P3 | ❌ |
| `en_punct_in_cjk` | 中文语境误用英文标点（英文引号包中文、中文后英文逗号/句号/括号） | P2 | ❌ |

**自定义规则**：用户可在 `references/language_rules.yaml` 中启用/禁用规则、调整严重度、添加自定义模式。

**修正原则**：只修正明确的符号错误，不改动语义。`autofix_safe: false` 的规则标记为"建议复核"而非直接修改。按严重程度排序：P1（必须修复）> P2（应当修复）> P3（建议修复）。

### Step 4.6：假排版清理

"假排版"指用排版无关手段冒充版式（空段堆叠、段首手工空格、段内手工换行、行尾空格），
会让样式系统失效、跨设备错位（借鉴 cn-paper-typesetter）。

**这一步通常无需手动跑**：`fix_format`（以及一键入口 `fix_paper`）会在 Step 5 统一样式之后、保存之前**自动**清理高确定性项（折叠连续空段、去段首/行尾手工空格），并记入修复记录。段内手工换行(w:br)只标记不自动删（可能是有意换行，需人工确认）。

仅当你想在修复前**单独查看**有哪些假排版（生成报告、决定是否人工干预）时，才单独运行检测脚本：

```bash
python scripts/check_fake_layout.py "论文_copy.docx" --output "输出目录/fake_layout.json"
```

### Step 5：自动修复

**标准修复**（python-docx）：
```bash
python scripts/fix_format.py "论文_copy.docx" --mode journal --output "输出目录/论文_repaired.docx"
```

**自动修复项**：页边距、正文字体字号、标题字体字号、行距、首行缩进、章节标题格式、参考文献字体格式

**脚注处理**：
- 字体样式：仅修改脚注字体为宋体 + TNR 10.5pt，不修改段落间距和行距
- 脚注内容提取：用 `extract_footnotes.py` 从 docx 的 `word/footnotes.xml` 直接解析（python-docx 高层 API 有时会漏掉脚注内容）。它会把每条脚注分类为"完整引文 / 再次引证 / 纯说明"，方便快速定位哪些需要核对 CNU 格式：

```bash
python scripts/extract_footnotes.py "论文_copy.docx" --output "输出目录/footnotes.json"
```

- 脚注引用格式检查：对分类为"完整引文"的脚注，检查是否符合 CNU 格式——中文书名用《》，英文书名用斜体，页码用 `第X页` / `p. X`；不符合的标记为 P2 建议修复，**不**自动改写内容（脚注内容修改属于学术判断，需人工核实）

> **改写脚注 XML 时的两个坑（务必避免内容丢失/重复）**：
> 1. **一条脚注可能含多个 `<w:p>` 段落**（如"左边…参见A。右边…参见B。"分两段）。只处理 `fn.find("w:p")` 第一段会导致其余段落的原文残留 → 内容重复。要 `findall("w:p")` 遍历所有段落，多余段落整段删除，正文只放在首段。
> 2. **正文 run 可能嵌套**在 hyperlink/smartTag 等元素里，`p.findall("w:r")`（仅直接子级）删不干净。要用 `p.findall(".//w:r")` 删除所有后代 run（保留含 `w:footnoteRef` 的引用标记 run）。
> 3. 改写出错时，从**未修改的副本** `_copy.docx` 重新读取 `word/footnotes.xml` 干净重建，不要在已损坏的文件上叠加修补。

**XML 级深度修复**（python-docx 无法处理的项目）：

python-docx 的高层 API 无法操作页眉边框、文本框等内容。对于这些问题，直接操作文档底层 XML：

```python
from docx import Document
from lxml import etree

doc = Document("论文_copy.docx")
# 通过 doc.element 访问完整 XML 树
# 通过 section.header 访问页眉，操作其 XML 子元素
```

| 问题 | XML 修复方式 |
|------|-------------|
| 页眉横线 | 操作页眉段落的 `w:pBdr` 元素，删除或修改 `w:bottom` 边框 |
| 文本框内容 | 直接读写 `w:txbxContent` 中的段落和 run |
| 公式识别 | 检测 `m:oMath` 元素，提取或修正公式内容 |
| 表格精确格式 | 操作 `w:tblBorders` 控制三线表边框 |

**三线表标准格式**（学术论文必须使用）：
```
顶线：1.5pt 粗线 (w:top size="12")
栏目线：0.75pt 细线 (w:insideH size="6"，仅表头下方)
底线：1.5pt 粗线 (w:bottom size="12")
其余所有边框：none
```

**混合字体设置**（中英文混排标题）：

标题中的英文和数字不应显示为宋体，应使用独立的西文字体：

```xml
<w:rPr>
  <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:eastAsia="SimHei"/>
</w:rPr>
```

| 元素 | 中文字体 (eastAsia) | 西文字体 (ascii/hAnsi) |
|------|-------------------|---------------------|
| 正文 | 宋体 SimSun | Cambria Math / Times New Roman |
| 标题 | 黑体 SimHei | Cambria Math |
| 图表题注 | 宋体 SimSun | Cambria Math |

**仍需人工处理的项目**：目录更新（需 Word 刷新 Ctrl+A → F9）、艺术字/SmartArt

### Step 5.5：公式处理（如有 LaTeX 公式）

如果论文中包含 LaTeX 公式，可以用脚本将其转为原生 Word 公式（OMML），而非嵌入图片：

```bash
# 插入块级公式（3 列无边框表布局：留白 | 公式居中 | 编号右对齐）
python scripts/formula.py "论文_unpacked/" "Q_n(x,a) = r + \gamma \max" 1 --anchor "由此可得"

# 仅输出 XML 片段（不修改文件）
python scripts/formula.py --latex "E=mc^2" --number 2
```

依赖：`pandoc >= 2.0`（LaTeX → OMML 转换）

如果 pandoc 不可用，退而求其次：提示用户在 Word 中手动插入公式，或使用 `mathml-to-docx.js`（Node.js，需 `npm install temml fast-xml-parser`）。

### Step 6：引文验证

**执行前必须询问用户**：Step 0 中已问过 .bib 文件，但如果当时没确认，此处再次询问。

```bash
# Step 6a: 本地验证（如有 .bib，必做）
python scripts/verify_local.py "论文_copy.docx" --bib refs.bib --output "输出目录/local_result.json"

# Step 6b: 外部验证（始终执行，用免费数据源）
python scripts/verify_external.py "论文_copy.docx" --sources openalex,crossref --output "输出目录/ext_result.json"

# Step 6c: 交叉校验（始终执行）
python scripts/cross_check.py "论文_copy.docx" --output "输出目录/cross_check.json"

# Step 6d: 增强检查（DOI 校验 + 字段完整性 + 正文 DOI/URL 提取）
python scripts/check_references_enhanced.py "论文_copy.docx" --bib refs.bib --output "输出目录/ref_enhanced.json"
# 如需检查 URL 可达性（较慢）：
python scripts/check_references_enhanced.py "论文_copy.docx" --bib refs.bib --check-urls --output "输出目录/ref_enhanced.json"
```

**增强检查项**：

| 检查项 | 说明 |
|--------|------|
| DOI 格式校验 | 验证每条引用的 DOI 是否符合 `10.xxxx/...` 规范 |
| BibTeX 必填字段 | 按条目类型检查必填字段（article 需 author/title/journal/year） |
| BibTeX 推荐字段 | 检查推荐字段（article 建议有 volume/pages/doi） |
| 缺失 DOI 候选 | 标记应有但缺失 DOI 的条目 |
| URL 可达性 | HEAD 请求检查 URL 是否可达（可选，较慢） |
| 正文 DOI/URL | 提取正文中出现的 DOI 和 URL 列表 |

如果用户有 .bib 文件且有出处 PDF，询问是否需要引文逐字核对：
```bash
python scripts/verify_quotes.py "论文_copy.docx" --sources "书.pdf:1-100" --output "输出目录/quote_result.json"
```

### Step 7：视觉复核

修复完成后，对输出文档做一轮视觉层面的校验，不能只依赖 JSON 报告。

**复核方式**：
1. 用 `render_preview.py` 通过 LibreOffice 把修复后的文档渲染成 PDF/PNG 页图，逐页肉眼检查
   （python-docx 无法渲染图片，必须用 LibreOffice；soffice 缺失时降级为提取关键段落格式属性 + 提示在 Word 中人工复核）：

```bash
python scripts/render_preview.py "输出目录/论文_repaired.docx" --out-dir "输出目录/_preview"
```

2. 逐页检查以下项目：

| 检查项 | 说明 |
|--------|------|
| 页边距与版心 | 左右上下是否均匀，装订侧是否留够 |
| 标题上下间距 | 各级标题的段前段后是否统一、稳定 |
| 表格完整性 | 是否有挤压、断裂、越界、跨页断行 |
| 图片与题注 | 图片和题注的相对位置是否正确（题注在图下方） |
| 分页质量 | 是否有明显孤行寡行、大块留白、标题沉底 |
| 页眉页脚 | 内容是否正确，页码是否连续 |
| 参考文献区 | 编号对齐、缩进悬挂、字号行距是否统一 |
| 符号修正效果 | Step 4 中标记的符号问题是否已正确修正 |

**如果发现视觉问题**：回到 Step 5 补修，然后重新复核，直到视觉层面无明显异常。

### Step 8：报告生成

**主报告是 HTML**（交互式、可折叠、彩色编码），用 `generate_html_report.py` 生成。Word 报告是可选导出（用于打印/批注）。

**before/after 的正确机制是 `repair_records`，不是"重跑检测"。** 修复脚本 `fix_format.py` 在修复时会记录每一处改动（改了什么、从什么值改成什么值），这份"修复记录"就是报告里"已修复内容"卡片的数据来源。报告把"原稿检测结果（`--format-issues`）"和"修复记录（`--repair-records`）"两份数据对照展示，所以**不需要也不应该**在修复后的文档上重跑 `check_format.py`——那样反而会因为检测器读不到"改动来由"而把已修复项又标成未修复（这正是早期版本的困惑来源）。

如果你走的是 Step 1/5 的细粒度脚本路径，`--repair-records` 是可选参数：`fix_format.py` 把修复记录作为返回值（不自动落盘），所以要拿到 `repair_records.json` 需要在调用时自己写出来；没有它时报告仍能生成，只是"已修复内容"卡片会缺少 before/after 明细。要完整的 before/after，最省事的是直接用 `fix_paper.py`。

```bash
# 生成 HTML 主报告（original + repaired + 原始问题 + 修复记录）
python scripts/generate_html_report.py "输出目录/论文_copy.docx" "输出目录/论文_repaired.docx" \
    --format-issues "输出目录/check_result.json" \
    --repair-records "输出目录/repair_records.json" \
    --mode journal \
    --output "输出目录/论文_report.html"

# 可选：Word 报告（用于打印批注；只含检测项，不展示 before/after diff）
python scripts/generate_report.py "输出目录/论文_copy.docx" "输出目录/论文_repaired.docx" "输出目录/格式检测报告.docx" \
    --format-issues "输出目录/check_result.json"
```

> 提示：若用 **Step 1 的一键编排器 `fix_paper.py`**（推荐），它会自动把修复记录传给 HTML 报告并写出 `论文名_report.html`，无需手动拼上面这些参数。

报告要求：中文用宋体，非中文用 Times New Roman。

**报告语言与术语**：报告全程使用中文标签。严禁使用英文词汇替代：
- ❌ Expected → ✅ **规范要求**
- ❌ Actual → ✅ **当前值**
- ❌ Status → ✅ **状态**
- ❌ Fixed / Not Fixed → ✅ **已修复 / 未修复**

**报告必须包含**：

1. **已修复内容卡片**（绿色背景）：列出本次修复的全部操作，每条含"修复前 → 修复后"对比。格式示例：
   ```
   ✅ 参考文献 [3] 添加 [法] 国籍标注
   ✅ 出版社"江苏人民出版社" → "南京：江苏人民出版社"
   ✅ 正文字号 10pt → 12pt
   ```

2. **仍需关注项**：未能自动修复的问题，标注原因（P1/P2/P3 + 原因说明）

3. **来源说明**：每条规则标注来源（📄 用户提供 / 🔧 合理补齐 / 📐 默认规范）

4. **引文验证摘要**（如运行了 Step 6）

5. **人工复核指引**（折叠区，放在报告最底部）：

   ```
   ## 人工复核指引
   
   以下项目需在 Word 中手动确认：
   
   1. 页边距：Word → 布局 → 页边距 → 自定义页边距
      规范要求：上X cm / 下X cm / 左X cm / 右X cm
   
   2. 脚注引用格式：检查每条脚注是否符合 CNU 格式
      中文：作者：《书名》，译者译，出版地：出版社，年份，第X页。
      英文：Author, *Title*, Place: Publisher, Year, p. X.
   
   3. 参考文献完整性：确认每条文献均有出版地、出版社、年份
   
   4. 目录更新：全选（Ctrl+A）→ F9 刷新目录
   
   5. 全文校对：重点检查图表编号连续性、交叉引用准确性
   ```

**隐藏空白章节**：报告中如果某个章节（如摘要、页眉页脚）检测结果为空（无问题也无内容），默认折叠或隐藏，不展示空卡片。

---

## 参考文献格式规则

详见 `references/CNU_citation_rules.md`，核心格式：

| 类型 | 中文格式 | 英文格式 |
|------|---------|---------|
| 专著 | 作者：《书名》，出版地：出版社，年份，第X页。 | Author, *Title*, Place: Publisher, Year, p. X. |
| 译著（外文作者） | [国籍]作者：《书名》，译者译，出版地：出版社，年份。 | — |
| 期刊 | 作者：《文章题名》，《期刊名》年份第X期。 | Author, "Article," *Journal*, vol. X, no. X (Year), pp. X-X. |
| 析出 | 作者：《章节名》，编者编：《书名》，出版地：出版社，年份，第X—X页。 | Author, "Ch.," in *Book*, ed. Editor, Place: Publisher, Year, p. X. |

**两个需要 Claude 用判断来处理的引文细节**（脚本不做，因为它们需要外部知识）：

- `[国籍]` 标注：参考文献列表中外国作者的中文译著，在作者名前补充 `[国籍]`（如 `[法]`、`[美]`）。脚注引用不加。**这要求知道作者是谁、是哪国人**——正则脚本无法判断，所以由 Claude 依据 `references/CNU_citation_rules.md` 零节的国籍对照表逐条处理；国籍不明时保留待人工核实，不要乱标。
- 出版地完整性：出版社名前若缺城市（如 `江苏人民出版社`），CNU 要求补成 `南京：江苏人民出版社`。`citation_repair.py` 能**检测并标记** P1"缺少出版社信息"，但**不会自动补城市**（需要"出版社→注册地"的外部知识）。由 Claude 参照常见出版社注册地补全，拿不准的标记待核实。

> 这两项之所以不写进脚本：作者国籍和出版社所在地都属于"事实查询"，硬编码查表既不全也会出错。交给 Claude 按参考表判断，比假装脚本能自动搞定更诚实、也更准。

格式优先级：CNU 规范 > CSL 文件规则 > GB/T 7714-2015 通用格式

---

## 脚本参考

所有脚本位于 `scripts/` 目录下：

| 脚本 | 用途 | 关键参数 |
|------|------|---------|
| `check_format.py` | 格式检测（含引用+脚注） | `论文.docx --output result.json` |
| `check_language.py` | 语言/符号检查（规则引擎，16 类规则族） | `论文.docx --rules language_rules.yaml --output result.json` |
| `check_content.py` | 内容结构检查（必需章节、关键词、缩略语） | `论文.docx --mode thesis --output result.json` |
| `check_fake_layout.py` | 假排版检测/清理（空段堆叠、手工空格、段内换行） | `论文.docx --output result.json` |
| `render_preview.py` | 视觉渲染复核（LibreOffice docx→PDF→PNG） | `论文.docx --out-dir _preview` |
| `check_references_enhanced.py` | 增强引用检查（DOI/字段/URL） | `论文.docx --bib refs.bib --check-urls --output result.json` |
| `fix_paper.py` | **一键编排**：建目录+复制+修复+记录改动+生成 HTML 报告（推荐入口） | `论文.docx --mode journal [--spec][--bib]` |
| `fix_format.py` | 格式修复（修复记录作为返回值，CLI 不落盘） | `论文_copy.docx --mode journal --output repaired.docx` |
| `generate_html_report.py` | **生成 HTML 主报告**（含 before/after，吃修复记录） | `original.docx repaired.docx --format-issues check.json --repair-records records.json --output report.html` |
| `generate_report.py` | 生成 Word 报告（可选打印版，仅检测项无 diff） | `original.docx repaired.docx report.docx --format-issues check.json` |
| `generate_template.py` | 生成预格式化论文模板 | `--mode thesis --output 模板.docx` |
| `parse_spec.py` | 从样本 .docx 提取格式规则 | `sample.docx --output spec.json` |
| `formula.py` | 插入块级 LaTeX 公式（OMML） | `unpacked/ "E=mc^2" 1 --anchor "文本"` |
| `mathml-to-docx.js` | MathML→OMML 转换 | Node.js，需 temml + fast-xml-parser |
| `comment.py` | 批量添加批注 | `unpacked/ 0 "批注内容"` |
| `csl_parser.py` | 解析 CSL 引用样式文件 | `style.csl --output csl_rules.json` |
| `validate_csl.py` | CSL XML 结构校验 | `style.csl` |
| `lint_csl_semantics.py` | CSL 语义检查（中文期刊常见错误） | `style.csl` |
| `smoke_test_zotero_style.py` | CSL Zotero 渲染冒烟测试 | `style.csl --item zh_journal=KEY` |
| `verify_local.py` | 本地数据库引文验证 | `论文.docx --bib refs.bib --output result.json` |
| `verify_external.py` | 外部 API 引文验证 | `论文.docx --sources openalex,crossref --output result.json` |
| `cross_check.py` | 正文引用<->参考文献交叉校验 | `论文.docx --output cross_check.json` |
| `verify_quotes.py` | 引文逐字核对（PDF 出处比对） | `论文.docx --sources "book.pdf:1-50" --output result.json` |
| `citation_repair.py` | 单独修复引用格式 | `论文.docx --bib refs.bib --output report.docx` |
| `extract_references.py` | 提取参考文献节 | `thesis.docx` |
| `extract_footnotes.py` | 提取并分类脚注（完整引文/再次引证/纯说明） | `论文.docx --output footnotes.json` |

**注意**：所有脚本始终在副本文件上操作，不会修改用户的原文档。

**参考模板**：`references/new_doc_template.js` — 完整的中文学术论文 docx-js 模板，含所有样式定义、三线表、公式布局、自动编号等实现，可作为从零创建论文的参考。

**语言规则配置**：`references/language_rules.yaml` — 16 类语言/符号检查规则，支持启用/禁用、严重度调整、自定义模式。

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

### CSL 模板与测试夹具

| 资源 | 路径 | 用途 |
|------|------|------|
| 中英混排脚注制底稿 | `references/bilingual-note-base.csl` | 中英混排 CSL 的稳定骨架，示范双 layout 分流 |
| 最小回归条目 | `fixtures/bilingual-note-cases.json` | 5 类测试条目：中文期刊/图书、英文期刊/图书、网页 |
| 目标输出基线 | `expected/bilingual-note-cases.md` | 上述夹具对应的期望输出，用于肉眼核对 |

---

## 故障排除

| 问题 | 解决方案 |
|------|---------|
| `ModuleNotFoundError: python-docx` | `pip install python-docx lxml` |
| 脚本运行报错 | 查看错误消息，通常是文件路径含空格或中文，用引号包裹 |
| 检测不到参考文献节 | 检查节标题是否为"参考文献"/"References"等标准名称 |
| 参考文献修复不完整 | CNU 格式复杂，部分需人工核实（P3 标记） |
| 页眉横线未修复 | 用 python-docx 直接操作页眉 XML 中的 `w:pBdr` 元素删除边框 |
| 用户文件是 .doc 格式 | 自动转换为 .docx 工作副本；如果自动转换失败，告知用户在 Word 中"另存为 .docx"后重试 |
| 报告显示已修复项为"未修复" | 报告缺少修复记录：Word 报告只吃 `--format-issues`（原始问题），无法展示 before/after。改用 `generate_html_report.py` 并传 `--repair-records`，或直接用 `fix_paper.py`（自动带记录）。**不要**在修复后的文档上重跑 `check_format.py` 来"刷新状态" |
| 脚注内容提取不完整 | python-docx 高层 API 可能漏掉部分脚注，用 `extract_footnotes.py --output` 直接解析 `word/footnotes.xml` |
| 参考文献缺少 [国籍] 标注 | 外国作者中文译著需在作者名前加 `[国籍]`，由 Claude 依 `references/CNU_citation_rules.md` 零节的国籍表逐条判断（脚本不做） |
| 出版社名称缺城市前缀 | `citation_repair.py` 会检测并标 P1，但补城市需外部知识，由 Claude 参照出版社注册地补全（如 `商务印书馆` → `北京：商务印书馆`） |
| CSL 语义检查报"中文期刊误出 vol." | lint_csl_semantics.py 已捕获，修改 CSL 中中文期刊分支的 locale 控制 |
| CSL 生成后 Zotero 安装失败 | 先跑 validate_csl.py 确认 XML 结构正确，再检查 info/id 是否唯一 |
| 中英混排 CSL 英文条目混入中文标点 | 检查 default-locale 设置，英文分支需显式 locale="en" |
| formula.py 报错 "pandoc not found" | 安装 pandoc >= 2.0，或改用 mathml-to-docx.js (Node.js) |
| 文本框内容无法读取 | python-docx 不支持文本框，需直接操作 XML 中的 `w:txbxContent` |
| 三线表边框仍可见 | 检查 `w:tblBorders` 中 `w:insideH` 和 `w:insideV` 是否设为 `none` |

---

## 注意事项

- 支持 `.docx` 和 `.doc` 格式（.doc 自动转换为 .docx 工作副本）
- **处理前必须备份**：始终先复制文件为副本，在副本上操作，不修改原文档
- 默认模式为 `journal`（期刊论文），学位论文需用户明确说明
- CNU 格式优先于 GB/T 7714 generic — 冲突时以 CNU 为准
- 参考文献节检测关键词：`参考文献`、`References`、`引用`、`Works Cited`、`Bibliography`
- 脚注检测基于 XML 解析，需要 `lxml` 依赖
- 知网 CNKI 数据源需要 Cookie，有封禁风险，引导用户使用小号
- 符号审查是默认步骤，不需要用户单独要求
- 执行方案必须经过用户确认才能开始修改
- 公式处理需要 `pandoc >= 2.0`（LaTeX→OMML），不可用时提示用户手动插入
- 批注处理需要 `defusedxml`（`pip install defusedxml`）
- 语言规则引擎需要 `PyYAML`（`pip install PyYAML`），不可用时回退到内置默认规则
- 混合字体设置：标题中文用黑体 (eastAsia)，英文/数字用 Cambria Math (ascii/hAnsi)
- 三线表：python-docx 的 `tbl.style` 无法精确控制边框，需直接操作 XML `w:tblBorders`
