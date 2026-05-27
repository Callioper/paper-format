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

### Step 0：信息收集 + 备份提醒（触发后第一步，不要跳过）

**处理前必做**：
1. 提醒用户："处理前建议备份原文件。我会复制一份副本进行处理，不会修改您的原文档。"
2. 将用户文件复制为 `论文_copy.docx`，后续所有脚本操作都在副本上执行。

用户触发技能后，**先收集信息再跑脚本**。分必问和选问：

**必问（缺任何一个都要问）**：
1. 论文文件路径在哪里？（引导用户上传或提供路径）
2. 是期刊论文（小论文）还是学位论文（大论文）？→ 默认 `journal`

**选问（有就问，没有就跳过）**：
3. 有没有格式正确的样本论文？→ 有的话用 parse_spec.py 提取规则
4. 有没有学校的格式规范文档？→ 有的话读取覆盖默认值
5. 参考文献格式标准 → 询问是否有 .csl 文件，参见下方"CSL 引用标准引导"
6. 有没有 .bib 文件？→ 参见下方"Better BibTeX 导出引导"
7. 需不需要引文校对？→ 如有 .bib，PDF 路径自动从 file 字段提取

#### CSL 引用标准引导

**什么是 .csl 文件？**
.csl（Citation Style Language）是一种 XML 格式的引用样式定义文件。它规定了参考文献的作者、标题、日期、页码等字段如何排版——比如作者名是否大写、书名用斜体还是书名号、页码用 p. 还是 pp.。学术期刊通常会提供 .csl 文件。

**话术**：
> "参考文献格式有 .csl 样式文件吗？.csl 是引用格式的定义文件，很多学术期刊和 Zotero 社区都提供。
> 如果没有，我可以帮你用 zotero-csl-builder 创建一个，或者你去 https://www.zotero.org/styles 搜索下载。
> 没有 .csl 的话，默认使用 CNU《外国文学评论》2024修订版格式。"

如果用户没有 .csl 且想要自定义格式：
- 推荐去 https://www.zotero.org/styles 搜索（有 10000+ 种期刊样式）
- 或触发 zotero-csl-builder 技能帮助创建

#### Better BibTeX .bib 导出引导

**什么是 .bib 文件？**
.bib 是 BibLaTeX/BibTeX 格式的参考文献数据库文件。Zotero 用户安装 Better BibTeX 插件后可以导出 .bib 文件，里面包含每条文献的完整元数据（作者、标题、出版社、年份等），还包含 PDF 文件路径（`file` 字段）。

**导出步骤（向用户说明）**：
> "如果你用 Zotero 管理参考文献，可以导出 .bib 文件来丰富引用信息和验证引文：
> 1. 安装 Better BibTeX 插件：https://github.com/retorquere/zotero-better-bibtex
> 2. 在 Zotero 中选择要导出的文献库/集合
> 3. 右键 → 导出文献库 → 选择 'Better BibTeX' 格式
> 4. 勾选 '保持更新'（Keep updated），这样你新增文献时 .bib 会自动更新
> 5. 保存为 .bib 文件，告诉我路径即可"

**.bib 中的 PDF 路径**：Better BibTeX 会自动在 `file` 字段中记录 PDF 路径：
```bibtex
@book{foucault1975,
  title = {规训与惩罚},
  author = {福柯},
  file = {:/Users/me/Zotero/storage/ABC123/规训与惩罚.pdf:application/pdf}
}
```
引文校准时自动从 .bib 提取 PDF 路径，无需用户手动指定。

#### 引文校验数据源引导

如果用户需要引文校验，逐步引导配置数据源，注明风险：

> "引文校验可以查询外部学术数据库验证引用的真实性。以下数据源供选择：
>
> **免费且无需配置（推荐先试这些）**：
> - OpenAlex — 2.5亿+ 论文，免费 API，最全面
> - CrossRef — DOI 元数据查询，免费
> - Semantic Scholar — AI 驱动的学术搜索，免费
>
> **需要配置（有风险）**：
> - 知网 CNKI — 需要从浏览器复制 Cookie，⚠️ 有封禁风险，建议用小号
> - 万方 Wanfang — 公开搜索页，无需配置，但反爬机制可能变化
>
> 你想用哪些数据源？默认用 OpenAlex + CrossRef，够覆盖大部分英文文献。"

**话术示例**（完整引导）：
> "收到！开始检查前确认几个事情：
> 1. 这篇是期刊论文还是学位论文？（默认期刊论文）
> 2. 参考文献格式有 .csl 文件吗？没有的话默认用 CNU 外国文学评论格式
> 3. 用 Zotero 管理参考文献吗？可以导出 .bib 文件来丰富引用信息
> 4. 需不需要验证引用的真实性？"

### Step 1：环境准备 + 规范确定

```
决策树：
  用户提供样本 .docx？
    ├─ 是 → parse_spec.py 提取 spec.json → 用 --spec 参数
    └─ 否 → 用户提供规范文档？
              ├─ 是 → Read 读取，提取数值覆盖默认值
              └─ 否 → 使用内置 UNIVERSAL_CHECKS 默认值
```

```bash
# 如有样本
python scripts/parse_spec.py "样本.docx" --output spec.json

# 确认依赖
pip install python-docx lxml
```

### Step 2：格式检测

根据 Step 0 确定的模式运行检测：

```bash
# 学位论文模式（检测封面/摘要/目录/正文/参考文献/致谢/附录）
python scripts/check_format.py "论文.docx" --mode thesis --output check_result.json

# 期刊论文模式（检测标题页/摘要/关键词/正文/参考文献/脚注）
python scripts/check_format.py "论文.docx" --mode journal --output check_result.json

# 如有 spec.json
python scripts/check_format.py "论文.docx" --mode thesis --spec spec.json --output check_result.json
```

### Step 3：问题诊断 + 用户确认

读取 `check_result.json` 的 `issues` 字段，向用户展示：

| 级别 | 含义 | 典型问题 |
|------|------|---------|
| P1（必须修复） | 格式明显错误 | 页边距偏差>2mm、标题字号错误、行距错误 |
| P2（应当修复） | 不规范 | 缩进单位不对、引用标点错误、表格非三线表 |
| P3（建议修复） | 需人工确认 | 页眉横线、目录需手动更新、引用格式需核实 |

**呈现格式**：展示问题数量分布表格，**询问用户是否继续修复**。等待用户确认后再执行 Step 4。

### Step 4：自动修复

用户确认后：

```bash
python scripts/fix_format.py "论文.docx" --mode thesis --output "论文_repaired.docx"

# 如有 spec.json
python scripts/fix_format.py "论文.docx" --spec spec.json --output "论文_repaired.docx"
```

**自动修复项**：页边距、正文字体字号、标题字体字号、行距、首行缩进、章节标题格式、参考文献格式

**无法自动修复（告知用户）**：页眉横线、目录更新、图片/公式位置、脚注格式

### Step 5：引文验证（可选，根据 Step 0 第 6、7 问决定）

```
决策树：
  用户有 .bib 文件？
    ├─ 是 → verify_local.py 本地验证 + 自动提取 PDF 路径
    │       用户需要引文校对？
    │         ├─ 是 → verify_quotes.py（PDF 路径从 .bib 的 file 字段自动获取）
    │         └─ 否 → 跳过引文校对
    └─ 否 → 跳过本地验证

  用户需要外部验证？
    ├─ 是 → 引导配置数据源（见下方"数据源配置引导"）
    └─ 否 → 跳过外部验证

  cross_check.py 交叉校验（始终执行）
```

```bash
# Step 5a: 本地验证（如有 .bib）
python scripts/verify_local.py "论文.docx" --bib refs.bib --output local_result.json

# Step 5b: 外部验证（如用户需要，逐步引导配置）
python scripts/verify_external.py "论文.docx" --sources openalex,crossref --output ext_result.json

# Step 5c: 交叉校验（始终执行）
python scripts/cross_check.py "论文.docx" --output cross_check.json

# Step 5d: 引文逐字核对（PDF 路径从 .bib 自动提取，或用户提供）
python scripts/verify_quotes.py "论文.docx" --sources "书名.pdf:1-100" --output quote_result.json

# 仅预览提取的引文（不做核对）
python scripts/verify_quotes.py "论文.docx" --extract-only
```

**引文逐字核对**：从论文正文中提取直接引文（引号内容），与出处 PDF 原文逐字比对。
- 一致：引文与原文完全匹配
- 基本一致：相似度 >= 78%，可能存在标点或空格差异
- 疑似不一致：相似度 55-78%，建议人工核实
- 未定位到出处：在 PDF 中未找到相似文本

#### 数据源配置引导

如果用户需要外部引文验证，逐步引导配置：

> "引文校验可以查询外部学术数据库验证引用的真实性。你想用哪些？
>
> **免费且无需配置（推荐）**：
> - OpenAlex — 2.5亿+ 论文，覆盖最全面
> - CrossRef — DOI 元数据查询
> - Semantic Scholar — AI 驱动的学术搜索
>
> **需要配置**：
> - 知网 CNKI — 需要从浏览器复制 Cookie，⚠️ 有账号封禁风险，建议用小号或访客模式
> - 万方 Wanfang — 公开搜索页，无需配置，但反爬机制可能变化
>
> 默认用 OpenAlex + CrossRef，覆盖大部分英文文献。中文文献建议用万方。"

向用户展示验证结果：匹配率、未匹配列表、孤立引用、未使用文献、引文核对状态。

### Step 6：报告生成

```bash
python scripts/generate_report.py "论文.docx" "论文_repaired.docx" "检测报告.docx" \
    --format-issues check_result.json
```

生成 A4 横向 .docx 报告，包含：
1. **格式问题**：6 列对比表（序号/位置/检查项/规范要求/当前值/修复建议）
2. **修复记录**：before → after 对比表
3. **参考文献报告**：P1/P2/P3 标记
4. **脚注报告**：脚注引用格式检测
5. **引文验证报告**（如有）：本地匹配 + 外部验证 + 交叉校验结果
6. **统计摘要**：各类别问题数量

向用户报告输出文件路径，提醒在 Word 中打开修复后的文件后手动更新目录（Ctrl+A → F9）。

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
