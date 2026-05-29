# 借鉴 cn-paper-typesetter 的 paper-format 优化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把 cn-paper-typesetter（纯提示词技能）相对 paper-format 多出来的两项"真本事"补进 paper-format：① 真正的"假排版"检测与清理（空行堆叠、手工空格缩进/居中、段内手工换行）；② 真正的视觉渲染复核（用已安装的 LibreOffice 把 docx 渲染成页图），并修正 Step 7 里"用 python-docx 渲染成图片"的不实表述。

**Architecture:** 两个新增独立脚本 + 最小集成。`check_fake_layout.py` 是纯段落分析（高度可 TDD）；其清理逻辑作为 `clean_fake_layout()` 供 `fix_format` 调用。`render_preview.py` 调 `soffice --headless --convert-to pdf` 再把 PDF 转 PNG（pdftoppm/pdf2image 优先，缺则返回 PDF 并提示），失败时优雅降级。两者都接入统一 CLI `paper_format.py` 与工作流。

**Tech Stack:** Python 3.9+、python-docx、lxml、LibreOffice（soffice，已确认安装于 `C:/Program Files/LibreOffice/program/soffice.exe`）、可选 poppler(`pdftoppm`)/`pdf2image`。

---

## 背景：与 cn-paper-typesetter 的差异分析

| 维度 | cn-paper-typesetter (manuscript-formatting) | paper-format（现状） |
|------|---------|---------|
| 形态 | 纯提示词（177 行，零脚本） | 脚本驱动（23 脚本 + 16 类规则引擎 + 引文验证 + CSL 流水线）|
| 提示理念（工作边界/执行方案/符号审查/说明口径）| 原创 | **已吸收**，措辞高度雷同 |
| **"假排版"清理** | 明确要求清掉"空行、手工空格、回车堆样式" | **缺失**——无任何检测/清理 |
| **视觉渲染复核** | 委托"文档技能渲染流程，把文档渲染成页图逐页检查" | **不实**——Step 7 称"用 python-docx 渲染成图片"，但 python-docx 无渲染能力 |

→ 只需补 ① 假排版 ② 真实渲染。其余 cn-paper-typesetter 的理念 paper-format 已覆盖，无需重复。

---

## File Structure

- **Create** `scripts/check_fake_layout.py` — 假排版检测（纯分析）+ `clean_fake_layout(doc)` 清理函数 + CLI。单一职责：识别并清理"假排版"。
- **Create** `scripts/render_preview.py` — docx→PDF→PNG 渲染复核 + CLI。单一职责：把文档渲染成可视页图。
- **Create** `tests/test_fake_layout.py` — 假排版检测/清理的单元测试。
- **Create** `tests/test_render_preview.py` — 渲染的集成测试（无 soffice 时跳过）。
- **Modify** `scripts/fix_format.py` — 在保存前调用 `clean_fake_layout(doc)`，记入 repair_records。
- **Modify** `paper_format.py` — 新增 `fake-layout` 与 `render` 两个子命令分派。
- **Modify** `SKILL.md` — 新增 Step 4.6（假排版清理），重写 Step 7（用真实渲染替代不实的 python-docx 渲染说法），脚本参考表补两行。

---

## Task 1: 假排版检测 `check_fake_layout.py`

**Files:**
- Create: `scripts/check_fake_layout.py`
- Test: `tests/test_fake_layout.py`

- [ ] **Step 1: 写失败测试**

```python
# tests/test_fake_layout.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from scripts.check_fake_layout import detect_fake_layout


def _doc_to_tmp(doc, tmp_path):
    p = tmp_path / "t.docx"
    doc.save(str(p))
    return str(p)


def test_detects_consecutive_blank_paragraphs(tmp_path):
    doc = Document()
    doc.add_paragraph("正文第一段。")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("")          # 3 连续空段
    doc.add_paragraph("正文第二段。")
    res = detect_fake_layout(_doc_to_tmp(doc, tmp_path))
    blanks = [i for i in res["issues"] if i["type"] == "consecutive_blank"]
    assert len(blanks) == 1
    assert blanks[0]["count"] == 3


def test_detects_leading_manual_spaces(tmp_path):
    doc = Document()
    doc.add_paragraph("　　这段用了两个全角空格冒充首行缩进。")  # 全角空格
    doc.add_paragraph("    这段用了四个半角空格。")
    res = detect_fake_layout(_doc_to_tmp(doc, tmp_path))
    leads = [i for i in res["issues"] if i["type"] == "manual_indent_spaces"]
    assert len(leads) == 2


def test_detects_intra_paragraph_break(tmp_path):
    doc = Document()
    p = doc.add_paragraph("第一行")
    br = OxmlElement("w:br"); p.runs[0]._r.append(br)
    p.add_run("被手工换行挤到第二行")
    res = detect_fake_layout(_doc_to_tmp(doc, tmp_path))
    assert any(i["type"] == "intra_para_break" for i in res["issues"])


def test_clean_doc_has_no_issues(tmp_path):
    doc = Document()
    doc.add_paragraph("正常段落一。")
    doc.add_paragraph("正常段落二。")
    res = detect_fake_layout(_doc_to_tmp(doc, tmp_path))
    assert res["total"] == 0
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd C:/Users/Administrator/.claude/skills/paper-format && python -m pytest tests/test_fake_layout.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.check_fake_layout'`

- [ ] **Step 3: 实现 `check_fake_layout.py`（检测部分）**

```python
#!/usr/bin/env python3
"""check_fake_layout.py - 检测并清理"假排版"。

"假排版"指用排版无关的手段冒充版式的做法，常见于学生稿：
  - 连续空段落堆叠制造垂直间距（应由段前/段后间距控制）
  - 段首手工空格冒充首行缩进（应由 first_line_indent 控制）
  - 段内手工换行(w:br)冒充分段或行距
  - 行尾游离空格
这些会让样式系统失效、跨设备错位，必须在统一样式前清掉。
"""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

_THIS = Path(__file__).resolve()
if str(_THIS.parents[1]) not in sys.path:
    sys.path.insert(0, str(_THIS.parents[1]))

from docx import Document
from docx.oxml.ns import qn

_LEADING_SPACE = ("　", " ", "\t")  # 全角空格 / 半角空格 / 制表符


def _is_blank(p) -> bool:
    return not p.text.strip()


def _has_break(p) -> bool:
    return p._p.find(qn("w:r") + "/" + qn("w:br")) is not None or \
        len(p._p.findall(".//" + qn("w:br"))) > 0


def detect_fake_layout(docx_path: str) -> dict:
    """Return {file, total, issues:[{type, location, detail, count?, autofix_safe}]}."""
    doc = Document(docx_path)
    paras = doc.paragraphs
    issues = []

    # 1) 连续空段（>=2）
    run_len = 0
    run_start = 0
    for i, p in enumerate(paras + [None]):
        if p is not None and _is_blank(p):
            if run_len == 0:
                run_start = i
            run_len += 1
        else:
            if run_len >= 2:
                issues.append({
                    "type": "consecutive_blank", "location": f"段 {run_start}-{run_start + run_len - 1}",
                    "count": run_len, "detail": f"{run_len} 个连续空段，疑似手工制造间距",
                    "autofix_safe": True,
                })
            run_len = 0

    # 2) 段首手工空格 / 3) 段内手工换行 / 4) 行尾空格
    for i, p in enumerate(paras):
        t = p.text
        if not t.strip():
            continue
        if t[:1] in _LEADING_SPACE:
            n = len(t) - len(t.lstrip("".join(_LEADING_SPACE)))
            issues.append({
                "type": "manual_indent_spaces", "location": f"段 {i}",
                "detail": f"段首 {n} 个手工空格，疑似冒充首行缩进/居中：{t[:20]!r}",
                "autofix_safe": True,
            })
        if len(p._p.findall(".//" + qn("w:br"))) > 0:
            issues.append({
                "type": "intra_para_break", "location": f"段 {i}",
                "detail": f"段内含手工换行(w:br)：{t[:20]!r}",
                "autofix_safe": False,  # 可能是有意换行，需人工判断
            })
        if t != t.rstrip() and t.strip():
            issues.append({
                "type": "trailing_space", "location": f"段 {i}",
                "detail": "行尾游离空格", "autofix_safe": True,
            })

    return {"file": docx_path, "total": len(issues), "issues": issues}


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="检测假排版（空行堆叠/手工空格/段内换行/行尾空格）")
    ap.add_argument("docx")
    ap.add_argument("--output", "-o")
    args = ap.parse_args()
    res = detect_fake_layout(args.docx)
    if args.output:
        Path(args.output).write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"检测到 {res['total']} 处假排版 → {args.output}")
    else:
        print(f"检测到 {res['total']} 处假排版：")
        for it in res["issues"]:
            print(f"  [{it['type']}] {it['location']} {it['detail']}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: 运行测试确认通过**

Run: `python -m pytest tests/test_fake_layout.py -v`
Expected: PASS（4 passed）

- [ ] **Step 5: 提交**

```bash
git add scripts/check_fake_layout.py tests/test_fake_layout.py
git commit -m "feat(paper-format): 新增假排版检测 check_fake_layout（借鉴 cn-paper-typesetter）"
```

---

## Task 2: 假排版清理 `clean_fake_layout()` + 接入 fix_format

**Files:**
- Modify: `scripts/check_fake_layout.py`（追加 `clean_fake_layout`）
- Modify: `scripts/fix_format.py`（保存前调用）
- Test: `tests/test_fake_layout.py`（追加清理测试）

- [ ] **Step 1: 写失败测试（追加到 tests/test_fake_layout.py 末尾）**

```python
from scripts.check_fake_layout import clean_fake_layout


def test_clean_collapses_blanks_and_strips_leading_spaces(tmp_path):
    doc = Document()
    doc.add_paragraph("正文第一段。")
    doc.add_paragraph(""); doc.add_paragraph(""); doc.add_paragraph("")
    doc.add_paragraph("　　第二段带全角空格。")
    n = clean_fake_layout(doc)
    out = tmp_path / "cleaned.docx"; doc.save(str(out))
    d2 = Document(str(out))
    texts = [p.text for p in d2.paragraphs]
    # 连续空段被折叠为至多 1 个
    assert "\n".join(texts).count("\n\n\n") == 0
    # 段首全角空格被去除
    assert any(t.startswith("第二段") for t in texts)
    assert n >= 2  # 至少清理了 空段折叠 + 段首空格
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_fake_layout.py::test_clean_collapses_blanks_and_strips_leading_spaces -v`
Expected: FAIL — `ImportError: cannot import name 'clean_fake_layout'`

- [ ] **Step 3: 实现 `clean_fake_layout`（追加到 check_fake_layout.py，main 之前）**

```python
def clean_fake_layout(doc) -> int:
    """In-place 清理可安全自动处理的假排版，返回处理动作数。

    只做高确定性的清理：折叠连续空段为单个、去段首手工空格、去行尾空格。
    段内手工换行(w:br)不自动删（可能有意），仅由 detect 标记待人工确认。
    """
    actions = 0
    paras = doc.paragraphs

    # 1) 折叠连续空段：保留每段连续空白的第 1 个，删除其余
    prev_blank = False
    for p in list(paras):
        blank = not p.text.strip()
        if blank and prev_blank:
            p._p.getparent().remove(p._p)
            actions += 1
        prev_blank = blank

    # 2) 段首手工空格 + 3) 行尾空格：改写首/尾 run 文本
    for p in doc.paragraphs:
        if not p.text.strip() or not p.runs:
            continue
        first = p.runs[0]
        stripped = first.text.lstrip("　 \t")
        if stripped != first.text:
            first.text = stripped
            actions += 1
        last = p.runs[-1]
        rstripped = last.text.rstrip("　 \t")
        if rstripped != last.text:
            last.text = rstripped
            actions += 1

    return actions
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_fake_layout.py -v`
Expected: PASS（5 passed）

- [ ] **Step 5: 接入 fix_format（在 `scripts/fix_format.py` 的"Reference font normalization"块之后、Save 之前插入）**

```python
    # ---------------------------------------------------------------
    # 假排版清理（借鉴 cn-paper-typesetter）：在统一样式之后、保存之前清掉
    # 空段堆叠、段首手工空格、行尾游离空格——它们会让样式系统失效。
    # ---------------------------------------------------------------
    try:
        from scripts.check_fake_layout import clean_fake_layout
        fake_actions = clean_fake_layout(doc)
        if fake_actions:
            records.append({"category": "paragraphs", "item": "fake_layout_cleanup",
                            "location": "全文", "note": f"清理假排版 {fake_actions} 处（空段折叠/段首空格/行尾空格）"})
    except Exception as e:
        records.append({"category": "paragraphs", "item": "fake_layout_cleanup", "error": str(e)})
```

- [ ] **Step 6: 回归 + 提交**

Run: `python -m pytest tests/test_fake_layout.py -v && python -c "import py_compile; py_compile.compile('scripts/fix_format.py', doraise=True)"`
Expected: PASS + 无编译错误

```bash
git add scripts/check_fake_layout.py scripts/fix_format.py tests/test_fake_layout.py
git commit -m "feat(paper-format): 假排版清理接入 fix_format（折叠空段/去手工空格）"
```

---

## Task 3: 真实视觉渲染 `render_preview.py`

**Files:**
- Create: `scripts/render_preview.py`
- Test: `tests/test_render_preview.py`

- [ ] **Step 1: 写集成测试（无 soffice 时跳过）**

```python
# tests/test_render_preview.py
import shutil
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pytest
from docx import Document
from scripts.render_preview import find_soffice, render_to_pdf


def test_find_soffice_returns_path_or_none():
    # 不应抛异常；返回 str 或 None
    res = find_soffice()
    assert res is None or isinstance(res, str)


def test_render_to_pdf_when_soffice_available(tmp_path):
    soffice = find_soffice()
    if not soffice:
        pytest.skip("soffice 不可用，跳过渲染集成测试")
    doc = Document(); doc.add_paragraph("渲染测试 render test 123。")
    src = tmp_path / "src.docx"; doc.save(str(src))
    pdf = render_to_pdf(str(src), str(tmp_path), soffice_path=soffice)
    assert pdf is not None and Path(pdf).exists() and Path(pdf).suffix == ".pdf"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_render_preview.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'scripts.render_preview'`

- [ ] **Step 3: 实现 `render_preview.py`**

```python
#!/usr/bin/env python3
"""render_preview.py - 把 docx 渲染成 PDF/页图，用于真正的视觉复核。

paper-format 旧 Step 7 声称"用 python-docx 渲染成图片"，但 python-docx 无渲染能力。
本脚本用已安装的 LibreOffice(soffice) 把 docx 转 PDF，再尽力转成 PNG 页图
（优先 pdftoppm，其次 pdf2image，都没有则返回 PDF 并提示）。soffice 缺失时优雅降级。
"""
from __future__ import annotations
import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

_WIN_SOFFICE = [
    r"C:/Program Files/LibreOffice/program/soffice.exe",
    r"C:/Program Files (x86)/LibreOffice/program/soffice.exe",
]


def find_soffice(explicit: str | None = None) -> str | None:
    if explicit and Path(explicit).exists():
        return explicit
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for p in _WIN_SOFFICE:
        if Path(p).exists():
            return p
    return None


def render_to_pdf(docx_path: str, out_dir: str, soffice_path: str | None = None) -> str | None:
    soffice = find_soffice(soffice_path)
    if not soffice:
        return None
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True, timeout=120,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    pdf = Path(out_dir) / (Path(docx_path).stem + ".pdf")
    return str(pdf) if pdf.exists() else None


def pdf_to_pngs(pdf_path: str, out_dir: str, dpi: int = 120) -> list[str]:
    """PDF→PNG 页图。优先 pdftoppm(poppler)，否则 pdf2image，否则返回空列表。"""
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    stem = Path(pdf_path).stem
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf_path,
                        str(Path(out_dir) / stem)], check=True, timeout=120,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return sorted(str(p) for p in Path(out_dir).glob(f"{stem}*.png"))
    try:
        from pdf2image import convert_from_path
        imgs = convert_from_path(pdf_path, dpi=dpi)
        out = []
        for i, im in enumerate(imgs, 1):
            fp = Path(out_dir) / f"{stem}-{i}.png"
            im.save(str(fp)); out.append(str(fp))
        return out
    except Exception:
        return []


def render_preview(docx_path: str, out_dir: str, soffice_path: str | None = None) -> dict:
    pdf = render_to_pdf(docx_path, out_dir, soffice_path)
    if not pdf:
        return {"ok": False, "reason": "LibreOffice(soffice) 不可用，无法渲染；请安装或在 Word 中人工复核",
                "pdf": None, "pngs": []}
    pngs = pdf_to_pngs(pdf, out_dir)
    return {"ok": True, "pdf": pdf, "pngs": pngs,
            "note": "" if pngs else "已生成 PDF；未找到 pdftoppm/pdf2image，未能转 PNG，可直接查看 PDF"}


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="把 docx 渲染成 PDF/PNG 用于视觉复核")
    ap.add_argument("docx")
    ap.add_argument("--out-dir", "-o", default=".")
    ap.add_argument("--soffice", help="soffice 可执行文件路径（默认自动探测）")
    args = ap.parse_args()
    res = render_preview(args.docx, args.out_dir, args.soffice)
    if res["ok"]:
        print(f"PDF: {res['pdf']}")
        print(f"PNG 页图: {len(res['pngs'])} 张" + (f"（{res['note']}）" if res['note'] else ""))
    else:
        print(f"渲染失败: {res['reason']}")
        sys.exit(1)


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_render_preview.py -v`
Expected: PASS（soffice 已安装，2 passed；无 soffice 则 1 passed 1 skipped）

- [ ] **Step 5: 端到端冒烟（用真实修复产物）**

Run: `python scripts/render_preview.py "test_output/迷楼_repaired.docx" --out-dir test_output/_preview`
Expected: 打印 `PDF: ...` 且 `test_output/_preview/` 下生成 PDF（若有 poppler 则另有 PNG）

- [ ] **Step 6: 提交**

```bash
git add scripts/render_preview.py tests/test_render_preview.py
git commit -m "feat(paper-format): 新增真实视觉渲染 render_preview（LibreOffice docx→PDF→PNG）"
```

---

## Task 4: 接入 CLI 与 SKILL.md（修正不实表述）

**Files:**
- Modify: `paper_format.py`（新增 `fake-layout`、`render` 子命令）
- Modify: `SKILL.md`（新增 Step 4.6；重写 Step 7；脚本表补两行）

- [ ] **Step 1: paper_format.py 新增分派（在 `content` 分支之后插入）**

```python
    elif command == "fake-layout":
        from scripts.check_fake_layout import main as fake_main
        fake_main()
    elif command == "render":
        from scripts.render_preview import main as render_main
        render_main()
```

- [ ] **Step 2: 更新 paper_format.py 的"Available:"提示行**

把 `Available:` 那行加入 `fake-layout` 和 `render` 两个命令名（紧跟 `content` 之后）。

- [ ] **Step 3: 验证 CLI**

Run: `python paper_format.py fake-layout --help && python paper_format.py render --help`
Expected: 两条均打印各自的 usage，无 `Unknown command`

- [ ] **Step 4: SKILL.md 新增 Step 4.6（在 Step 4 符号审查之后、Step 5 之前）**

```markdown
### Step 4.6：假排版清理

"假排版"指用排版无关手段冒充版式（空段堆叠、段首手工空格、段内手工换行、行尾空格），
会让样式系统失效、跨设备错位，必须在统一样式前清掉（借鉴 cn-paper-typesetter）。

```bash
python scripts/check_fake_layout.py "论文_copy.docx" --output "输出目录/fake_layout.json"
```

`fix_format` 会在统一样式后自动清理高确定性项（折叠连续空段、去段首/行尾手工空格）；
段内手工换行(w:br)只标记不自动删（可能是有意换行，需人工确认）。
```

- [ ] **Step 5: SKILL.md 重写 Step 7（替换"用 python-docx 将文档按页渲染成图片"的不实表述）**

把 Step 7 复核方式第 1 条由：
`1. 用 python-docx 将修复后的文档按页渲染成图片（或提取关键段落的格式属性）`
改为：

```markdown
1. 用 `render_preview.py` 通过 LibreOffice 把修复后的文档渲染成 PDF/PNG 页图，逐页肉眼检查
   （python-docx 无法渲染图片，必须用 LibreOffice；soffice 缺失时降级为提取关键段落格式属性 + 提示人工在 Word 中复核）：

```bash
python scripts/render_preview.py "输出目录/论文_repaired.docx" --out-dir "输出目录/_preview"
```
```

- [ ] **Step 6: SKILL.md 脚本参考表补两行（在 check_content 行附近）**

```markdown
| `check_fake_layout.py` | 假排版检测/清理（空段堆叠、手工空格、段内换行） | `论文.docx --output result.json` |
| `render_preview.py` | 视觉渲染复核（LibreOffice docx→PDF→PNG） | `论文.docx --out-dir _preview` |
```

- [ ] **Step 7: 全量编译 + 提交**

Run: `python -c "import py_compile,glob; [py_compile.compile(f,doraise=True) for f in glob.glob('scripts/*.py')+['paper_format.py']]; print('OK')"`
Expected: `OK`

```bash
git add paper_format.py SKILL.md
git commit -m "docs(paper-format): 接入 fake-layout/render 子命令，修正 Step 7 不实的渲染表述"
```

---

## Self-Review

**1. Spec coverage：**
- ① 假排版 → Task 1（检测）+ Task 2（清理+接入 fix_format）✅
- ② 真实渲染 → Task 3（render_preview）+ Task 4 Step 5（修正 Step 7 不实表述）✅
- CLI/文档接入 → Task 4 ✅
- cn-paper-typesetter 其余理念（执行方案/符号审查/说明口径）paper-format 已覆盖，无需新增 ✅

**2. Placeholder scan：** 各步均含完整可运行代码与确切命令、预期输出，无 TBD/TODO/"类似上文"。✅

**3. Type consistency：** `detect_fake_layout`/`clean_fake_layout`/`find_soffice`/`render_to_pdf`/`pdf_to_pngs`/`render_preview` 在测试与实现、CLI 分派中签名一致；repair_records 用既有字段 `category/item/location/note`。✅

## 风险与降级
- **render PNG 依赖 poppler/pdf2image**：缺失时只产 PDF 并提示——不阻断流程（Task 3 已降级）。
- **soffice 缺失**：render_preview 返回 ok=False 提示人工复核；fix_format 不依赖渲染，不受影响。
- **clean_fake_layout 删空段**：只折叠"连续"空段保留 1 个，不会清掉所有空段，降低误伤；段内换行只标记不删。
