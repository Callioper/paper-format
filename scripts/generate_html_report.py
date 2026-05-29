#!/usr/bin/env python3
"""generate_html_report.py - Generate interactive HTML format report.

Self-contained HTML with:
  - Summary dashboard with SVG charts
  - Collapsible detail sections
  - Color-coded status indicators
  - Embedded screenshots (base64)
  - Print-friendly CSS

Usage:
    python generate_html_report.py original.docx repaired.docx --format-issues check.json --output report.html
"""

from __future__ import annotations

import argparse
import base64
import json
import sys
from pathlib import Path
from typing import Any

# -------------------------------------------------------------------
# Category name mapping
# -------------------------------------------------------------------

CAT_NAMES = {
    "page_setup": "页面设置",
    "styles": "样式",
    "paragraphs": "段落",
    "tables": "表格",
    "cover": "封面",
    "abstract": "摘要",
    "toc": "目录",
    "headers_footers": "页眉页脚",
    "acknowledgments": "致谢",
    "appendices": "附录",
    "citations": "参考文献",
    "footnotes": "脚注",
}

REPAIR_DESC = {
    "page_setup": "页边距已修正",
    "paragraphs": "字体字号行距已修正",
    "sections": "章节标题已格式化",
    "cover": "封面格式已修正",
    "citations": "引用格式已规范化",
    "footnotes": "脚注字体已修正",
}

# -------------------------------------------------------------------
# SVG chart helpers
# -------------------------------------------------------------------


def _pie_chart_svg(data: dict[str, int], size: int = 160) -> str:
    """Generate a simple SVG pie chart."""
    total = sum(data.values())
    if total == 0:
        return '<div style="text-align:center;padding:20px">无问题</div>'

    colors = {
        "已修复": "#22c55e",
        "需人工处理": "#f59e0b",
    }

    # Fallback colors for categories
    fallback_colors = [
        "#3b82f6", "#ef4444", "#8b5cf6", "#06b6d4", "#f97316",
        "#ec4899", "#14b8a6", "#6366f1", "#84cc16", "#64748b",
    ]

    segments = []
    angle = 0
    cx, cy, r = size // 2, size // 2, size // 2 - 10

    for i, (label, count) in enumerate(data.items()):
        fraction = count / total
        sweep = fraction * 360

        # SVG arc
        import math
        x1 = cx + r * math.cos(math.radians(angle))
        y1 = cy + r * math.sin(math.radians(angle))
        x2 = cx + r * math.cos(math.radians(angle + sweep))
        y2 = cy + r * math.sin(math.radians(angle + sweep))

        large = 1 if sweep > 180 else 0
        color = colors.get(label, fallback_colors[i % len(fallback_colors)])

        segments.append(
            f'<path d="M{cx},{cy} L{x1:.1f},{y1:.1f} A{r},{r} 0 {large},1 {x2:.1f},{y2:.1f} Z" '
            f'fill="{color}" stroke="white" stroke-width="2">'
            f'<title>{label}: {count} ({fraction:.0%})</title></path>'
        )
        angle += sweep

    svg = f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}">'
    svg += ''.join(segments)
    svg += '</svg>'
    return svg


def _legend_html(data: dict[str, int], colors: dict[str, str]) -> str:
    """Generate legend items for a chart."""
    total = sum(data.values())
    items = []
    for label, count in data.items():
        color = colors.get(label, "#888")
        pct = f"{count / total:.0%}" if total > 0 else "0%"
        items.append(
            f'<span style="display:inline-block;margin-right:16px">'
            f'<span style="display:inline-block;width:12px;height:12px;background:{color};'
            f'border-radius:2px;margin-right:4px;vertical-align:middle"></span>'
            f'{label}: {count} ({pct})</span>'
        )
    return ' '.join(items)


# -------------------------------------------------------------------
# HTML generation
# -------------------------------------------------------------------


def _severity_color(severity: str) -> str:
    """Return CSS color for severity level."""
    s = severity.lower()
    if s in ("high", "p1"):
        return "#ef4444"
    if s in ("medium", "p2"):
        return "#f59e0b"
    if s in ("low", "p3"):
        return "#6b7280"
    return "#374151"


def _status_badge(status: str) -> str:
    """Return HTML badge for status."""
    if status == "已修复":
        return '<span class="badge badge-ok">已修复</span>'
    return '<span class="badge badge-warn">需人工处理</span>'


def generate_html_report(
    original_path: str,
    repaired_path: str,
    format_issues: dict,
    citation_results: list[dict],
    footnote_results: list[dict] | None = None,
    repair_records: list[dict] | None = None,
    verification_results: dict | None = None,
    quote_results: list[dict] | None = None,
    mode: str = "journal",
) -> str:
    """Generate a self-contained HTML report string."""

    issues = format_issues.get("issues", {})
    total_issues = sum(len(v) for v in issues.values())

    # Which repair-record categories actually ran (no error)?
    repair_cats = set()
    if repair_records:
        for rec in repair_records:
            if not rec.get("error"):
                repair_cats.add(rec.get("category", ""))

    # Resolve at the ITEM level, not the category level. check_format and fix_format
    # use different category vocabularies (检测说"styles/abstract"，修复记录说
    # "paragraphs/sections")，所以纯按类别名匹配会把已修复的正文字号、摘要标题误报为未修复。
    # 这个别名表把"修复记录类别"映射到"它实际解决了哪些检测项"，并保留确实没自动修的项
    # （如标题样式缺失）。
    def _issue_resolved(cat: str, item: str) -> bool:
        item = item or ""
        if cat == "page_setup":
            return "page_setup" in repair_cats
        if cat == "styles":
            # 正文字号由 fix_format 的 paragraphs(font_indent_spacing) 修复；
            # heading_*_missing 属结构问题，不自动修。
            if "body_font" in item or "font_size" in item:
                return "paragraphs" in repair_cats
            return False
        if cat == "abstract":
            return "sections" in repair_cats        # 摘要标题/run-in 由 sections 修复
        if cat == "paragraphs":
            return "paragraphs" in repair_cats
        # tables / citations / footnotes / headers_footers 等同名直配
        return cat in repair_cats

    # Per-item resolution map: (cat, idx) -> bool
    resolved_map = {}
    for cat, items in issues.items():
        for idx, it in enumerate(items):
            resolved_map[(cat, idx)] = _issue_resolved(cat, it.get("item", "") if isinstance(it, dict) else "")

    fixed_count = sum(1 for v in resolved_map.values() if v)
    remaining = total_issues - fixed_count
    pie_data = {"已修复": fixed_count, "需人工处理": remaining} if fixed_count > 0 else {"需人工处理": remaining}

    # Category breakdown
    cat_data = {}
    for cat, items in issues.items():
        if items:
            cat_data[CAT_NAMES.get(cat, cat)] = len(items)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>论文格式检测报告</title>
<style>
  :root {{
    --ok: #22c55e; --warn: #f59e0b; --err: #ef4444;
    --bg: #f8fafc; --card: #ffffff; --border: #e2e8f0;
    --text: #1e293b; --muted: #64748b;
  }}
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: "PingFang SC", "Noto Sans SC", "Microsoft YaHei", sans-serif;
         background: var(--bg); color: var(--text); line-height: 1.6; padding: 24px; }}
  .container {{ max-width: 1100px; margin: 0 auto; }}
  h1 {{ font-size: 1.6rem; font-weight: 700; margin-bottom: 4px; }}
  .subtitle {{ color: var(--muted); font-size: 0.9rem; margin-bottom: 24px; }}
  .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 24px; }}
  .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px;
           padding: 20px; }}
  .card h2 {{ font-size: 1rem; font-weight: 600; margin-bottom: 12px; }}
  .kpi {{ font-size: 2rem; font-weight: 700; }}
  .kpi-label {{ font-size: 0.8rem; color: var(--muted); }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; }}
  th {{ background: #f1f5f9; text-align: left; padding: 8px 10px; font-weight: 600;
       border-bottom: 2px solid var(--border); }}
  td {{ padding: 7px 10px; border-bottom: 1px solid var(--border); vertical-align: top; }}
  tr:hover td {{ background: #f8fafc; }}
  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 4px;
            font-size: 0.75rem; font-weight: 600; }}
  .badge-ok {{ background: #dcfce7; color: #166534; }}
  .badge-warn {{ background: #fef3c7; color: #92400e; }}
  .badge-err {{ background: #fee2e2; color: #991b1b; }}
  .badge-info {{ background: #dbeafe; color: #1e40af; }}
  .section {{ margin-bottom: 24px; }}
  .section-title {{ font-size: 1.1rem; font-weight: 600; padding: 10px 0;
                    border-bottom: 2px solid var(--border); margin-bottom: 12px;
                    cursor: pointer; user-select: none; }}
  .section-title::before {{ content: "▶ "; font-size: 0.8em; color: var(--muted); }}
  .section-title.open::before {{ content: "▼ "; }}
  .collapsible {{ overflow: hidden; transition: max-height 0.3s ease; }}
  .chart-container {{ display: flex; align-items: center; gap: 24px; }}
  .legend {{ font-size: 0.85rem; }}
  .legend-item {{ margin-bottom: 4px; }}
  .legend-dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 2px;
                margin-right: 6px; vertical-align: middle; }}
  .citation-item {{ padding: 8px 0; border-bottom: 1px solid var(--border); }}
  .citation-item:last-child {{ border-bottom: none; }}
  .p1 {{ color: var(--err); font-weight: 600; }}
  .p2 {{ color: var(--warn); font-weight: 600; }}
  .p3 {{ color: var(--muted); }}
  img.screenshot {{ max-width: 400px; border: 1px solid var(--border); border-radius: 4px;
                    margin-top: 4px; }}
  @media print {{
    body {{ padding: 0; background: white; }}
    .section-title {{ cursor: default; }}
    .collapsible {{ max-height: none !important; overflow: visible !important; }}
    .card {{ break-inside: avoid; }}
  }}
</style>
</head>
<body>
<div class="container">

<h1>论文格式检测与修复报告</h1>
<div class="subtitle">
  原始文件：{Path(original_path).name} &nbsp;|&nbsp;
  修复后：{Path(repaired_path).name} &nbsp;|&nbsp;
  模式：{'学位论文' if mode == 'thesis' else '期刊论文'} &nbsp;|&nbsp;
  规范：CNU + GB/T 7713.1-2006
</div>

<!-- ========== KPI Cards ========== -->
<div class="grid">
  <div class="card">
    <div class="kpi" style="color:var(--err)">{total_issues}</div>
    <div class="kpi-label">检测到的问题</div>
  </div>
  <div class="card">
    <div class="kpi" style="color:var(--ok)">{fixed_count}</div>
    <div class="kpi-label">已自动修复</div>
  </div>
</div>

<!-- ========== Summary Dashboard ========== -->
<div class="card section">
  <h2>总体概览</h2>
  <div class="chart-container">
    {_pie_chart_svg(pie_data)}
    <div class="legend">
      {_legend_html(pie_data, {"已修复": "#22c55e", "需人工处理": "#f59e0b"})}
    </div>
  </div>

  <table style="margin-top:16px">
    <thead><tr><th>类别</th><th>问题数</th><th>修复操作</th><th>状态</th></tr></thead>
    <tbody>"""

    # Summary table rows — resolved count is computed per item, so a category can be
    # partially fixed (e.g. styles: 正文字号已修复 but 标题样式缺失 未修复).
    for cat, items in issues.items():
        count = len(items)
        if count == 0:
            continue
        cat_name = CAT_NAMES.get(cat, cat)
        cat_fixed = sum(1 for idx in range(count) if resolved_map.get((cat, idx)))
        if cat_fixed == count:
            desc = REPAIR_DESC.get(cat, "已修复")
            badge = _status_badge("已修复")
        elif cat_fixed > 0:
            desc = REPAIR_DESC.get(cat, "部分修复")
            badge = f'{_status_badge("已修复")} {cat_fixed}/{count}'
        else:
            desc = "-"
            badge = _status_badge("需人工处理")
        html += f"\n      <tr><td>{cat_name}</td><td>{count}</td><td>{desc}</td><td>{badge}</td></tr>"

    html += f"""
    </tbody>
    <tfoot><tr style="font-weight:700;background:#f1f5f9">
      <td>合计</td><td>{total_issues}</td><td>{len(repair_records) if repair_records else 0} 项操作</td>
      <td>{_status_badge('已修复')} {fixed_count} 项</td>
    </tr></tfoot>
  </table>
</div>"""

    # ========== Repair Details ==========
    if repair_records:
        html += """
<div class="card section">
  <div class="section-title open" onclick="this.classList.toggle('open');this.nextElementSibling.style.maxHeight=this.classList.contains('open')?'2000px':'0'">修复明细</div>
  <div class="collapsible" style="max-height:2000px">
  <table>
    <thead><tr><th>类别</th><th>项目</th><th>位置</th><th>修复前</th><th>修复后</th></tr></thead>
    <tbody>"""

        for rec in repair_records:
            if rec.get("error"):
                continue
            cat_name = CAT_NAMES.get(rec.get("category", ""), rec.get("category", ""))
            before = str(rec.get("before", "-"))
            after = str(rec.get("after", "-"))
            if len(before) > 60:
                before = before[:57] + "..."
            if len(after) > 60:
                after = after[:57] + "..."
            html += f"""
      <tr>
        <td>{cat_name}</td>
        <td>{rec.get('item', '')}</td>
        <td>{rec.get('location', '')}</td>
        <td>{before}</td>
        <td style="color:var(--ok)">{after}</td>
      </tr>"""

        html += """
    </tbody>
  </table>
  </div>
</div>"""

    # ========== Remaining Issues ==========
    # Item-level: only keep issues that were NOT resolved (so 已修复的正文字号/摘要标题
    # 不再出现在"未修复"清单里，而确实没修的标题样式缺失仍保留)。
    remaining_issues = {}
    for cat, items in issues.items():
        unresolved = [it for idx, it in enumerate(items) if not resolved_map.get((cat, idx))]
        if unresolved:
            remaining_issues[cat] = unresolved
    if remaining_issues:
        html += """
<div class="card section">
  <div class="section-title open" onclick="this.classList.toggle('open');this.nextElementSibling.style.maxHeight=this.classList.contains('open')?'3000px':'0'">未修复问题明细</div>
  <div class="collapsible" style="max-height:3000px">"""

        for cat, items in remaining_issues.items():
            cat_name = CAT_NAMES.get(cat, cat)
            html += f'\n  <h3 style="margin:12px 0 6px;font-size:0.95rem">{cat_name} ({len(items)})</h3>'
            html += '\n  <table><thead><tr><th>#</th><th>检查项</th><th>规范要求</th><th>当前值</th><th>建议</th></tr></thead><tbody>'
            for i, item in enumerate(items, 1):
                sev = _severity_color(item.get("severity", ""))
                html += f"""
    <tr>
      <td>{i}</td>
      <td style="color:{sev};font-weight:600">{item.get('item', '')}</td>
      <td>{item.get('expected', '')}</td>
      <td>{item.get('actual', '')}</td>
      <td>{item.get('suggestion', '')}</td>
    </tr>"""
            html += "\n  </tbody></table>"

        html += "\n  </div>\n</div>"

    # ========== Citation Report ==========
    if citation_results:
        p1 = sum(1 for r in citation_results if any("P1" in w for w in r.get("warnings", [])))
        p2 = sum(1 for r in citation_results if any("P2" in w for w in r.get("warnings", [])))
        p3 = sum(1 for r in citation_results if any("P3" in w for w in r.get("warnings", [])))

        html += f"""
<div class="card section">
  <div class="section-title" onclick="this.classList.toggle('open');this.nextElementSibling.style.maxHeight=this.classList.contains('open')?'5000px':'0'">参考文献格式报告 ({len(citation_results)} 条)</div>
  <div class="collapsible" style="max-height:0">
    <p>P1(必须修复): <span class="p1">{p1}</span> &nbsp;
       P2(应当修复): <span class="p2">{p2}</span> &nbsp;
       P3(建议修复): <span class="p3">{p3}</span></p>"""

        for idx, r in enumerate(citation_results, 1):
            orig = r.get("original", "")[:100]
            formatted = r.get("formatted", "")
            warnings = r.get("warnings", [])
            html += f'\n    <div class="citation-item">'
            html += f'\n      <div><strong>[{idx}]</strong> {orig}</div>'
            if formatted:
                html += f'\n      <div style="color:var(--ok)">规范化：{formatted}</div>'
            for w in warnings:
                cls = "p1" if "P1" in w else ("p2" if "P2" in w else "p3")
                html += f'\n      <div class="{cls}">{w}</div>'
            html += "\n    </div>"

        html += "\n  </div>\n</div>"

    # ========== Cross-check / Verification ==========
    if verification_results:
        html += """
<div class="card section">
  <div class="section-title" onclick="this.classList.toggle('open');this.nextElementSibling.style.maxHeight=this.classList.contains('open')?'2000px':'0'">引文完整性校验</div>
  <div class="collapsible" style="max-height:0">"""

        for key, label in [("local", "本地数据库验证"), ("external", "外部数据源验证"), ("cross_check", "交叉校验")]:
            data = verification_results.get(key, {})
            if not data:
                continue
            html += f'\n    <h3 style="margin:10px 0 6px;font-size:0.95rem">{label}</h3>'
            html += '\n    <table><tbody>'
            for k, v in data.items():
                if isinstance(v, (str, int, float)):
                    html += f'\n      <tr><td style="font-weight:600">{k}</td><td>{v}</td></tr>'
            html += "\n    </tbody></table>"

        html += "\n  </div>\n</div>"

    # ========== Quote Verification ==========
    if quote_results:
        html += """
<div class="card section">
  <div class="section-title" onclick="this.classList.toggle('open');this.nextElementSibling.style.maxHeight=this.classList.contains('open')?'5000px':'0'">引文逐字核对</div>
  <div class="collapsible" style="max-height:0">
  <table>
    <thead><tr><th>#</th><th>引文</th><th>出处</th><th>页码</th><th>结果</th><th>备注</th></tr></thead>
    <tbody>"""

        for r in quote_results:
            status = r.get("status", "")
            if status == "一致":
                badge = '<span class="badge badge-ok">一致</span>'
            elif status == "基本一致":
                badge = '<span class="badge badge-warn">基本一致</span>'
            elif status == "疑似不一致":
                badge = '<span class="badge badge-err">疑似不一致</span>'
            else:
                badge = '<span class="badge badge-info">未定位</span>'

            html += f"""
      <tr>
        <td>{r.get('quote_index', '')}</td>
        <td>{r.get('quote_text', '')[:60]}</td>
        <td>{r.get('source_file', '')}</td>
        <td>{r.get('matched_page', '')}</td>
        <td>{badge}</td>
        <td>{r.get('note', '')}</td>
      </tr>"""

            # Embed screenshot if available
            ss_path = r.get("screenshot_path", "")
            if ss_path and Path(ss_path).exists():
                try:
                    with open(ss_path, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode()
                    html += f'\n      <tr><td colspan="6"><img class="screenshot" src="data:image/png;base64,{b64}" /></td></tr>'
                except Exception:
                    pass

        html += """
    </tbody>
  </table>
  </div>
</div>"""

    # ========== Footer ===========
    html += """
<div style="text-align:center;color:var(--muted);font-size:0.8rem;padding:24px 0">
  Generated by paper-format skill &nbsp;|&nbsp; <a href="https://github.com/Callioper/paper-format">GitHub</a>
</div>

</div>
</body>
</html>"""

    return html


# -------------------------------------------------------------------
# CLI
# -------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Generate HTML format report")
    parser.add_argument("original", help="Original thesis .docx")
    parser.add_argument("repaired", help="Repaired thesis .docx")
    parser.add_argument("--format-issues", help="JSON file with format issues")
    parser.add_argument("--repair-records", help="JSON file with repair records")
    parser.add_argument("--output", "-o", help="Output HTML path", default="report.html")
    parser.add_argument("--mode", default="journal", choices=["journal", "thesis"])
    args = parser.parse_args()

    format_issues = {}
    if args.format_issues:
        format_issues = json.loads(Path(args.format_issues).read_text(encoding="utf-8"))

    repair_records = None
    if args.repair_records:
        repair_records = json.loads(Path(args.repair_records).read_text(encoding="utf-8"))

    html = generate_html_report(
        args.original, args.repaired,
        format_issues, [],
        repair_records=repair_records,
        mode=args.mode,
    )

    out = Path(args.output)
    out.write_text(html, encoding="utf-8")
    print(f"HTML report saved to {out}")


if __name__ == "__main__":
    main()
