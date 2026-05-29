#!/usr/bin/env python3
"""check_language.py — 基于规则引擎的中文学术论文语言/符号检查。

读取 references/language_rules.yaml 中的规则配置，对 .docx 文件逐段扫描，
输出 JSON 格式的检查报告。

Usage:
    python scripts/check_language.py "论文.docx" --output language_result.json
    python scripts/check_language.py "论文.docx" --rules references/language_rules.yaml
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError:
    yaml = None

try:
    from docx import Document
except ImportError:
    print("需要 python-docx: pip install python-docx", file=sys.stderr)
    sys.exit(1)


# ── 默认规则（当 YAML 不可用时的硬编码回退） ───────────────

_DEFAULT_RULES = {
    "cjk_latin_spacing": {
        "enabled": True, "severity": "P2", "autofix_safe": True,
        "rules": [
            {"name": "missing_space_cjk_then_latin",
             "pattern": r"([一-鿿])([A-Za-z0-9])",
             "fix": r"\1 \2"},
            {"name": "missing_space_latin_then_cjk",
             "pattern": r"([A-Za-z0-9])([一-鿿])",
             "fix": r"\1 \2"},
        ],
    },
    "repeated_punctuation": {
        "enabled": True, "severity": "P2", "autofix_safe": False,
        "patterns": ["。。。。", "。。", "、、", "！！", "？？", ",,", r"\.\.\."],
    },
    "mixed_quote_style": {
        "enabled": True, "severity": "P2", "autofix_safe": False,
    },
    "bracket_mismatch": {
        "enabled": True, "severity": "P1", "autofix_safe": False,
        "pairs": [["（", "）"], ["(", ")"], ["《", "》"], ["「", "」"],
                  ["【", "】"], ["[", "]"]],
    },
    "quote_mismatch": {
        "enabled": True, "severity": "P1", "autofix_safe": False,
        "pairs": [["“", "”"], ["「", "」"]],
    },
    "dash_style": {
        "enabled": True, "severity": "P2", "autofix_safe": False,
        "rules": [
            {"name": "two_hyphens_as_dash", "pattern": "--", "fix": "——"},
        ],
    },
    "ellipsis_style": {
        "enabled": True, "severity": "P2", "autofix_safe": True,
        "rules": [
            {"name": "three_dots", "pattern": r"\.{3}", "fix_zh": "……", "fix_en": "..."},
        ],
    },
    "fullwidth_halfwidth_mix": {
        "enabled": True, "severity": "P2", "autofix_safe": False,
    },
    "weak_phrases": {
        "enabled": True, "severity": "P3", "autofix_safe": False,
        "patterns": ["众所周知", "不言而喻", "显而易见", "毋庸置疑",
                     "笔者认为", "值得指出的是", "有必要指出"],
    },
    "zh_en_symbol_spacing": {
        "enabled": True, "severity": "P2", "autofix_safe": True,
    },
}


# ── 规则加载 ──────────────────────────────────────────────

def load_rules(rules_path: Path | None) -> dict[str, Any]:
    """加载规则配置。优先 YAML，回退硬编码默认值。"""
    if rules_path and rules_path.exists() and yaml:
        with open(rules_path, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    return _DEFAULT_RULES


# ── 检查器 ──────────────────────────────────────────────

def _check_regex_patterns(text: str, patterns: list[str], rule_name: str) -> list[dict]:
    """通用正则模式匹配。"""
    issues = []
    for pat in patterns:
        for m in re.finditer(pat, text):
            issues.append({
                "rule": rule_name,
                "match": m.group(),
                "position": m.start(),
                "context": text[max(0, m.start() - 10):m.end() + 10],
            })
    return issues


def _check_bracket_mismatch(text: str, pairs: list[list[str]], rule_name: str) -> list[dict]:
    """检测括号配对。"""
    issues = []
    for left, right in pairs:
        left_count = text.count(left)
        right_count = text.count(right)
        if left_count != right_count:
            issues.append({
                "rule": rule_name,
                "detail": f"'{left}' 出现 {left_count} 次，'{right}' 出现 {right_count} 次",
                "left": left, "right": right,
                "left_count": left_count, "right_count": right_count,
            })
    return issues


def _check_fullwidth_halfwidth_mix(text: str, rule_name: str) -> list[dict]:
    """检测同一段落中全角/半角标点混用。"""
    issues = []
    pairs = [
        ("，", ",", "逗号"),
        ("。", ".", "句号"),
        ("；", ";", "分号"),
        ("：", ":", "冒号"),
    ]
    for fw, hw, label in pairs:
        if fw in text and hw in text:
            # 排除英文句子中的正常半角句号
            if label == "句号":
                # 如果有英文句子，半角句号是正常的
                if re.search(r"[A-Za-z]+\.", text):
                    continue
            issues.append({
                "rule": rule_name,
                "detail": f"全角'{fw}'与半角'{hw}'混用（{label}）",
            })
    # 括号单独检查
    if "（" in text and "(" in text:
        issues.append({
            "rule": rule_name,
            "detail": "全角'（'与半角'('混用（括号）",
        })
    return issues


def _check_mixed_quotes(text: str, rule_name: str) -> list[dict]:
    """检测引号风格混用。"""
    issues = []
    has_curly = "“" in text or "”" in text  # ""
    has_corner = "「" in text or "」" in text
    if has_curly and has_corner:
        issues.append({
            "rule": rule_name,
            "detail": "弯引号""与直角引号「」混用",
        })
    return issues


def _check_en_punct_in_cjk(text: str, rule_name: str) -> list[dict]:
    """检测中文语境中误用英文标点。

    判据是「与中文字符直接邻接」——这样能精准抓到"用英文标点包裹/分隔中文"的错误，
    又不会误伤中文句子里正常嵌入的英文短语（如"杰夫·范德米尔（Jeff VanderMeer）"里
    英文词后的逗号，因为它前面是英文字母而非中文）。
    """
    issues = []

    def add(m, detail):
        issues.append({
            "rule": rule_name,
            "match": m.group(),
            "position": m.start(),
            "context": text[max(0, m.start() - 12):m.end() + 12],
            "detail": detail,
        })

    # 1) 中文字符后紧跟英文逗号/分号/冒号/问号/叹号 → 应用中文标点
    cn = {",": "，", ";": "；", ":": "：", "?": "？", "!": "！"}
    for m in re.finditer(r"([一-鿿])([,;:?!])", text):
        en = m.group(2)
        add(m, f"中文后误用英文'{en}'，应为中文'{cn[en]}'")

    # 2) 中文字符后紧跟英文句点（排除小数、省略号、英文缩写如 e.g.）
    for m in re.finditer(r"([一-鿿])\.(?!\.)(?=\s|$|[一-鿿“”\"'）)】」])", text):
        add(m, "中文后误用英文'.'，应为中文句号'。'")

    # 3) 英文直引号包裹含中文的内容 → 应用中文引号 “”/「」
    for m in re.finditer(r"\"[^\"\n]*[一-鿿][^\"\n]*\"", text):
        add(m, "英文引号\"\"包裹中文，应用中文引号“”或「」")
    for m in re.finditer(r"'[^'\n]*[一-鿿][^'\n]*'", text):
        add(m, "英文单引号''包裹中文，应用中文引号‘’")

    # 4) 英文圆括号直接贴着中文（开括号后即中文，或中文后即闭括号）→ 应用全角（）
    for m in re.finditer(r"\([一-鿿]|[一-鿿]\)", text):
        add(m, "中文内容用了英文圆括号()，应为全角（）")

    return issues


def check_paragraph(text: str, rules: dict[str, Any]) -> list[dict]:
    """对单个段落执行所有启用的规则检查。"""
    if not text or not text.strip():
        return []

    all_issues: list[dict] = []

    for rule_name, rule_cfg in rules.items():
        if not isinstance(rule_cfg, dict) or not rule_cfg.get("enabled", True):
            continue

        severity = rule_cfg.get("severity", "P3")

        # 正则规则族（cjk_latin_spacing, dash_style, ellipsis_style 等）
        if "rules" in rule_cfg:
            for sub_rule in rule_cfg["rules"]:
                pat = sub_rule.get("pattern")
                if pat:
                    for m in re.finditer(pat, text):
                        all_issues.append({
                            "rule": rule_name,
                            "sub_rule": sub_rule.get("name", ""),
                            "severity": severity,
                            "autofix_safe": rule_cfg.get("autofix_safe", False),
                            "match": m.group(),
                            "position": m.start(),
                            "context": text[max(0, m.start() - 15):m.end() + 15],
                            "description": sub_rule.get("description", ""),
                        })

        # 简单模式匹配（repeated_punctuation, weak_phrases, connector_blacklist）
        if "patterns" in rule_cfg:
            issues = _check_regex_patterns(
                text, rule_cfg["patterns"], rule_name
            )
            for iss in issues:
                iss["severity"] = severity
                iss["autofix_safe"] = rule_cfg.get("autofix_safe", False)
            all_issues.extend(issues)

        # 括号配对检查
        if rule_name == "bracket_mismatch" and "pairs" in rule_cfg:
            issues = _check_bracket_mismatch(text, rule_cfg["pairs"], rule_name)
            for iss in issues:
                iss["severity"] = severity
            all_issues.extend(issues)

        # 引号配对检查
        if rule_name == "quote_mismatch" and "pairs" in rule_cfg:
            issues = _check_bracket_mismatch(text, rule_cfg["pairs"], rule_name)
            for iss in issues:
                iss["severity"] = severity
            all_issues.extend(issues)

        # 全角/半角混用
        if rule_name == "fullwidth_halfwidth_mix":
            issues = _check_fullwidth_halfwidth_mix(text, rule_name)
            for iss in issues:
                iss["severity"] = severity
            all_issues.extend(issues)

        # 引号风格混用
        if rule_name == "mixed_quote_style":
            issues = _check_mixed_quotes(text, rule_name)
            for iss in issues:
                iss["severity"] = severity
            all_issues.extend(issues)

        # 中文语境误用英文标点
        if rule_name == "en_punct_in_cjk":
            issues = _check_en_punct_in_cjk(text, rule_name)
            for iss in issues:
                iss["severity"] = severity
                iss["autofix_safe"] = rule_cfg.get("autofix_safe", False)
            all_issues.extend(issues)

    return all_issues


# ── 文档扫描 ──────────────────────────────────────────────

def scan_document(docx_path: str | Path, rules: dict[str, Any]) -> dict[str, Any]:
    """扫描整个文档，返回检查报告。"""
    doc = Document(str(docx_path))
    results: list[dict[str, Any]] = []
    summary: dict[str, int] = {}

    for para_idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # 确定段落位置
        location = "正文"
        style_name = para.style.name if para.style else ""
        if "Heading" in style_name or "标题" in style_name:
            location = f"标题({style_name})"
        elif "TOC" in style_name:
            location = "目录"
        elif "Footnote" in style_name:
            location = "脚注"

        issues = check_paragraph(text, rules)
        if issues:
            for iss in issues:
                iss["paragraph_index"] = para_idx
                iss["location"] = location
                iss["paragraph_text"] = text[:80] + ("..." if len(text) > 80 else "")
                rule = iss["rule"]
                summary[rule] = summary.get(rule, 0) + 1
            results.extend(issues)

    # 按严重程度排序
    severity_order = {"P1": 0, "P2": 1, "P3": 2}
    results.sort(key=lambda x: (severity_order.get(x.get("severity", "P3"), 3), x["rule"]))

    total = len(results)
    by_severity = {}
    for iss in results:
        sev = iss.get("severity", "P3")
        by_severity[sev] = by_severity.get(sev, 0) + 1

    return {
        "file": str(docx_path),
        "total_issues": total,
        "by_severity": by_severity,
        "by_rule": summary,
        "issues": results,
    }


# ── CLI ──────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="中文学术论文语言/符号检查（规则引擎）")
    parser.add_argument("docx", help="待检查的 .docx 文件")
    parser.add_argument("--output", "-o", help="输出 JSON 路径")
    parser.add_argument("--rules", "-r", help="规则 YAML 路径（默认 references/language_rules.yaml）")
    parser.add_argument("--severity", "-s", default="P1,P2,P3",
                        help="只报告指定严重程度（逗号分隔，如 P1,P2）")
    args = parser.parse_args()

    docx_path = Path(args.docx)
    if not docx_path.exists():
        print(f"文件不存在: {docx_path}", file=sys.stderr)
        return 1

    # 加载规则
    rules_path = Path(args.rules) if args.rules else None
    if rules_path is None:
        # 尝试默认位置
        candidates = [
            Path(__file__).parent.parent / "references" / "language_rules.yaml",
            Path("references/language_rules.yaml"),
        ]
        for c in candidates:
            if c.exists():
                rules_path = c
                break

    rules = load_rules(rules_path)

    # 扫描文档
    report = scan_document(docx_path, rules)

    # 按严重程度过滤
    allowed_severities = set(args.severity.split(","))
    report["issues"] = [
        iss for iss in report["issues"]
        if iss.get("severity", "P3") in allowed_severities
    ]
    report["total_issues"] = len(report["issues"])

    # 输出
    if args.output:
        Path(args.output).write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"检查完成: {report['total_issues']} 个问题 → {args.output}")
    else:
        print(json.dumps(report, ensure_ascii=False, indent=2))

    # 打印摘要
    print(f"\n=== 语言检查摘要 ===")
    print(f"总问题数: {report['total_issues']}")
    for sev in ["P1", "P2", "P3"]:
        count = report["by_severity"].get(sev, 0)
        if count:
            print(f"  {sev}: {count}")
    for rule, count in sorted(report["by_rule"].items(), key=lambda x: -x[1]):
        print(f"  {rule}: {count}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
