#!/usr/bin/env python3
"""paper-check: Universal entry point for thesis format checking.

Works with any AI agent (Claude Code, Hermes, OpenClaw, Codex, etc.)

Usage:
    python paper_format.py paper   thesis.docx [--mode journal] [--spec spec.json] [--bib refs.bib]   # 一键修复+报告（推荐）
    python paper_format.py check    thesis.docx [--mode journal] [--spec spec.json]
    python paper_format.py fix      thesis.docx [--output repaired.docx]
    python paper_format.py lang     thesis.docx --rules references/language_rules.yaml   # 16 类中英符号检查
    python paper_format.py content  thesis.docx [--mode thesis]                          # 内容结构检查
    python paper_format.py footnotes thesis.docx [--output footnotes.json]
    python paper_format.py report   original.docx repaired.docx report.docx
    python paper_format.py html-report original.docx repaired.docx --output report.html
    python paper_format.py template [--mode thesis] [--output template.docx]
    python paper_format.py verify   thesis.docx --bib refs.bib
    python paper_format.py verify   thesis.docx --sources openalex,crossref
    python paper_format.py refs     thesis.docx --bib refs.bib [--check-urls]
    python paper_format.py cross    thesis.docx
    python paper_format.py quotes   thesis.docx --sources "book.pdf:1-50"
    python paper_format.py spec     sample.docx [--output spec.json]
    python paper_format.py csl      style.csl [--output rules.json]
    python paper_format.py validate-csl style.csl
    python paper_format.py lint-csl style.csl
    python paper_format.py cite     thesis.docx --bib refs.bib [--output report.docx]
"""

from __future__ import annotations

import sys
from pathlib import Path

# Ensure package root is on sys.path
_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))


def main():
    # UTF-8 output for Windows
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
        except Exception:
            pass

    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    command = sys.argv[1]
    # Pass remaining args to the sub-command
    sys.argv = [sys.argv[0]] + sys.argv[2:]

    if command == "paper":
        from scripts.fix_paper import main as paper_main
        paper_main()
    elif command == "check":
        from scripts.check_format import main as check_main
        check_main()
    elif command == "fix":
        from scripts.fix_format import main as fix_main
        fix_main()
    elif command in ("lang", "language"):
        from scripts.check_language import main as lang_main
        lang_main()
    elif command == "content":
        from scripts.check_content import main as content_main
        content_main()
    elif command == "footnotes":
        from scripts.extract_footnotes import main as fn_main
        fn_main()
    elif command in ("html-report", "html"):
        from scripts.generate_html_report import main as html_main
        html_main()
    elif command in ("refs", "refs-enhanced"):
        from scripts.check_references_enhanced import main as refs_main
        refs_main()
    elif command == "validate-csl":
        from scripts.validate_csl import main as vcsl_main
        vcsl_main()
    elif command == "lint-csl":
        from scripts.lint_csl_semantics import main as lcsl_main
        lcsl_main()
    elif command == "report":
        from scripts.generate_report import main as report_main
        report_main()
    elif command == "template":
        from scripts.generate_template import main as template_main
        template_main()
    elif command == "verify":
        # Try local first (has --bib), fall back to external
        if any(arg in sys.argv for arg in ("--bib", "--ris", "--xml")):
            from scripts.verify_local import main as local_main
            local_main()
        else:
            from scripts.verify_external import main as ext_main
            ext_main()
    elif command == "cross":
        from scripts.cross_check import main as cross_main
        cross_main()
    elif command == "quotes":
        from scripts.verify_quotes import main as quotes_main
        quotes_main()
    elif command == "spec":
        from scripts.parse_spec import main as spec_main
        spec_main()
    elif command == "csl":
        from scripts.csl_parser import main as csl_main
        csl_main()
    elif command == "cite":
        from scripts.citation_repair import main as cite_main
        cite_main()
    elif command in ("help", "--help", "-h"):
        print(__doc__)
    else:
        print(f"Unknown command: {command}")
        print("Available: paper, check, fix, lang, content, footnotes, report, html-report, "
              "template, verify, refs, cross, quotes, spec, csl, validate-csl, lint-csl, cite")
        sys.exit(1)


if __name__ == "__main__":
    main()
