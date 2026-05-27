#!/usr/bin/env python3
"""paper-check: Universal entry point for thesis format checking.

Works with any AI agent (Claude Code, Hermes, OpenClaw, Codex, etc.)

Usage:
    python paper_check.py check  thesis.docx [--mode journal] [--spec spec.json]
    python paper_check.py fix    thesis.docx [--output repaired.docx]
    python paper_check.py report original.docx repaired.docx report.docx
    python paper_check.py template [--mode thesis] [--output template.docx]
    python paper_check.py verify thesis.docx --bib refs.bib
    python paper_check.py verify thesis.docx --sources openalex,crossref
    python paper_check.py cross  thesis.docx
    python paper_check.py quotes thesis.docx --sources "book.pdf:1-50"
    python paper_check.py spec   sample.docx [--output spec.json]
    python paper_check.py csl    style.csl [--output rules.json]
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

    if command == "check":
        from scripts.check_format import main as check_main
        check_main()
    elif command == "fix":
        from scripts.fix_format import main as fix_main
        fix_main()
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
        print(f"Available: check, fix, report, template, verify, cross, quotes, spec, csl, cite")
        sys.exit(1)


if __name__ == "__main__":
    main()
