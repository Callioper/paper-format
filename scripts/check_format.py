#!/usr/bin/env python3
"""check_format.py - Comprehensive thesis format checking.

Orchestrates format checks via scripts.checks.format_checks module.
Supports two document modes and spec-driven checking.

Modes:
  - thesis:  cover, abstract, TOC, body, references, acknowledgments, appendices
  - journal: title page, abstract, keywords, body, references, footnotes
"""

from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
from typing import Any

_THIS = Path(__file__).resolve()
_REPO_ROOT = _THIS.parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from scripts.extract_references import detect_references_section
from scripts.citation_repair import process_citations, process_footnote_citations
from scripts.extract_footnotes import extract_footnotes
from scripts.checks.format_checks import (
    _pt,
    _check_styles,
    _check_body_paragraphs,
    _check_cover,
    _check_abstract,
    _check_toc,
    _check_tables,
    _check_headers_footers,
    _check_acknowledgments,
    _check_appendices,
)


# -------------------------------------------------------------------
# Document modes
# -------------------------------------------------------------------

DOC_MODES = {
    "thesis": [
        "cover", "abstract", "toc", "body", "references", "acknowledgments", "appendices"
    ],
    "journal": [
        "title_page", "abstract", "keywords", "body", "references", "footnotes"
    ],
}

# -------------------------------------------------------------------
# Default format checklist (GB/T 7713.1-2006 + CNU common rules)
# -------------------------------------------------------------------

UNIVERSAL_CHECKS = {
    "page_setup": {
        "paper_size": "A4 (210mm x 297mm)",
        "margins": {"top": 28, "bottom": 22, "left": 30, "right": 20},
        "binding_edge": 0,
    },
    "fonts": {
        "chinese": "SimSun / Song typeface",
        "english": "Times New Roman",
        "body_size": 12,
    },
    "line_spacing": 1.5,
    "first_line_indent": 2,  # characters
    "heading_sizes": {1: 16, 2: 15, 3: 14, 4: 12},
}


def _load_spec(spec_path: str | Path | None) -> dict | None:
    """Load a spec JSON file produced by parse_spec.py."""
    if not spec_path:
        return None
    p = Path(spec_path)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def _get_spec_value(spec: dict | None, *keys, default=None):
    """Safely traverse nested spec dict."""
    if not spec:
        return default
    current = spec
    for k in keys:
        if isinstance(current, dict) and k in current:
            current = current[k]
        else:
            return default
    return current



# -------------------------------------------------------------------
# Main check function
# -------------------------------------------------------------------

def check_format(
    thesis_path: str,
    spec_path: str | None = None,
    bib_path: str | None = None,
    mode: str = "journal",
    spec: dict | None = None,
) -> dict:
    """
    Run comprehensive format checks on a thesis .docx.

    Args:
        thesis_path: Path to thesis .docx
        spec_path: Path to spec JSON from parse_spec.py (optional)
        bib_path: Path to Better BibTeX .bib file for citation enrichment
        mode: Document mode - "thesis" or "journal"
        spec: Pre-loaded spec dict (alternative to spec_path)

    Returns a dict with keys:
      - page_setup, styles, paragraphs, tables, citations, footnotes, issues
    """
    from docx import Document
    from scripts.lib.doc_helpers import create_working_copy

    thesis_path = str(thesis_path)
    p = Path(thesis_path)
    if not p.exists():
        return {"error": f"File not found: {thesis_path}", "thesis": thesis_path}
    if p.suffix.lower() != ".docx":
        return {"error": f"Only .docx supported: {p.suffix}", "thesis": thesis_path}

    try:
        doc = Document(thesis_path)
    except Exception as e:
        return {"error": f"Cannot open document: {e}", "thesis": thesis_path}

    # Load spec if provided
    if spec is None and spec_path:
        spec = _load_spec(spec_path)

    # Get heading sizes from spec or defaults
    heading_sizes = dict(UNIVERSAL_CHECKS["heading_sizes"])
    if spec:
        for level in range(1, 5):
            h_size = _get_spec_value(spec, "styles", f"Heading {level}", "size")
            if h_size:
                heading_sizes[level] = round(h_size)

    issues: dict = {
        "page_setup": [],
        "styles": [],
        "paragraphs": [],
        "tables": [],
        "citations": [],
        "footnotes": [],
    }

    # Mode-specific issue categories
    sections = DOC_MODES.get(mode, DOC_MODES["thesis"])
    if "cover" in sections:
        issues["cover"] = []
    if "abstract" in sections or "title_page" in sections:
        issues["abstract"] = []
    if "toc" in sections:
        issues["toc"] = []
    if "acknowledgments" in sections:
        issues["acknowledgments"] = []
    if "appendices" in sections:
        issues["appendices"] = []
    issues["headers_footers"] = []

    # --- Core checks ---
    try:
        _check_styles(doc, issues, heading_sizes)
    except Exception as e:
        issues["styles"].append({"item": "check_error", "severity": "low",
                                 "suggestion": f"Style check error: {e}"})

    try:
        _check_body_paragraphs(doc, issues)
    except Exception as e:
        issues["paragraphs"].append({"item": "check_error", "severity": "low",
                                     "suggestion": f"Paragraph check error: {e}"})

    # --- Page setup ---
    margins: dict = {}
    try:
        section = doc.sections[0]
        pw = section.page_width
        ph = section.page_height
        if pw is None or ph is None:
            margins = {}
            issues["page_setup"].append({"item": "page_dimensions_missing",
                                         "severity": "medium",
                                         "suggestion": "Document has no page dimensions, set to A4"})
        else:
            page_width_mm = pw / 914400 * 25.4
            if abs(page_width_mm - 210) > 1:
                issues["page_setup"].append({
                    "item": "paper_size",
                    "expected": "A4 (210mm)",
                    "actual": f"{page_width_mm:.0f}mm",
                    "severity": "high",
                    "suggestion": "Set paper to A4",
                })
            try:
                tm = section.top_margin
                bm = section.bottom_margin
                lm = section.left_margin
                rm = section.right_margin
                if all(v is not None for v in [tm, bm, lm, rm]):
                    margins = {
                        "top": tm / 914400 * 25.4,
                        "bottom": bm / 914400 * 25.4,
                        "left": lm / 914400 * 25.4,
                        "right": rm / 914400 * 25.4,
                    }
                    for side, expected_mm in UNIVERSAL_CHECKS["page_setup"]["margins"].items():
                        actual_mm = margins[side]
                        if abs(actual_mm - expected_mm) > 2:
                            issues["page_setup"].append({
                                "item": f"margin_{side}",
                                "expected": f"{expected_mm}mm",
                                "actual": f"{actual_mm:.1f}mm",
                                "severity": "high",
                                "suggestion": f"Set {side} margin to {expected_mm}mm",
                            })
                else:
                    issues["page_setup"].append({"item": "margins_missing",
                                                  "severity": "medium",
                                                  "suggestion": "No margins set"})
            except Exception:
                pass
    except (IndexError, AttributeError) as e:
        issues["page_setup"].append({"item": "section_error", "severity": "high",
                                     "suggestion": f"Cannot read page setup: {e}"})
        margins = {}

    # --- References + citation check ---
    citation_count = 0
    try:
        ref_result = detect_references_section(thesis_path)
        if ref_result.found:
            citation_count = ref_result.citation_count
            print(f"[citation_repair] Found {citation_count} citations")
            citation_results = process_citations(
                ref_result.raw_citations, bib_path=bib_path
            )
            for r in citation_results:
                warns = r.get("warnings", [])
                if warns:
                    issues["citations"].append({
                        "item": "citation_format",
                        "original": r.get("original", "")[:80],
                        "formatted": r.get("formatted", ""),
                        "warnings": warns,
                        "bib_status": r.get("bib_status", r.get("zotero_status", "")),
                        "severity": max(
                            (w.split("[")[1].split("]")[0] for w in warns if "[" in w),
                            default="low",
                        ),
                    })
        else:
            print("[citation_repair] No references section detected")
    except Exception as e:
        print(f"[citation_repair] Error: {e}")
        issues["citations"].append({"item": "detection_error", "severity": "low",
                                    "suggestion": f"Citation detection error: {e}"})

    # --- Footnote check ---
    try:
        footnote_items = extract_footnotes(thesis_path)
        if footnote_items:
            print(f"[footnote_check] Found {len(footnote_items)} footnotes")
            fn_results = process_footnote_citations(
                footnote_items, bib_path=bib_path
            )
            for r in fn_results:
                warns = r.get("warnings", [])
                if warns:
                    issues["footnotes"].append({
                        "item": "footnote_format",
                        "footnote_id": r.get("footnote_id"),
                        "original": r.get("raw_footnote_text", r.get("original", ""))[:80],
                        "formatted": r.get("formatted", ""),
                        "warnings": warns,
                        "severity": max(
                            (w.split("[")[1].split("]")[0] for w in warns if "[" in w),
                            default="low",
                        ),
                    })
        else:
            print("[footnote_check] No footnotes found")
    except Exception as e:
        print(f"[footnote_check] Error: {e}")
        issues["footnotes"].append({"item": "detection_error", "severity": "low",
                                    "suggestion": f"Footnote detection error: {e}"})

    # --- Mode-specific checks ---
    if "cover" in sections:
        try:
            _check_cover(doc, issues, spec)
        except Exception as e:
            issues["cover"].append({"item": "check_error", "severity": "low",
                                    "suggestion": f"Cover check error: {e}"})

    if "abstract" in sections or "title_page" in sections:
        try:
            _check_abstract(doc, issues, spec)
        except Exception as e:
            issues["abstract"].append({"item": "check_error", "severity": "low",
                                       "suggestion": f"Abstract check error: {e}"})

    if "toc" in sections:
        try:
            _check_toc(doc, issues, spec)
        except Exception as e:
            issues["toc"].append({"item": "check_error", "severity": "low",
                                  "suggestion": f"TOC check error: {e}"})

    try:
        _check_tables(doc, issues, spec)
    except Exception as e:
        issues["tables"].append({"item": "check_error", "severity": "low",
                                 "suggestion": f"Table check error: {e}"})

    try:
        _check_headers_footers(doc, issues, spec)
    except Exception as e:
        issues["headers_footers"].append({"item": "check_error", "severity": "low",
                                          "suggestion": f"Header/footer check error: {e}"})

    if "acknowledgments" in sections:
        try:
            _check_acknowledgments(doc, issues, spec)
        except Exception as e:
            issues["acknowledgments"].append({"item": "check_error", "severity": "low",
                                              "suggestion": f"Acknowledgments check error: {e}"})

    if "appendices" in sections:
        try:
            _check_appendices(doc, issues, spec)
        except Exception as e:
            issues["appendices"].append({"item": "check_error", "severity": "low",
                                         "suggestion": f"Appendices check error: {e}"})

    return {
        "thesis": str(thesis_path),
        "spec": str(spec_path) if spec_path else "default (GB/T 7713.1-2006)",
        "mode": mode,
        "page_setup": margins,
        "citation_count": citation_count,
        "issues": issues,
    }


# -------------------------------------------------------------------
# CLI entry point
# -------------------------------------------------------------------

def main():
    from scripts.lib import setup_utf8_output
    setup_utf8_output()
    parser = argparse.ArgumentParser(
        description="Comprehensive thesis format checker (with citation repair)"
    )
    parser.add_argument("thesis", help="Path to thesis .docx")
    parser.add_argument(
        "--spec", "-s",
        help="Path to spec JSON file (from parse_spec.py)",
    )
    parser.add_argument("--bib", help="Path to Better BibTeX .bib file")
    parser.add_argument("--output", "-o", help="Path for JSON output")
    parser.add_argument(
        "--mode", "-m", choices=["thesis", "journal"], default="journal",
        help="Document mode: thesis (default) or journal",
    )
    args = parser.parse_args()

    result = check_format(
        args.thesis, args.spec,
        bib_path=args.bib,
        mode=args.mode,
    )

    if args.output:
        Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Results written to {args.output}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
