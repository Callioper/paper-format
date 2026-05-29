#!/usr/bin/env python3
"""fix_paper.py - One-pass paper fix: orchestration only.

This is a thin orchestration layer. The actual format repair lives in
`fix_format.fix_format()` so there is a SINGLE source of truth — every fix and
future improvement there (楷体保留、run-in 摘要、参考文献字号、CNU 引文修复…) applies here
automatically. fix_paper only adds: output-directory management, footnote font
normalization (which fix_format does not touch), optional .bib-enriched citation
repair, and HTML report generation.

Usage:
    python fix_paper.py "论文.docx" --mode journal [--spec spec.json] [--bib refs.bib]
"""

from __future__ import annotations
import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

_THIS = Path(__file__).resolve()
_SKILL_ROOT = _THIS.parents[1]
if str(_SKILL_ROOT) not in sys.path:
    sys.path.insert(0, str(_SKILL_ROOT))

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def create_output_dir(thesis_path: Path) -> Path:
    """Create an output directory named after the thesis file."""
    output_dir = thesis_path.parent / thesis_path.stem
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def create_copy(src: Path, output_dir: Path) -> Path:
    """Create a working copy in the output directory (never touch the original)."""
    copy_path = output_dir / f"{src.stem}_copy{src.suffix}"
    if not copy_path.exists():
        shutil.copy2(str(src), str(copy_path))
    return copy_path


def _normalize_footnote_fonts(docx_path: Path) -> int:
    """Set every footnote run to 宋体 + Times New Roman 10.5pt (font only, no spacing).

    fix_format operates through python-docx on the body and does not reach
    word/footnotes.xml, so footnote font normalization is done here at the XML level.
    Uses .//w:r so nested runs (hyperlinks etc.) are covered; the footnote reference
    mark run is left untouched.
    """
    fixed = 0
    with zipfile.ZipFile(str(docx_path)) as z:
        if "word/footnotes.xml" not in z.namelist():
            return 0
        root = etree.fromstring(z.read("word/footnotes.xml"))

    for fn in root.findall("w:footnote", NS):
        if fn.get(f"{{{W}}}type", "") in ("separator", "continuationSeparator"):
            continue
        for run in fn.findall(".//w:r", NS):
            if run.find(".//w:footnoteRef", NS) is not None:
                continue
            if run.find("w:t", NS) is None:
                continue
            rPr = run.find("w:rPr", NS)
            if rPr is None:
                rPr = etree.SubElement(run, f"{{{W}}}rPr")
                run.insert(0, rPr)
            rFonts = rPr.find("w:rFonts", NS)
            if rFonts is None:
                rFonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
            rFonts.set(f"{{{W}}}eastAsia", "宋体")
            rFonts.set(f"{{{W}}}ascii", "Times New Roman")
            rFonts.set(f"{{{W}}}hAnsi", "Times New Roman")
            for tag in ("sz", "szCs"):
                el = rPr.find(f"w:{tag}", NS)
                if el is None:
                    el = etree.SubElement(rPr, f"{{{W}}}{tag}")
                el.set(f"{{{W}}}val", "21")  # 10.5pt = 21 half-points
            fixed += 1

    if fixed:
        tmp = docx_path.with_suffix(".fn.tmp")
        with zipfile.ZipFile(str(docx_path)) as zin, \
                zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "word/footnotes.xml":
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(item, data)
        shutil.move(str(tmp), str(docx_path))
    return fixed


def fix_paper(
    thesis_path: str,
    mode: str = "journal",
    spec_path: str | None = None,
    bib_path: str | None = None,
) -> dict:
    """Orchestrate a full one-pass fix. Returns dict with output paths + records."""
    from scripts.fix_format import fix_format

    thesis = Path(thesis_path)
    if not thesis.exists():
        return {"error": f"File not found: {thesis}"}

    output_dir = create_output_dir(thesis)
    copy_path = create_copy(thesis, output_dir)
    repaired_path = output_dir / f"{thesis.stem}_repaired{thesis.suffix}"

    # ---- Core format repair: delegate to the single source of truth ----
    # When a .bib is supplied we skip fix_format's conservative (no-enrichment)
    # citation pass and run the .bib-enriched citation_repair below instead, so the
    # references section is only processed once.
    fmt = fix_format(
        str(copy_path),
        spec_path=spec_path,
        output_path=str(repaired_path),
        repair_citations=(bib_path is None),
        mode=mode,
    )
    if fmt.get("error"):
        return {"error": fmt["error"]}
    records = list(fmt.get("repair_records", []))

    # fix_format makes its own internal "<stem>_copy_copy.docx" working copy; sweep it.
    for stray in output_dir.glob(f"{copy_path.stem}_copy*"):
        if stray != repaired_path:
            try:
                stray.unlink()
            except OSError:
                pass

    # ---- Footnote font normalization (fix_format does not touch footnotes.xml) ----
    fn_fixed = _normalize_footnote_fonts(repaired_path)
    if fn_fixed:
        records.append({
            "category": "footnotes", "item": "font_only", "location": "全文脚注",
            "runs_fixed": fn_fixed, "note": "仅统一脚注字体为宋体+TNR 10.5pt，不改段落间距/行距",
        })

    # ---- Optional .bib-enriched citation repair ----
    if bib_path:
        try:
            from scripts.citation_repair import process_citations, repair_citations_in_docx
            from scripts.extract_references import detect_references_section
            from docx import Document
            ref_res = detect_references_section(str(repaired_path))
            if ref_res.found:
                cit_results = process_citations(ref_res.raw_citations, bib_path=bib_path)
                doc2 = Document(str(repaired_path))
                repair_citations_in_docx(str(repaired_path), cit_results, doc=doc2)
                doc2.save(str(repaired_path))
                warns = sum(1 for r in cit_results if r.get("warnings"))
                records.append({
                    "category": "citations", "item": "citation_repair",
                    "count": ref_res.citation_count, "with_warnings": warns,
                    "note": ".bib 元数据丰富化",
                })
        except Exception as e:
            records.append({"category": "citations", "item": "citation_repair", "error": str(e)})

    return {
        "output_dir": str(output_dir),
        "copy_path": str(copy_path),
        "repaired_path": str(repaired_path),
        "records": records,
        "thesis_stem": thesis.stem,
    }


def main():
    parser = argparse.ArgumentParser(description="One-pass paper fix (orchestration over fix_format)")
    parser.add_argument("thesis", help="Path to thesis .docx")
    parser.add_argument("--mode", "-m", default="journal", choices=["journal", "thesis"])
    parser.add_argument("--spec", "-s", help="Path to spec JSON")
    parser.add_argument("--bib", help="Path to .bib file")
    args = parser.parse_args()

    result = fix_paper(args.thesis, mode=args.mode, spec_path=args.spec, bib_path=args.bib)
    if result.get("error"):
        print(f"Error: {result['error']}")
        sys.exit(1)

    output_dir = Path(result["output_dir"])
    repaired = result["repaired_path"]
    stem = result.get("thesis_stem", "paper")

    print(f"Output directory: {result['output_dir']}")
    print(f"Working copy: {result['copy_path']}")
    print(f"Repaired file: {repaired}")
    print(f"Repair actions: {len(result['records'])}")
    for r in result["records"]:
        print(f"  - {r.get('category')}/{r.get('item')}: {r.get('note', '')}")

    # Persist repair records (so the report / re-runs can use them)
    (output_dir / "repair_records.json").write_text(
        json.dumps(result["records"], ensure_ascii=False, indent=2), encoding="utf-8")

    # Run format check on the repaired file
    check_result = {}
    try:
        from scripts.check_format import check_format
        check_result = check_format(repaired, mode=args.mode)
        (output_dir / "check_result.json").write_text(
            json.dumps(check_result, ensure_ascii=False, indent=2), encoding="utf-8")
        total = sum(len(v) for v in check_result.get("issues", {}).values())
        print(f"\nRemaining issues: {total}")
    except Exception as e:
        print(f"\nCheck failed: {e}")

    # Generate HTML report
    try:
        from scripts.generate_html_report import generate_html_report
        html = generate_html_report(
            result["copy_path"], repaired, check_result, [],
            repair_records=result["records"], mode=args.mode,
        )
        html_path = output_dir / f"{stem}_report.html"
        html_path.write_text(html, encoding="utf-8")
        print(f"HTML report: {html_path}")
    except Exception as e:
        print(f"HTML report failed: {e}")


if __name__ == "__main__":
    main()
