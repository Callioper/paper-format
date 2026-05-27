#!/usr/bin/env python3
"""fix_paper.py - One-pass complete paper fix with output directory management.

Usage:
    python fix_paper.py "论文.docx" --mode journal [--spec spec.json] [--bib refs.bib]
"""

from __future__ import annotations
import argparse
import json
import os
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8') if hasattr(sys.stdout, 'reconfigure') else None

_THIS = Path(__file__).resolve()
_REPO_ROOT = _THIS.parents[0]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT.parent))

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def _apply_font(run, cjk: str, latin: str, size_pt: float):
    """Set CJK + Latin font and size on a run."""
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cjk)
    rFonts.set(qn('w:ascii'), latin)
    rFonts.set(qn('w:hAnsi'), latin)
    run.font.size = Pt(size_pt)


def create_output_dir(thesis_path: Path) -> Path:
    """Create output directory named after the thesis file."""
    stem = thesis_path.stem
    output_dir = thesis_path.parent / stem
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def create_copy(src: Path, output_dir: Path) -> Path:
    """Create a working copy in the output directory."""
    import shutil
    copy_path = output_dir / f"{src.stem}_copy{src.suffix}"
    if not copy_path.exists():
        shutil.copy2(str(src), str(copy_path))
    return copy_path


def fix_paper(
    thesis_path: str,
    mode: str = "journal",
    spec_path: str | None = None,
    bib_path: str | None = None,
) -> dict:
    """Run complete fix pipeline. Returns dict with output paths."""
    thesis = Path(thesis_path)
    if not thesis.exists():
        return {"error": f"File not found: {thesis}"}

    # Create output directory
    output_dir = create_output_dir(thesis)
    copy_path = create_copy(thesis, output_dir)
    repaired_path = output_dir / f"{thesis.stem}_repaired{thesis.suffix}"

    # Load spec if provided
    spec = None
    if spec_path:
        spec_p = Path(spec_path)
        if spec_p.exists():
            spec = json.loads(spec_p.read_text(encoding='utf-8'))

    # Load document
    doc = Document(str(copy_path))

    records = []

    # =============================================
    # 1. Page margins
    # =============================================
    try:
        section = doc.sections[0]
        old_margins = {}
        for side, attr in [("top", "top_margin"), ("bottom", "bottom_margin"),
                           ("left", "left_margin"), ("right", "right_margin")]:
            val = getattr(section, attr, None)
            if val is not None:
                old_margins[side] = round(val / 914400 * 25.4, 1)

        section.top_margin = Cm(2.8)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

        records.append({
            "category": "page_setup", "item": "margins",
            "location": "全文",
            "before": old_margins,
            "after": {"top": 2.8, "bottom": 2.2, "left": 3.0, "right": 2.0},
        })
    except Exception as e:
        records.append({"category": "page_setup", "item": "margins", "error": str(e)})

    # =============================================
    # 2. Normal style
    # =============================================
    try:
        normal = doc.styles['Normal']
        normal.font.size = Pt(12)
        normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        normal.paragraph_format.line_spacing = 1.5
        rPr = normal.font._element
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), '宋体')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    except Exception:
        pass

    # =============================================
    # 3. Fix all body paragraphs
    # =============================================
    body_fixed = 0
    heading_fixed = 0
    in_refs = False
    ref_fixed = 0

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        sname = para.style.name if para.style else ''
        if sname.startswith('Heading'):
            continue

        # References section
        if re.match(r'^参考文献|^References', text) and len(text) < 20:
            in_refs = True

        if in_refs:
            if re.match(r'^致\s*谢|^附\s*录', text):
                in_refs = False
                continue
            if re.match(r'^\[?\d+\]', text):
                for run in para.runs:
                    _apply_font(run, '宋体', 'Times New Roman', 12)
                pPr = para._p.get_or_add_pPr()
                ind = pPr.find(qn('w:ind'))
                if ind is None:
                    ind = OxmlElement('w:ind')
                    pPr.append(ind)
                ind.set(qn('w:left'), '480')
                ind.set(qn('w:hanging'), '480')
                ref_fixed += 1
            elif in_refs:
                # Reference title
                for run in para.runs:
                    _apply_font(run, '黑体', 'Times New Roman', 16)
                    run.font.bold = True
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

        # Heading detection: centered + large font + short text
        is_centered = para.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
        has_large = any(r.font.size and r.font.size > Pt(13) for r in para.runs)

        if is_centered and has_large and len(text) < 40:
            sz = 16 if len(text) < 20 else 15
            for run in para.runs:
                _apply_font(run, '黑体', 'Times New Roman', sz)
                run.font.bold = True
            heading_fixed += 1
            continue

        # Body: fix font only (not spacing for footnotes)
        for run in para.runs:
            _apply_font(run, '宋体', 'Times New Roman', 12)
            body_fixed += 1
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        para.paragraph_format.line_spacing = 1.5

    records.append({
        "category": "paragraphs", "item": "font_and_format",
        "location": "全文",
        "body_fixed": body_fixed,
        "headings_fixed": heading_fixed,
        "refs_fixed": ref_fixed,
    })

    # =============================================
    # 3b. Set footnote numbering: restart each page, use circled numbers
    # =============================================
    try:
        for section in doc.sections:
            sectPr = section._sectPr
            # Add or find footnotePr
            fnPr = sectPr.find(qn('w:footnotePr'))
            if fnPr is None:
                fnPr = OxmlElement('w:footnotePr')
                sectPr.append(fnPr)
            # Restart each page
            numRestart = fnPr.find(qn('w:numRestart'))
            if numRestart is None:
                numRestart = OxmlElement('w:numRestart')
                fnPr.append(numRestart)
            numRestart.set(qn('w:val'), 'eachPage')
            # Number format: decimal (1,2,3) - Word doesn't support circled via XML
            # but we can set the footnote reference style
            numFmt = fnPr.find(qn('w:numFmt'))
            if numFmt is None:
                numFmt = OxmlElement('w:numFmt')
                fnPr.append(numFmt)
            numFmt.set(qn('w:val'), 'decimal')
        records.append({"category": "footnotes", "item": "numbering_restart", "note": "每页重新编号"})
    except Exception as e:
        records.append({"category": "footnotes", "item": "numbering_restart", "error": str(e)})

    # =============================================
    # 4. Save to temp first, then fix footnotes
    # =============================================
    temp_path = repaired_path.with_suffix('.tmp.docx')
    doc.save(str(temp_path))
    del doc  # Release file handle

    # =============================================
    # 5. Fix footnotes via XML (font only, no spacing changes)
    # =============================================
    fn_fixed = 0

    with zipfile.ZipFile(str(temp_path), 'r') as zin:
        with zipfile.ZipFile(str(repaired_path), 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/footnotes.xml':
                    xml = zin.read('word/footnotes.xml')
                    root = etree.fromstring(xml)

                    for fn in root.findall('w:footnote', NS):
                        fn_type = fn.get(f'{{{W}}}type', '')
                        if fn_type in ('separator', 'continuationSeparator'):
                            continue

                        # Fix font ONLY (not paragraph spacing)
                        for run in fn.findall('.//w:r', NS):
                            rPr = run.find('w:rPr', NS)
                            if rPr is None:
                                rPr = etree.SubElement(run, f'{{{W}}}rPr')
                                run.insert(0, rPr)

                            rFonts = rPr.find('w:rFonts', NS)
                            if rFonts is None:
                                rFonts = etree.SubElement(rPr, f'{{{W}}}rFonts')
                            rFonts.set(f'{{{W}}}eastAsia', '宋体')
                            rFonts.set(f'{{{W}}}ascii', 'Times New Roman')
                            rFonts.set(f'{{{W}}}hAnsi', 'Times New Roman')

                            # 10.5pt (五号) = 21 half-points
                            for tag in ('sz', 'szCs'):
                                el = rPr.find(f'w:{tag}', NS)
                                if el is None:
                                    el = etree.SubElement(rPr, f'{{{W}}}{tag}')
                                el.set(f'{{{W}}}val', '21')

                            fn_fixed += 1

                        # Clean extra spaces in text nodes
                        for t in fn.findall('.//w:t', NS):
                            if t.text:
                                t.text = re.sub(r'\s+', ' ', t.text).strip()

                    zout.writestr(item, etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.unlink(str(temp_path))

    records.append({
        "category": "footnotes", "item": "font_only",
        "location": "全文脚注",
        "runs_fixed": fn_fixed,
        "note": "仅修改字体，未修改段落间距和行距",
    })

    # =============================================
    # 6. Citation repair (if .bib provided)
    # =============================================
    if bib_path:
        try:
            from scripts.citation_repair import process_citations, repair_citations_in_docx
            from scripts.extract_references import detect_references_section
            ref_res = detect_references_section(str(repaired_path))
            if ref_res.found:
                cit_results = process_citations(ref_res.raw_citations, bib_path=bib_path)
                doc2 = Document(str(repaired_path))
                repair_citations_in_docx(str(repaired_path), cit_results, doc=doc2)
                doc2.save(str(repaired_path))
                warns = sum(1 for r in cit_results if r.get("warnings"))
                records.append({
                    "category": "citations", "item": "citation_repair",
                    "count": ref_res.citation_count,
                    "with_warnings": warns,
                })
        except Exception as e:
            records.append({"category": "citations", "item": "citation_repair", "error": str(e)})

    return {
        "output_dir": str(output_dir),
        "copy_path": str(copy_path),
        "repaired_path": str(repaired_path),
        "records": records,
    }


def main():
    parser = argparse.ArgumentParser(description="One-pass paper fix with output directory")
    parser.add_argument("thesis", help="Path to thesis .docx")
    parser.add_argument("--mode", "-m", default="journal", choices=["journal", "thesis"])
    parser.add_argument("--spec", "-s", help="Path to spec JSON")
    parser.add_argument("--bib", help="Path to .bib file")
    args = parser.parse_args()

    result = fix_paper(args.thesis, mode=args.mode, spec_path=args.spec, bib_path=args.bib)

    if result.get("error"):
        print(f"Error: {result['error']}")
        sys.exit(1)

    print(f"Output directory: {result['output_dir']}")
    print(f"Working copy: {result['copy_path']}")
    print(f"Repaired file: {result['repaired_path']}")
    print(f"Repair actions: {len(result['records'])}")
    for r in result['records']:
        print(f"  - {r.get('category')}/{r.get('item')}: {r.get('note', '')}")


if __name__ == "__main__":
    main()
