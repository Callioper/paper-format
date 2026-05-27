#!/usr/bin/env python3
"""csl_parser.py - Lightweight CSL (Citation Style Language) XML parser.

Extracts formatting rules from a .csl file for use by citation_repair.py.
Not a full citeproc implementation — extracts key formatting patterns:
  - Author formatting (delimiters, text-case, et-al rules)
  - Title formatting (quotes, italic, brackets)
  - Date formatting (suffixes, delimiters)
  - Page formatting (single/range prefixes, delimiters)
  - Bibliography layout (et-al threshold, entry spacing)

Usage:
    python csl_parser.py style.csl --output csl_rules.json
    python csl_parser.py style.csl  # prints to stdout
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    from lxml import etree
except ImportError:
    import xml.etree.ElementTree as etree


_CSL_NS = {"csl": "http://purl.org/net/xbiblio/csl"}


def _ns(tag: str) -> str:
    """Wrap tag in CSL namespace."""
    return f"{{{_CSL_NS['csl']}}}{tag}"


def _find(element, tag: str):
    """Find child with CSL namespace."""
    return element.find(f"csl:{tag}", _CSL_NS)


def _findall(element, tag: str):
    """Find all children with CSL namespace."""
    return element.findall(f"csl:{tag}", _CSL_NS)


def parse_info(root) -> dict:
    """Extract <info> metadata."""
    info_el = _find(root, "info")
    if info_el is None:
        return {}
    info: dict[str, Any] = {}
    title = _find(info_el, "title")
    if title is not None and title.text:
        info["title"] = title.text.strip()
    summary = _find(info_el, "summary")
    if summary is not None and summary.text:
        info["summary"] = summary.text.strip()
    # citation-format from category
    for cat in _findall(info_el, "category"):
        cf = cat.get("citation-format")
        if cf:
            info["citation_format"] = cf
    return info


def parse_bibliography(root) -> dict:
    """Extract <bibliography> attributes."""
    bib = _find(root, "bibliography")
    if bib is None:
        return {}
    result: dict[str, Any] = {}
    for attr in ("et-al-min", "et-al-use-first", "second-field-align",
                 "entry-spacing", "hanging-indent", "line-spacing"):
        val = bib.get(attr)
        if val is not None:
            key = attr.replace("-", "_")
            result[key] = int(val) if val.isdigit() else val
    return result


def parse_citation(root) -> dict:
    """Extract <citation> layout rules."""
    cit = _find(root, "citation")
    if cit is None:
        return {}
    result: dict[str, Any] = {}
    layout = _find(cit, "layout")
    if layout is not None:
        for attr in ("prefix", "suffix", "delimiter", "vertical-align"):
            val = layout.get(attr)
            if val:
                result[f"layout_{attr.replace('-', '_')}"] = val
    # collapse
    collapse = cit.get("collapse")
    if collapse:
        result["collapse"] = collapse
    return result


def parse_macro_author(root) -> dict:
    """Extract author formatting rules from <macro name="author">."""
    for macro in _findall(root, "macro"):
        if macro.get("name") == "author":
            result: dict[str, Any] = {}
            names = _find(macro, "names")
            if names is None:
                return result
            name = _find(names, "name")
            if name is not None:
                result["delimiter"] = name.get("delimiter", ",")
                result["sort_separator"] = name.get("sort-separator", ", ")
                # text-case on name-part
                for np in _findall(name, "name-part"):
                    part_name = np.get("name", "")
                    tc = np.get("text-case")
                    if tc:
                        result[f"{part_name}_text_case"] = tc
                # and / et-al
                result["and"] = name.get("and", "")
            # et-al
            et_al = _find(macro, "et-al")
            if et_al is None and name is not None:
                # et-al attributes on names element
                for attr in ("et-al-min", "et-al-use-first"):
                    val = names.get(attr)
                    if val:
                        result[attr.replace("-", "_")] = int(val) if val.isdigit() else val
            # substitute
            substitute = _find(names, "substitute")
            if substitute is not None:
                result["has_substitute"] = True
            return result
    return {}


def parse_macro_title(root) -> dict:
    """Extract title formatting rules from <macro name="title">."""
    for macro in _findall(root, "macro"):
        if macro.get("name") == "title":
            result: dict[str, Any] = {}
            # Check for quotes
            text_el = _find(macro, "text")
            if text_el is not None:
                quotes = text_el.get("quotes")
                if quotes == "true":
                    result["quotes"] = True
                font_style = text_el.get("font-style")
                if font_style:
                    result["font_style"] = font_style
            # Check group for type suffix like [M] [J]
            for group in _findall(macro, "group"):
                prefix = group.get("prefix", "")
                suffix = group.get("suffix", "")
                if "[" in prefix or "[" in suffix:
                    result["type_identifier"] = True
                    result["type_prefix"] = prefix
                    result["type_suffix"] = suffix
            return result
    return {}


def parse_macro_publisher(root) -> dict:
    """Extract publisher formatting rules from <macro name="publisher">."""
    for macro in _findall(root, "macro"):
        if macro.get("name") == "publisher":
            result: dict[str, Any] = {}
            for group in _findall(macro, "group"):
                d = group.get("delimiter")
                if d:
                    result["delimiter"] = d
                for child in group:
                    if child.tag == _ns("group"):
                        inner_d = child.get("delimiter")
                        if inner_d:
                            result["inner_delimiter"] = inner_d
            return result
    return {}


def parse_locale(root) -> dict:
    """Extract locale-specific formatting (dates, terms)."""
    result: dict[str, Any] = {}
    for locale in _findall(root, "locale"):
        lang = locale.get("xml:lang", locale.get("{http://www.w3.org/XML/1998/namespace}lang", "default"))
        locale_info: dict[str, Any] = {}

        # Date formatting
        date = _find(locale, "date")
        if date is not None:
            date_info: dict[str, Any] = {}
            date_info["form"] = date.get("form", "")
            for dp in _findall(date, "date-part"):
                part_name = dp.get("name", "")
                part_info: dict[str, Any] = {}
                for attr in ("suffix", "form", "range-delimiter"):
                    val = dp.get(attr)
                    if val:
                        part_info[attr] = val
                if part_info:
                    date_info[part_name] = part_info
            if date_info:
                locale_info["date"] = date_info

        # Terms
        terms = _find(locale, "terms")
        if terms is not None:
            terms_dict: dict[str, str] = {}
            for term in _findall(terms, "term"):
                name = term.get("name", "")
                if name and term.text:
                    terms_dict[name] = term.text.strip()
            if terms_dict:
                locale_info["terms"] = terms_dict

        if locale_info:
            result[lang] = locale_info
    return result


def parse_macros_misc(root) -> dict:
    """Extract page and volume formatting from relevant macros."""
    result: dict[str, Any] = {}

    for macro in _findall(root, "macro"):
        name = macro.get("name", "")
        if name not in ("page", "volume", "year-volume-issue", "issued-year", "accessed-date"):
            continue

        macro_info: dict[str, Any] = {}
        for label in _findall(macro, "label"):
            form = label.get("form", "")
            if form:
                macro_info["label_form"] = form

        for text_el in _findall(macro, "text"):
            for attr in ("prefix", "suffix"):
                val = text_el.get(attr)
                if val:
                    macro_info[f"text_{attr}"] = val

        if macro_info:
            result[name] = macro_info

    return result


def parse_csl(csl_path: str | Path) -> dict:
    """
    Parse a CSL file and extract formatting rules.

    Returns a structured dict suitable for JSON serialization.
    """
    csl_path = Path(csl_path)
    if not csl_path.exists():
        return {"error": f"文件不存在: {csl_path}"}
    if csl_path.suffix.lower() not in (".csl", ".xml"):
        return {"error": f"仅支持 .csl/.xml 格式: {csl_path.suffix}"}

    try:
        parser = etree.XMLParser(remove_blank_text=True) if hasattr(etree, 'XMLParser') else None
        if parser:
            tree = etree.parse(str(csl_path), parser)
            root = tree.getroot()
        else:
            tree = etree.parse(str(csl_path))
            root = tree.getroot()
    except Exception as e:
        return {"error": f"无法解析 CSL 文件: {e}"}

    # Strip namespace for xml.etree compatibility
    if not hasattr(root, 'nsmap'):
        # xml.etree — namespace is in the tag
        pass

    rules: dict[str, Any] = {
        "source": str(csl_path),
        "info": parse_info(root),
        "bibliography": parse_bibliography(root),
        "citation": parse_citation(root),
        "author": parse_macro_author(root),
        "title": parse_macro_title(root),
        "publisher": parse_macro_publisher(root),
        "locale": parse_locale(root),
        "misc_macros": parse_macros_misc(root),
    }

    return rules


def main():
    import sys as _sys
    _this = Path(__file__).resolve()
    _repo_root = _this.parents[1]
    if str(_repo_root) not in _sys.path:
        _sys.path.insert(0, str(_repo_root))
    from scripts.lib import setup_utf8_output
    setup_utf8_output()
    parser = argparse.ArgumentParser(description="Parse CSL file to formatting rules")
    parser.add_argument("csl", help="Path to .csl file")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    rules = parse_csl(args.csl)

    if args.output:
        Path(args.output).write_text(
            json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"Rules written to {args.output}")
    else:
        print(json.dumps(rules, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
