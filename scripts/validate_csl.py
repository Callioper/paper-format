#!/usr/bin/env python3
"""Validate that a CSL file is well-formed and has the expected root structure."""

from __future__ import annotations

import argparse
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


CSL_NAMESPACE = "http://purl.org/net/xbiblio/csl"
STYLE_TAG = f"{{{CSL_NAMESPACE}}}style"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate a CSL file.")
    parser.add_argument("csl_path", help="Path to the .csl file")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    csl_path = Path(args.csl_path).expanduser().resolve()

    if not csl_path.is_file():
        print(f"Missing CSL file: {csl_path}", file=sys.stderr)
        return 1

    try:
        tree = ET.parse(csl_path)
    except ET.ParseError as exc:
        print(f"XML parse error: {exc}", file=sys.stderr)
        return 1

    root = tree.getroot()
    if root.tag != STYLE_TAG:
        print(
            "Invalid root element: expected "
            f"{{{CSL_NAMESPACE}}}style but found {root.tag}",
            file=sys.stderr,
        )
        return 1

    missing = []
    for tag_name in ("info", "citation", "bibliography"):
        if root.find(f"{{{CSL_NAMESPACE}}}{tag_name}") is None:
            missing.append(tag_name)

    if missing:
        print(
            "Missing required top-level CSL elements: " + ", ".join(missing),
            file=sys.stderr,
        )
        return 1

    print(f"CSL validation passed: {csl_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
