#!/usr/bin/env python3
"""Smoke-test a Zotero CSL style against local Zotero items."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path


CSL_NAMESPACE = "http://purl.org/net/xbiblio/csl"
API_BASE = "http://127.0.0.1:23119/api/users/0"
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Smoke-test a CSL style through Zotero local API.")
    parser.add_argument("csl_path", help="Path to the .csl file")
    parser.add_argument(
        "--style-dir",
        default="~/Zotero/styles",
        help="Zotero styles directory. Default: %(default)s",
    )
    parser.add_argument(
        "--item",
        action="append",
        default=[],
        metavar="LABEL=KEY",
        help="Map a test label to a Zotero item key, e.g. zh_journal=HXKEDRBQ",
    )
    parser.add_argument(
        "--restart-cmd",
        help="Optional shell command used to restart Zotero after copying the style.",
    )
    parser.add_argument(
        "--wait-seconds",
        type=float,
        default=2.0,
        help="Seconds to wait after copying or restarting before probing Zotero.",
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON instead of text")
    return parser.parse_args()


def parse_style_id(csl_path: Path) -> str:
    tree = ET.parse(csl_path)
    root = tree.getroot()
    id_node = root.find(f"{{{CSL_NAMESPACE}}}info/{{{CSL_NAMESPACE}}}id")
    if id_node is None or not (id_node.text or "").strip():
        raise ValueError("Style is missing info/id")
    return (id_node.text or "").strip()


def parse_items(raw_items: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for raw in raw_items:
        if "=" not in raw:
            raise ValueError(f"Invalid --item value: {raw}. Expected LABEL=KEY.")
        label, key = raw.split("=", 1)
        mapping[label.strip()] = key.strip()
    return mapping


def request_json(url: str) -> object:
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


def ensure_api_ready(wait_seconds: float) -> None:
    deadline = time.time() + max(wait_seconds, 1.0) + 10.0
    last_error: str | None = None
    while time.time() < deadline:
        try:
            request_json(f"{API_BASE}/items?limit=1")
            return
        except Exception as exc:  # noqa: BLE001
            last_error = str(exc)
            time.sleep(0.5)
    raise RuntimeError(f"Zotero local API is not ready: {last_error}")


def fetch_rendered(style_id: str, item_key: str) -> dict[str, object]:
    params = urllib.parse.urlencode({"include": "data,citation", "style": style_id})
    return request_json(f"{API_BASE}/items/{urllib.parse.quote(item_key)}?{params}")  # type: ignore[return-value]


def summarize_rendered(label: str, item_key: str, payload: dict[str, object]) -> dict[str, object]:
    data = payload.get("data", {})
    if not isinstance(data, dict):
        data = {}
    citation = payload.get("citation")
    citation_text = citation if isinstance(citation, str) else ""
    plain_text = re.sub(r"<[^>]+>", "", citation_text)
    checks = {
        "contains_url": "http://" in plain_text or "https://" in plain_text,
        "contains_cn_date": any(token in plain_text for token in ("年", "月", "日")),
        "contains_cn_pages": "页" in plain_text,
        "contains_en_vol": "vol." in plain_text,
        "contains_en_no": "no." in plain_text,
        "contains_en_pp": "pp." in plain_text,
        "contains_cn_quotes": "《" in plain_text or "》" in plain_text,
    }
    return {
        "label": label,
        "item_key": item_key,
        "item_type": data.get("itemType"),
        "title": data.get("title"),
        "language": data.get("language"),
        "citation": citation_text,
        "plain_text": plain_text,
        "checks": checks,
    }


def install_style(csl_path: Path, style_dir: Path) -> Path:
    style_dir.mkdir(parents=True, exist_ok=True)
    target = style_dir / csl_path.name
    shutil.copy2(csl_path, target)
    return target


def run_restart(command: str) -> None:
    import subprocess

    subprocess.run(command, shell=True, check=True)


def render_text_report(results: list[dict[str, object]]) -> str:
    lines = []
    for row in results:
        lines.append(f"[{row['label']}] {row['item_key']} {row.get('title')}")
        lines.append(f"citation: {row['citation']}")
        checks = row.get("checks", {})
        if isinstance(checks, dict):
            findings = ", ".join(f"{key}={value}" for key, value in checks.items())
            lines.append(f"checks: {findings}")
        lines.append("")
    return "\n".join(lines).rstrip()


def evaluate_result(row: dict[str, object]) -> list[str]:
    label = str(row.get("label", ""))
    checks = row.get("checks", {})
    if not isinstance(checks, dict):
        return ["missing checks"]

    failures: list[str] = []
    is_zh = label.startswith("zh_")
    is_en = label.startswith("en_")

    if is_zh:
        if checks.get("contains_url"):
            failures.append("中文条目不应输出 URL")
        if checks.get("contains_en_vol"):
            failures.append("中文条目不应输出 vol.")
        if checks.get("contains_en_no"):
            failures.append("中文条目不应输出 no.")
        if checks.get("contains_en_pp"):
            failures.append("中文条目不应输出 pp.")

    if is_en:
        if checks.get("contains_cn_date"):
            failures.append("英文条目不应输出中文日期")
        if checks.get("contains_cn_pages"):
            failures.append("英文条目不应输出中文页码标签")
        if checks.get("contains_cn_quotes"):
            failures.append("英文条目不应输出中文书名号")

    return failures


def main() -> int:
    args = parse_args()
    csl_path = Path(args.csl_path).expanduser().resolve()
    if not csl_path.is_file():
        print(f"Missing CSL file: {csl_path}", file=sys.stderr)
        return 1

    try:
        style_id = parse_style_id(csl_path)
        items = parse_items(args.item)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if not items:
        print("No test items provided. Pass one or more --item LABEL=KEY mappings.", file=sys.stderr)
        return 1

    style_dir = Path(args.style_dir).expanduser().resolve()
    install_target = install_style(csl_path, style_dir)

    if args.restart_cmd:
        run_restart(args.restart_cmd)
    else:
        time.sleep(args.wait_seconds)

    try:
        ensure_api_ready(args.wait_seconds)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    results: list[dict[str, object]] = []
    failures: list[str] = []
    semantic_failures: list[str] = []
    for label, item_key in items.items():
        try:
            payload = fetch_rendered(style_id, item_key)
        except urllib.error.HTTPError as exc:
            failures.append(f"{label}={item_key}: HTTP {exc.code}")
            continue
        except Exception as exc:  # noqa: BLE001
            failures.append(f"{label}={item_key}: {exc}")
            continue
        row = summarize_rendered(label, item_key, payload)
        results.append(row)
        row_failures = evaluate_result(row)
        for failure in row_failures:
            semantic_failures.append(f"{label}={item_key}: {failure}")

    output = {
        "style_path": str(csl_path),
        "installed_to": str(install_target),
        "style_id": style_id,
        "results": results,
        "failures": failures,
        "semantic_failures": semantic_failures,
    }

    if args.json:
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        print(f"Installed: {install_target}")
        print(f"Style ID: {style_id}")
        print(render_text_report(results))
        if failures:
            print("Failures:")
            for failure in failures:
                print(f"- {failure}")
        if semantic_failures:
            print("Semantic failures:")
            for failure in semantic_failures:
                print(f"- {failure}")

    return 1 if failures or semantic_failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
