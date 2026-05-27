#!/usr/bin/env python3
"""Package paper-format for universal distribution.

Creates a clean zip file without Claude Code-specific files.

Usage:
    python package.py [--output paper-check.zip]
"""

import os
import sys
import zipfile
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent

# Files/dirs to EXCLUDE from universal package
EXCLUDE = {
    ".remember",
    "test_output",
    "__pycache__",
    ".DS_Store",
    ".claude",
    "evals",
    "*.pyc",
    ".codex-plugin",
    "plugin.json",
}

# Files/dirs to INCLUDE explicitly (overrides exclude)
INCLUDE_PATTERNS = [
    "*.py",
    "*.md",
    "*.csl",
    "*.txt",
    "*.json",
]


def should_include(path: Path, root: Path) -> bool:
    """Check if a file should be included in the package."""
    rel = path.relative_to(root)
    parts = rel.parts

    # Check excluded dirs
    for part in parts:
        if part in EXCLUDE:
            return False
        if part.startswith("."):
            return False

    # Check file extensions
    if path.is_file():
        if path.suffix in (".pyc", ".pyo"):
            return False
        if path.suffix in (".py", ".md", ".csl", ".txt", ".json"):
            return True
        return False

    return True


def package(output_path: str | None = None):
    """Create a clean zip package."""
    if output_path is None:
        output_path = str(SKILL_DIR.parent / "paper-check.zip")

    output = Path(output_path)
    count = 0

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(SKILL_DIR):
            root_path = Path(root)

            # Filter directories in-place to skip excluded ones
            dirs[:] = [
                d for d in dirs
                if d not in EXCLUDE and not d.startswith(".")
            ]

            for fname in sorted(files):
                fpath = root_path / fname
                if not should_include(fpath, SKILL_DIR):
                    continue

                arcname = f"paper-format/{fpath.relative_to(SKILL_DIR)}"
                zf.write(fpath, arcname)
                count += 1

    size_kb = output.stat().st_size / 1024
    print(f"Packaged {count} files -> {output} ({size_kb:.0f} KB)")
    print(f"\nTo install on another agent:")
    print(f"  unzip {output.name} -d <agent-skills-dir>/")
    print(f"  cd paper-format && pip install -r requirements.txt")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else None
    package(out)
