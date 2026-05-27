"""doc_helpers.py - Document handling utilities for paper-check.

Provides:
  - create_working_copy(): always work on a copy to preserve the user's original file
"""

from __future__ import annotations

import shutil
from pathlib import Path


def create_working_copy(src_path: str | Path) -> Path:
    """Create a working copy of the input .docx file.

    Always operates on a copy to preserve the user's original file.
    Returns the path to the copy.

    Copy naming: thesis.docx -> thesis_copy.docx
    """
    src = Path(src_path)
    if not src.exists():
        raise FileNotFoundError(f"文件不存在: {src}")

    stem = src.stem
    suffix = src.suffix
    parent = src.parent

    copy_path = parent / f"{stem}_copy{suffix}"
    # Avoid overwriting existing copies
    counter = 1
    while copy_path.exists():
        copy_path = parent / f"{stem}_copy{counter}{suffix}"
        counter += 1

    shutil.copy2(str(src), str(copy_path))
    return copy_path
