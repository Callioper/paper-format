"""paper-check scripts library.

Common utilities for all paper-check scripts.
"""

from __future__ import annotations
import sys


def setup_utf8_output() -> None:
    """Reconfigure stdout/stderr to UTF-8 encoding.

    Fixes Windows GBK encoding errors when printing Unicode characters.
    Safe to call on all platforms — no-op if already UTF-8.
    """
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass
    if hasattr(sys.stderr, 'reconfigure'):
        try:
            sys.stderr.reconfigure(encoding='utf-8')
        except Exception:
            pass
