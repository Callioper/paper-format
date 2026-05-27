"""Quote verification data models.

Used by verify_quotes.py for direct quotation verification against source PDFs.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class QuoteRecord:
    """A direct quotation extracted from the thesis body."""
    index: int = 0
    quote_text: str = ""
    context_before: str = ""  # 30 chars before the quote
    context_after: str = ""   # 30 chars after the quote
    paragraph_index: int = 0  # which paragraph in the document
    source_hint: str = ""     # author/year/page hint from footnote or text

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "quote_text": self.quote_text,
            "context_before": self.context_before[:50],
            "context_after": self.context_after[:50],
            "paragraph_index": self.paragraph_index,
            "source_hint": self.source_hint,
        }


@dataclass
class QuoteVerificationResult:
    """Verification result for a single quotation."""
    quote_index: int = 0
    quote_text: str = ""
    source_file: str = ""
    matched_page: int = 0
    status: str = ""  # "一致" / "基本一致" / "疑似不一致" / "未定位到出处"
    confidence: float = 0.0
    matched_text: str = ""  # the actual text found in PDF
    screenshot_path: str = ""
    note: str = ""

    def to_dict(self) -> dict:
        return {
            "quote_index": self.quote_index,
            "quote_text": self.quote_text[:120],
            "source_file": self.source_file,
            "matched_page": self.matched_page,
            "status": self.status,
            "confidence": round(self.confidence, 3),
            "matched_text": self.matched_text[:120],
            "screenshot_path": self.screenshot_path,
            "note": self.note,
        }
