import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pytest
from docx import Document
from scripts.render_preview import find_soffice, render_to_pdf


def test_find_soffice_returns_path_or_none():
    res = find_soffice()
    assert res is None or isinstance(res, str)


def test_render_to_pdf_when_soffice_available(tmp_path):
    soffice = find_soffice()
    if not soffice:
        pytest.skip("soffice 不可用，跳过渲染集成测试")
    doc = Document()
    doc.add_paragraph("渲染测试 render test 123。")
    src = tmp_path / "src.docx"
    doc.save(str(src))
    pdf = render_to_pdf(str(src), str(tmp_path), soffice_path=soffice)
    assert pdf is not None and Path(pdf).exists() and Path(pdf).suffix == ".pdf"
