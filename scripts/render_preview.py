#!/usr/bin/env python3
"""render_preview.py - 把 docx 渲染成 PDF/页图，用于真正的视觉复核。

paper-format 旧 Step 7 声称"用 python-docx 渲染成图片"，但 python-docx 无渲染能力。
本脚本用已安装的 LibreOffice(soffice) 把 docx 转 PDF，再尽力转成 PNG 页图
（优先 pdftoppm，其次 pdf2image，都没有则返回 PDF 并提示）。soffice 缺失时优雅降级。
"""
from __future__ import annotations
import argparse
import shutil
import subprocess
import sys
from pathlib import Path

_WIN_SOFFICE = [
    r"C:/Program Files/LibreOffice/program/soffice.exe",
    r"C:/Program Files (x86)/LibreOffice/program/soffice.exe",
]


def find_soffice(explicit: str | None = None) -> str | None:
    if explicit and Path(explicit).exists():
        return explicit
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for p in _WIN_SOFFICE:
        if Path(p).exists():
            return p
    return None


def render_to_pdf(docx_path: str, out_dir: str, soffice_path: str | None = None) -> str | None:
    soffice = find_soffice(soffice_path)
    if not soffice:
        return None
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True, timeout=120,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    pdf = Path(out_dir) / (Path(docx_path).stem + ".pdf")
    return str(pdf) if pdf.exists() else None


def pdf_to_pngs(pdf_path: str, out_dir: str, dpi: int = 120) -> list[str]:
    """PDF→PNG 页图。优先 pdftoppm(poppler)，否则 pdf2image，否则返回空列表。"""
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    stem = Path(pdf_path).stem
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf_path,
                        str(Path(out_dir) / stem)], check=True, timeout=120,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return sorted(str(p) for p in Path(out_dir).glob(f"{stem}*.png"))
    # pdf2image 缺失（ImportError）与真实转换失败要区分：前者表示"没有转换器"，
    # 应安静返回 []；后者（PDF 损坏等）应让异常向上传播，由 render_preview 报告，
    # 不要把两者都吞成空列表而让调用方误以为只是缺工具。
    try:
        from pdf2image import convert_from_path
    except ImportError:
        return []
    imgs = convert_from_path(pdf_path, dpi=dpi)
    out = []
    for i, im in enumerate(imgs, 1):
        fp = Path(out_dir) / f"{stem}-{i}.png"
        im.save(str(fp))
        out.append(str(fp))
    return out


def render_preview(docx_path: str, out_dir: str, soffice_path: str | None = None) -> dict:
    # soffice 崩溃/超时也应落到 ok:False（而非抛裸 traceback），兑现优雅降级契约。
    try:
        pdf = render_to_pdf(docx_path, out_dir, soffice_path)
    except (subprocess.SubprocessError, OSError) as e:
        return {"ok": False, "reason": f"LibreOffice 渲染失败：{e}", "pdf": None, "pngs": []}
    if not pdf:
        return {"ok": False, "reason": "LibreOffice(soffice) 不可用，无法渲染；请安装或在 Word 中人工复核",
                "pdf": None, "pngs": []}
    # PNG 转换是尽力而为：失败不影响"已出 PDF"，但要把原因写进 note，不静默吞掉。
    try:
        pngs = pdf_to_pngs(pdf, out_dir)
        note = "" if pngs else "已生成 PDF；未找到 pdftoppm/pdf2image，未能转 PNG，可直接查看 PDF"
    except Exception as e:
        pngs = []
        note = f"已生成 PDF；PNG 转换失败（{e}），可直接查看 PDF"
    return {"ok": True, "pdf": pdf, "pngs": pngs, "note": note}


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="把 docx 渲染成 PDF/PNG 用于视觉复核")
    ap.add_argument("docx")
    ap.add_argument("--out-dir", "-o", default=".")
    ap.add_argument("--soffice", help="soffice 可执行文件路径（默认自动探测）")
    args = ap.parse_args()
    res = render_preview(args.docx, args.out_dir, args.soffice)
    if res["ok"]:
        print(f"PDF: {res['pdf']}")
        print(f"PNG 页图: {len(res['pngs'])} 张" + (f"（{res['note']}）" if res['note'] else ""))
    else:
        print(f"渲染失败: {res['reason']}")
        sys.exit(1)


if __name__ == "__main__":
    main()
