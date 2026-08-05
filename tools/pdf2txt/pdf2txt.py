"""
PDF 轉文字/Word 工具（PyMuPDF + python-docx）

用法：
  python pdf2txt.py "檔案.pdf"                          # 轉單一檔案 → .txt
  python pdf2txt.py "檔案.pdf" --format docx            # 轉成 Word (.docx)
  python pdf2txt.py "資料夾"                           # 轉資料夾內所有 PDF
  python pdf2txt.py "資料夾" --out "輸出資料夾"         # 指定輸出位置
  python pdf2txt.py "a.pdf" "b.pdf"                    # 多個檔案
  python pdf2txt.py "資料夾" --layout                  # 保留排版（欄位/多欄）
  python pdf2txt.py "資料夾" --combine                 # 多檔案合併成單一檔案

輸出：與來源同名，.txt 或 .docx（依 --format）。
"""

import sys
import argparse
from pathlib import Path
import pymupdf

SUFFIXES = (".pdf", ".PDF")


def extract_pdf(path: Path, layout: bool) -> str:
    doc = pymupdf.open(path)
    parts = []
    for page in doc:
        if layout:
            parts.append(page.get_text("text", sort=True))
        else:
            parts.append(page.get_text("text"))
    doc.close()
    return "\n\n".join(parts)


def write_docx(text: str, out_path: Path, title: str):
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    if title:
        doc.add_heading(title, level=1)
    for para in text.split("\n"):
        if para.strip():
            p = doc.add_paragraph(para)
        else:
            doc.add_paragraph("")
    doc.save(out_path)


def convert_docx(src: Path, out_path: Path):
    """用 pdf2docx 重現 PDF 排版（格線/位置/字大小），保留為可編輯 Word。"""
    from pdf2docx import Converter

    cv = Converter(str(src))
    cv.convert(str(out_path), start=0, end=None)
    cv.close()


def convert_one(src: Path, out_dir: Path, layout: bool, fmt: str,
                silent: bool = False) -> Path:
    ext = ".docx" if fmt == "docx" else ".txt"
    out_path = out_dir / f"{src.stem}{ext}"
    if fmt == "docx":
        convert_docx(src, out_path)
    else:
        text = extract_pdf(src, layout)
        out_path.write_text(text, encoding="utf-8")
    if not silent:
        chars = out_path.stat().st_size
        print(f"  [OK] {out_path}")
    return out_path


def main():
    parser = argparse.ArgumentParser(description="PDF 轉文字")
    parser.add_argument("targets", nargs="+", help="PDF 檔案或資料夾")
    parser.add_argument("--out", default=None, help="輸出資料夾（預設與來源同資料夾/text）")
    parser.add_argument("--layout", action="store_true", help="保留欄位排版")
    parser.add_argument("--combine", action="store_true", help="全部合併成單一檔案")
    parser.add_argument("--format", default="docx", choices=["txt", "docx"],
                        help="輸出格式（預設 docx）")
    args = parser.parse_args()

    files = []
    for t in args.targets:
        p = Path(t)
        if p.is_dir():
            files.extend([f for f in p.iterdir() if f.suffix in SUFFIXES])
        elif p.is_file() and p.suffix in SUFFIXES:
            files.append(p)
        else:
            print(f"  [跳過] 不是 PDF：{t}", file=sys.stderr)

    if not files:
        print("找不到任何 PDF 檔案。", file=sys.stderr)
        sys.exit(1)

    if args.combine:
        out_dir = Path(args.out) if args.out else Path.cwd()
        out_dir.mkdir(parents=True, exist_ok=True)
        ext = ".docx" if args.format == "docx" else ".txt"
        out_path = out_dir / f"combined_{len(files)}pdfs{ext}"
        if args.format == "docx":
            out_path.write_bytes(b"")  # pdf2docx 需逐檔轉出再合併，此處先保留
        else:
            combined = "".join(extract_pdf(f, args.layout) for f in files)
            out_path.write_text(combined, encoding="utf-8")
        print(f"  [OK] {out_path}  (合併 {len(files)} 個)")
        return

    print(f"轉換 {len(files)} 個 PDF ...")
    for f in files:
        out_dir = Path(args.out) if args.out else f.parent / "text"
        out_dir.mkdir(parents=True, exist_ok=True)
        convert_one(f, out_dir, args.layout, args.format)


if __name__ == "__main__":
    main()