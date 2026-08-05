"""
掃描書 PDF → 圖文並存的 Word（原圖 + Gemini OCR 文字）

用法：
  python scan2docx.py "掃描書.pdf"
  python scan2docx.py "掃描書.pdf" --pages 1-5            # 只處理指定頁
  python scan2docx.py "掃描書.pdf" --out "輸出.docx"       # 指定輸出檔名
  python scan2docx.py "掃描書.pdf" --model gemini-3.1-flash-lite
  python scan2docx.py "掃描書.pdf" --skip-ocr             # 只要原圖、不 OCR

逐頁智能判斷（Gemini 回報頁面類型）：
  - text_only（純文字頁）→ 不插圖，只留 OCR 文字
  - image_only（純圖無字）→ 保留原圖，不貼文字
  - mixed（圖文混合）→ 保留原圖 + OCR 文字

排版：
  - 連續排版：不做逐頁分頁，跨頁未完句子自動接合
  - 標題（章/節/目錄）用較大字體（16pt 黑體），正文 12pt 宋體、首行縮排
  - 孤行頁碼自動移除；段落依空行/縮排/對白引號重整

流程：
  1. 每頁抽成高解析 PNG
  2. Gemini 分析頁面類型 + OCR（快取於 _scan2docx_*/pages/*.ocr.txt）
  3. 組 Word：依類型決定插圖與文字

前置：~/.gemini.env 內有 GEMINI_API_KEY；需連網。
"""

import sys
import argparse
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


def parse_pages(spec: str, total: int):
    """解析 '1-5'、'3'、'1,3,5' 等頁碼規格，回傳 0-based 索引清單。"""
    if not spec:
        return list(range(total))
    idxs = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = int(a), int(b)
            idxs.extend(range(a - 1, b))
        else:
            idxs.append(int(part) - 1)
    return sorted({i for i in idxs if 0 <= i < total})


def extract_pages(pdf_path: Path, out_dir: Path, idxs, dpi: int = 200):
    import pymupdf
    doc = pymupdf.open(pdf_path)
    for i in idxs:
        pix = doc[i].get_pixmap(dpi=dpi)
        png = out_dir / f"page_{i + 1:03d}.png"
        pix.save(png)
    doc.close()


def ocr_page(img_path: Path, model: str):
    """回傳 (kind, text)。kind: text_only | image_only | mixed"""
    import base64
    import json
    from google import genai
    from google.genai import types

    key = None
    env = Path.home() / ".gemini.env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("GEMINI_API_KEY="):
                key = line.split("=", 1)[1].strip().strip('"').strip("'")
    if not key:
        print("錯誤：找不到 GEMINI_API_KEY（~/.gemini.env）", file=sys.stderr)
        sys.exit(1)

    client = genai.Client(api_key=key)
    prompt = (
        "你是專業的影像分析助手。請分析這張圖片並回傳 JSON，格式如下：\n"
        '{"type": "text_only" 或 "image_only" 或 "mixed", "text": "..."}\n'
        "判斷規則：\n"
        "- 若圖片是純文字頁（幾乎只有文字）→ type=text_only\n"
        "- 若圖片是純插圖/照片/無文字 → type=image_only\n"
        "- 若圖片同時有文字與插圖/照片 → type=mixed\n"
        "text 欄位：用 OCR 完整提取圖片中的文字，嚴格保留原文語言（不翻譯），"
        "維持標點、段落與閱讀順序（直排中文依由上至下、由右至左）。\n"
        "文字整理要求：\n"
        "1. 依照原書的段落分塊：不同段落之間用一個空行（\\n\\n）分隔。\n"
        "2. 段落內部不要換行（把斷行接回同一段）。\n"
        "3. 頁碼（頁面邊緣或底部的獨立數字）一律忽略，不要輸出。\n"
        "4. 純數字行不要輸出。\n"
        "image_only 時 text 留空字串。\n"
        "只回傳 JSON，不要加任何其他內容或 markdown 標記。"
    )
    image = types.Part.from_bytes(
        data=img_path.read_bytes(), mime_type="image/png")
    import time

    def _call():
        return client.models.generate_content(
            model=model,
            contents=[prompt, image],
        )

    resp = None
    last_wait = 0
    for attempt in range(8):
        try:
            resp = _call()
            break
        except Exception as e:
            msg = str(e)
            wait = 0
            if "429" in msg and "retryDelay" in msg:
                import re as _re
                m = _re.search(r"retryDelay': '(\d+)s", msg)
                if m:
                    wait = int(m.group(1))
            wait = max(wait, last_wait * 2, 3)
            last_wait = wait
            print(f"    429 限流，等待 {wait} 秒後重試（{attempt + 1}/8）...",
                  file=sys.stderr)
            time.sleep(wait)
    if resp is None:
        raise RuntimeError("Gemini OCR 多次重試仍失敗（限流）")
    raw = resp.text.strip() if resp.text else ""
    try:
        data = json.loads(raw.strip("` \n").removeprefix("json"))
    except Exception:
        data = {"type": "mixed", "text": raw}
    kind = data.get("type") if data.get("type") in (
        "text_only", "image_only", "mixed") else "mixed"
    return kind, (data.get("text") or "").strip()


def reflow_text(text: str) -> list:
    """重整 OCR 文字為段落清單：
    - 移除孤行頁碼（如『8』，全行僅數字）
    - 空行、行首縮排、或行首為對話引號（「『）視為新段落
    - 其餘斷行直接接回同一段
    """
    import re

    paras: list[str] = []
    cur: list[str] = []
    prev_was_toc = False
    for ln in text.splitlines():
        s = ln.strip()
        if not s:
            if cur:
                paras.append("".join(cur))
                cur = []
            continue
        if re.fullmatch(r"\d{1,4}", s):  # 孤行頁碼
            continue
        toc_line = bool(re.match(r"^第[一二三四五六七八九十百千\d]+章\s*\S", s))
        starts_new = (
            s.startswith("　") or s.startswith("  ") or s.startswith("\t")
            or s.startswith("“") or s.startswith("「") or s.startswith("『")
            or (toc_line and prev_was_toc)   # 目錄每章各成一行
        )
        if cur and starts_new:
            paras.append("".join(cur))
            cur = [s]
        else:
            cur.append(s)
        prev_was_toc = toc_line
    if cur:
        paras.append("".join(cur))
    return [p for p in paras if p.strip()]


def is_heading(s: str) -> bool:
    """判斷是否為標題（章/節/目錄等）"""
    import re
    t = s.strip()
    if not t:
        return False
    if re.match(r"^第[一二三四五六七八九十百千\d]+章", t):
        return True
    if re.match(r"^\d+[\.、．]\s*\S", t):
        return True
    norm = re.sub(r"\s", "", t)
    if norm in {"目录", "内容提要", "后记", "作者像", "序", "前言", "题记"}:
        return True
    return False


def is_terminal(s: str) -> bool:
    """是否以句末標點（或閉引號）結束"""
    import re
    return bool(
        re.search(r"[。！？…\u2026」』”）)]\s*$", s)
        or re.search(r"[.!?]\s*$", s)
    )


def build_docx(pages_dir: Path, out_path: Path, idxs, model: str,
               skip_ocr: bool, dpi: int = 200):
    import re
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "SimSun"
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
    style.paragraph_format.first_line_indent = Pt(24)
    style.paragraph_format.space_after = Pt(6)

    items = []          # ("text"|"heading", 文字) 或 ("img", 路徑)
    prev_terminal = True   # 上一頁最後一段是否句末結束
    prev_has_img = False

    for n, i in enumerate(idxs):
        png = pages_dir / f"page_{i + 1:03d}.png"
        if not png.exists():
            continue
        kind = "mixed"
        text = ""
        if not skip_ocr:
            txt = pages_dir / f"page_{i + 1:03d}.ocr.txt"
            if txt.exists():
                try:
                    kind, text = txt.read_text(encoding="utf-8").split("\n", 1)
                except Exception:
                    kind, text = "mixed", ""
            else:
                kind, text = ocr_page(png, model)
                txt.write_text(f"{kind}\n{text}", encoding="utf-8")

        page_has_img = kind != "text_only"
        if page_has_img:
            items.append(("img", str(png)))

        paras = reflow_text(text) if (text and kind != "image_only") else []
        in_page_heading = False
        for j, para in enumerate(paras):
            heading = is_heading(para) or (
                in_page_heading and len(para) <= 15 and not is_terminal(para)
            )
            # 章節標題頁（如「第一章」）→ 分頁 + 置中
            if heading and re.match(r"^第[一二三四五六七八九十百千\d]+章\s*$", para):
                items.append(("page_break", None))
                items.append(("chapter_title", para))
                continue
            # 跨頁接合：上頁未句末結束 → 併入下頁首段
            if (
                j == 0 and not page_has_img and not prev_has_img
                and items and items[-1][0] in ("text", "heading")
                and not prev_terminal and not heading
            ):
                items[-1] = (items[-1][0], items[-1][1] + para)
            else:
                items.append(("heading", para) if heading else ("text", para))
            if heading:
                in_page_heading = True
        prev_terminal = (not paras) or is_terminal(paras[-1])
        prev_has_img = page_has_img

    toc_mode = False   # 目前是否在目錄段落內
    for kind_, payload in items:
        if kind_ == "img":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(payload, width=Inches(4.5))
            continue
        if kind_ == "page_break":
            doc.add_page_break()
            continue
        if kind_ == "chapter_title":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(payload)
            r.bold = True
            r.font.size = Pt(22)
            r.font.name = "SimHei"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(72)
            p.paragraph_format.space_after = Pt(24)
            continue

        norm = re.sub(r"\s", "", payload)
        if kind_ == "heading" and norm == "目录":
            toc_mode = True
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(payload)
            r.bold = True
            r.font.size = Pt(16)
            r.font.name = "SimHei"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_after = Pt(12)
            continue

        p = doc.add_paragraph()
        r = p.add_run(payload)
        in_toc_entry = (
            toc_mode and bool(re.match(r"^第[一二三四五六七八九十百千\d]+章", payload))
        )
        if kind_ == "heading" and not in_toc_entry:
            r.bold = True
            r.font.size = Pt(16)
            r.font.name = "SimHei"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(12)
            toc_mode = False
        else:
            if in_toc_entry:
                # 目錄條目：獨立行、不縮排
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.left_indent = Pt(24)
    doc.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="掃描書 → 圖文 Word")
    parser.add_argument("pdf", help="掃描 PDF 路徑")
    parser.add_argument("--pages", default="", help="頁碼，如 1-5 或 1,3,5")
    parser.add_argument("--out", default=None, help="輸出 .docx 路徑")
    parser.add_argument("--model", default="gemini-3.1-flash-lite",
                        choices=["gemini-3.6-flash", "gemini-3.1-flash-lite",
                                 "gemini-2.5-pro"],
                        help="Gemini 模型")
    parser.add_argument("--skip-ocr", action="store_true", help="只要原圖")
    parser.add_argument("--dpi", type=int, default=200, help="抽圖解析度")
    args = parser.parse_args()

    pdf = Path(args.pdf)
    if not pdf.exists():
        print(f"錯誤：找不到 {pdf}", file=sys.stderr)
        sys.exit(1)

    import pymupdf
    total = len(pymupdf.open(pdf))
    idxs = parse_pages(args.pages, total)

    work = pdf.parent / f"_scan2docx_{pdf.stem}"
    pages_dir = work / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)

    print(f"共 {total} 頁，處理 {len(idxs)} 頁（{args.pages or '全部'}）")
    print("步驟 1/3：抽取頁面圖 ...")
    extract_pages(pdf, pages_dir, idxs, args.dpi)

    if not args.skip_ocr:
        print("步驟 2/3：Gemini OCR 逐頁讀取 ...")
        import time
        for i in idxs:
            png = pages_dir / f"page_{i + 1:03d}.png"
            txt = pages_dir / f"page_{i + 1:03d}.ocr.txt"
            if not txt.exists():
                kind, text = ocr_page(png, args.model)
                txt.write_text(f"{kind}\n{text}", encoding="utf-8")
                time.sleep(4)  # 避免觸發每分鐘上限
            else:
                kind = txt.read_text(encoding="utf-8").split("\n", 1)[0]
            print(f"  頁 {i + 1} ✓（{kind}）")
    else:
        print("步驟 2/3：跳過 OCR（--skip-ocr）")

    out = Path(args.out) if args.out else pdf.parent / f"{pdf.stem}_scan.docx"
    print(f"步驟 3/3：組 Word → {out}")
    build_docx(pages_dir, out, idxs, args.model, args.skip_ocr, args.dpi)
    print(f"完成：{out}")


if __name__ == "__main__":
    main()
