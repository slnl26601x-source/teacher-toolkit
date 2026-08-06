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
        "判斷規則（type 影響是否輸出原圖，務必依『內容主體』判斷）：\n"
        "- 頁面以文字為主（插圖/照片只是小範圍點綴或章首裝飾，正文仍佔大半）"
        "→ type=text_only，此時會輸出全部文字、不保留整頁圖。\n"
        "- 頁面幾乎只有插圖/照片、文字極少或只有圖說 → type=image_only。\n"
        "- 頁面文字與插圖/照片都佔相當比例、彼此並列 → type=mixed。\n"
        "text 欄位：用 OCR 完整提取圖片中的文字，嚴格保留原文語言（不翻譯），"
        "維持標點、段落與閱讀順序（直排中文依由上至下、由右至左）。\n"
        "文字整理要求：\n"
        "1. 依照原書的段落分塊：不同段落之間用一個空行（\\n\\n）分隔。\n"
        "2. 段落內部不要換行（把斷行接回同一段）。\n"
        "3. 頁面頂部/底部的獨立頁碼（純數字的孤行）忽略，不要輸出。\n"
        "4. 若該頁是目錄/目次：條目後方或右側的頁碼數字「要保留」，"
        "直接接在該條目同一行末尾（例如『第一章　認識自然　15』）。\n"
        "5. 純數字行（孤行頁碼）不要輸出。\n"
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
        # JSON 解析失敗：寬鬆抽取 "type" 與 "text"，避免把整包 JSON 當文字存檔
        import re as _re
        tm = _re.search(r'"type"\s*:\s*"(\w+)"', raw)
        kind2 = tm.group(1) if tm and tm.group(1) in (
            "text_only", "image_only", "mixed") else "mixed"
        text2 = ""
        im = _re.search(r'"text"\s*:\s*"((?:[^"\\]|\\.)*)"', raw, _re.S)
        if im:
            try:
                text2 = json.loads('"' + im.group(1) + '"')
            except Exception:
                text2 = im.group(1)
        # 若內容含未跳脫引號導致截斷，改從整段 JSON 抓取直到最後一個結尾引號
        if "}," not in text2 and text2 and not text2.rstrip().endswith(("。", "！", "？", "\"", "」", "』")):
            m2 = _re.search(r'"text"\s*:\s*"(.+)"\s*\}?\s*$', raw, _re.S)
            if m2:
                try:
                    text2 = json.loads('"' + m2.group(1) + '"')
                except Exception:
                    text2 = m2.group(1)
        data = {"type": kind2, "text": text2}
    kind = data.get("type") if data.get("type") in (
        "text_only", "image_only", "mixed") else "mixed"
    return kind, (data.get("text") or "").strip()


_rapid_engine = None


def rapidocr_text(png: Path) -> str:
    """免費本機 OCR（RapidOCR，離線、無 API 額度）。失敗回傳空字串。"""
    global _rapid_engine
    try:
        if _rapid_engine is None:
            from rapidocr_onnxruntime import RapidOCR
            _rapid_engine = RapidOCR()
        result, _ = _rapid_engine(str(png))
    except Exception:
        return ""
    if not result:
        return ""
    lines = []
    for box, text, score in result:
        t = (text or "").strip()
        if t:
            lines.append(t)
    return "\n".join(lines)


_CJK_RE = r"[\u2e80-\u2fff\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uF900-\uFAFF\u3040-\u30ff\uFF01-\uFF60]"


def fullwidth_punct(s: str) -> str:
    """半形標點轉全形：相鄰為中文字（CJK）時轉換，位於英文詞內時保留。

    - ( ) → （ ）   , → ，   . → 。   ? → ？   ! → ！
    - : → ：   ; → ；   「『」等既為全形不變
    - 英文單字內的 - " ' .（如 Nokus-Ele、"Buffalo Bill"、St.、People's）維持半形
    """
    import re
    if not s or not re.search(r"[,()?!:;.']", s):
        return s
    full = {",": "，", "(": "（", ")": "）", "?": "？",
            "!": "！", ":": "：", ";": "；", ".": "。"}
    out = []
    for i, ch in enumerate(s):
        if ch not in full:
            out.append(ch)
            continue
        left = s[i - 1] if i > 0 else ""
        right = s[i + 1] if i + 1 < len(s) else ""
        # 兩側皆非中文（英文/數字/空白/符號）→ 視為英文語境，保留半形
        if not re.match(_CJK_RE, left) and not re.match(_CJK_RE, right):
            out.append(ch)
        else:
            out.append(full[ch])
    return "".join(out)


def reflow_text(text: str, toc: bool = False) -> list:
    """重整 OCR 文字為段落清單：
    - 移除孤行頁碼（如『8』，全行僅數字）
    - 空行、行首縮排、或行首為對話引號（「『）視為新段落
    - 其餘斷行直接接回同一段
    - toc=True（目錄頁）：每一行各自成一段，不整併
    """
    import re

    paras: list[str] = []
    cur: list[str] = []
    prev_was_toc = False
    for ln in text.splitlines():
        s = fullwidth_punct(ln.strip())
        if not s:
            if cur:
                paras.append("".join(cur))
                cur = []
            continue
        if re.fullmatch(r"\d{1,4}", s):  # 孤行頁碼
            continue
        if toc:
            # 目錄頁：每行各自成段（含章/節標題與頁碼行）
            if cur:
                paras.append("".join(cur))
                cur = []
            paras.append(s)
            continue
        toc_line = bool(re.match(r"^第[一二三四五六七八九十百千\d]+章\s*\S", s))
        starts_new = (
            s.startswith("　") or s.startswith("  ") or s.startswith("\t")
            or s.startswith("“") or s.startswith("「") or s.startswith("『")
            or bool(re.match(r"^[*＊●○■□・·]", s))   # 項次/註記標記分行
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


def _merge_symbol_paras(paras: list) -> list:
    """將連續的單一符號段落（如 * * *、- - -）合併成一行，如『＊ ＊ ＊』"""
    import re
    out: list = []
    buf: list = []
    for p in paras:
        t = p.strip()
        if re.fullmatch(r"[*＊●○•·-]{1,3}", t) or re.fullmatch(r"[*＊●○•·-]", t):
            buf.append(t)
        else:
            if buf:
                out.append("  ".join(buf))
                buf = []
            out.append(p)
    if buf:
        out.append("  ".join(buf))
    return out


_CHAPTER_RE = r"^第[一二三四五六七八九十百千參貳\d]+[章篇編部]"
_CHAPTER_ONLY_RE = r"^第[一二三四五六七八九十百千參貳\d]+[章篇編部]\s*$"


def is_heading(s: str) -> bool:
    """判斷是否為標題（章/節/目錄等）"""
    import re
    t = s.strip()
    if not t:
        return False
    if re.match(_CHAPTER_RE, t):
        return True
    if re.match(r"^\d+[\.、．]\s*\S", t):
        return True
    norm = re.sub(r"\s", "", t)
    if norm in {"目录", "目次", "内容提要", "后记", "作者像", "序", "前言",
                "题记", "推薦序", "推荐序", "導言", "导言", "自序", "譯序",
                "譯後記", "跋"}:
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
               skip_ocr: bool, dpi: int = 200, retry_model: str = "gemini-3.6-flash",
               retry_max: int = 5):
    import re
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "PMingLiU"
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")
    style.paragraph_format.first_line_indent = Pt(24)
    style.paragraph_format.space_after = Pt(6)

    items = []          # ("text"|"heading", 文字) 或 ("img", 路徑)
    prev_terminal = True   # 上一頁最後一段是否句末結束
    prev_has_img = False
    toc_region = False   # 是否在目錄區（逢「目次」頁進入，續頁沿用）
    prev_toc_page = False
    retry_used = [0]     # 強模型重試次數計數
    empty_pages = []     # 仍空白待補的頁碼

    for n, i in enumerate(idxs):
        png = pages_dir / f"page_{i + 1:03d}.png"
        if not png.exists():
            continue
        kind = "mixed"
        text = ""
        txt = pages_dir / f"page_{i + 1:03d}.ocr.txt"
        if txt.exists():
            try:
                kind, text = txt.read_text(encoding="utf-8").split("\n", 1)
            except Exception:
                kind, text = "mixed", ""
        elif not skip_ocr:
            kind, text = ocr_page(png, model)
            txt.write_text(f"{kind}\n{text}", encoding="utf-8")
        elif skip_ocr:
            continue   # 跳過 OCR 且無快取 → 略過該頁
        if not skip_ocr:
            # 防呆：應有文字卻回傳空 → 先免費 RapidOCR，再限量強模型重試
            if kind in ("text_only", "mixed") and not (text or "").strip():
                rt = rapidocr_text(png)
                if rt.strip():
                    kind, text = "text_only", rt
                    txt.write_text(f"{kind}\n{text}", encoding="utf-8")
                    print(f"    頁 {i + 1} 文字為空，已用 RapidOCR 補回 {len(rt)} 字元",
                          file=sys.stderr)
                elif retry_used[0] < retry_max:
                    retry_used[0] += 1
                    print(f"    頁 {i + 1} RapidOCR 也空，改用 {retry_model} 重試 "
                          f"（{retry_used[0]}/{retry_max}）...", file=sys.stderr)
                    kind2, text2 = ocr_page(png, retry_model)
                    if text2.strip():
                        kind, text = kind2, text2
                        txt.write_text(f"{kind}\n{text}", encoding="utf-8")
                    else:
                        empty_pages.append(i + 1)
                else:
                    empty_pages.append(i + 1)
                    print(f"    頁 {i + 1} 仍空白（強模型額度已用罄，記錄待補）",
                          file=sys.stderr)

        page_has_img = kind != "text_only"
        if page_has_img:
            items.append(("img", str(png)))

        if "目次" in (text or ""):
            toc_region = True            # 進入目錄區（續頁沿用）
        elif toc_region and text:
            # 目錄續頁：整頁皆為短行 → 仍在目錄；出現長句 → 結束
            longest = max((len(l.strip()) for l in text.splitlines()), default=0)
            if longest > 45:
                toc_region = False
        is_toc_page = toc_region
        paras = (reflow_text(text, toc=is_toc_page)
                 if (text and kind != "image_only") else [])
        if not is_toc_page:
            paras = _merge_symbol_paras(paras)
        if is_toc_page:
            prev_terminal = True   # 目錄頁結束不跨頁接合
        if is_toc_page and not prev_toc_page:
            items.append(("toc_start", None))
        elif not is_toc_page and prev_toc_page:
            items.append(("toc_end", None))
        prev_toc_page = is_toc_page
        in_page_heading = False
        single_short_page = (
            len(paras) == 1 and 0 < len(paras[0].strip()) <= 18
            and not is_terminal(paras[0].strip())
        )
        for j, para in enumerate(paras):
            heading = is_heading(para) or (
                not is_toc_page and (
                    single_short_page
                    or (in_page_heading and len(para) <= 15 and not is_terminal(para))
                )
            )
            # 章節標題頁（如「第一章」「第一篇」）→ 分頁 + 置中
            if heading and re.match(_CHAPTER_ONLY_RE, para):
                items.append(("page_break", None))
                items.append(("chapter_title", para))
                if j + 1 < len(paras):
                    sub = paras[j + 1].strip()
                    if 0 < len(sub) <= 18 and not is_terminal(sub):
                        items[-1] = ("chapter_title", para + "  " + sub)
                        paras[j + 1] = ""   # 副標併入，避免重複輸出
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
        prev_terminal = True if (is_toc_page or single_short_page) else (
            (not paras) or is_terminal(paras[-1]))
        prev_has_img = page_has_img

    toc_mode = False   # 目前是否在目錄段落內
    for kind_, payload in items:
        if kind_ == "toc_start":
            toc_mode = True
            continue
        if kind_ == "toc_end":
            toc_mode = False
            continue
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
        if kind_ == "heading" and norm in ("目录", "目次"):
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
        toc_main = (
            re.match(_CHAPTER_RE, payload)
            or bool(re.match(r"^(推薦序|推荐序|導言|导言|自序|前言|後記|后记|題記|跋)", payload))
        )
        toc_num = bool(re.match(r"^\d+\s*\S", payload)) or bool(
            re.match(r"^[一二三四五六七八九十]+[、.．]", payload))
        in_toc_entry = toc_mode and (toc_main or toc_num or bool(
            re.match(_CHAPTER_RE, payload)))
        if kind_ == "heading" and not in_toc_entry:
            r.bold = True
            r.font.size = Pt(16)
            r.font.name = "SimHei"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(12)
            if toc_mode:
                # 目錄內的標題（如 推薦序/導言）→ 當主層級條目
                r.bold = True
                r.font.size = Pt(14)
                r.font.name = "SimHei"
                r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.left_indent = Pt(0)
        else:
            if toc_mode or in_toc_entry:
                # 目錄條目：獨立行，依層級分字型大小
                p.paragraph_format.first_line_indent = Pt(0)
                if toc_main:
                    r.bold = True
                    r.font.size = Pt(14)
                    r.font.name = "SimHei"
                    r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
                    p.paragraph_format.left_indent = Pt(0)
                    p.paragraph_format.space_before = Pt(8)
                elif toc_num:
                    r.bold = True
                    r.font.size = Pt(12.5)
                    r.font.name = "PMingLiU"
                    r._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")
                    p.paragraph_format.left_indent = Pt(24)
                else:
                    r.font.size = Pt(12)
                    r.font.name = "PMingLiU"
                    r._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")
                    p.paragraph_format.left_indent = Pt(48)
    doc.save(out_path)
    if empty_pages:
        print(f"\n⚠ 以下頁面仍無文字（RapidOCR 與強模型均失敗），請手動補：{empty_pages}",
              file=sys.stderr)


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
    parser.add_argument("--retry-model", default="gemini-3.6-flash",
                        choices=["gemini-3.6-flash", "gemini-3.1-flash-lite",
                                 "gemini-2.5-pro"],
                        help="空文字時的強模型（預設 gemini-3.6-flash）")
    parser.add_argument("--retry-max", type=int, default=5,
                        help="強模型每日/每次上限次數（預設 5）")
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
    build_docx(pages_dir, out, idxs, args.model, args.skip_ocr, args.dpi,
               args.retry_model, args.retry_max)
    print(f"完成：{out}")


if __name__ == "__main__":
    main()
