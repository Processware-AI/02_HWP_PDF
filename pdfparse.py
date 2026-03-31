"""
PDF → Markdown 변환 엔진 (공용 모듈)

pdfplumber를 사용하여 PDF를 구조적 Markdown으로 변환합니다.
- 폰트 크기 · 볼드 여부로 제목 수준 자동 추론
- 표 자동 감지 → Markdown 테이블
- 글머리 기호, 페이지 번호, 목차 등 아티팩트 후처리
"""

import re

# ---------------------------------------------------------------------------
# PDF → 원시 Markdown 추출
# ---------------------------------------------------------------------------

def pdf_to_markdown(pdf_path: str) -> str:
    """pdfplumber로 PDF를 파싱하여 깨끗한 Markdown 문자열을 반환한다."""
    import pdfplumber

    # ── 1차 패스: 본문 폰트 크기(최빈값) 계산 ──
    size_counts: dict[float, int] = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for ch in (page.chars or []):
                sz = round(ch.get("size", 0), 1)
                if sz > 0:
                    size_counts[sz] = size_counts.get(sz, 0) + 1
    body_size = max(size_counts, key=size_counts.get) if size_counts else 10.0

    # ── 2차 패스: 페이지별 텍스트 + 테이블 추출 ──
    raw_parts: list[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.find_tables() or []
            table_bboxes = [t.bbox for t in tables]

            def _in_table(char: dict) -> bool:
                x, y = char["x0"], char["top"]
                for x0, top, x1, bottom in table_bboxes:
                    if x0 <= x <= x1 and top <= y <= bottom:
                        return True
                return False

            # ── 표 외부 글자를 줄 단위로 묶기 ──
            chars = [c for c in (page.chars or []) if not _in_table(c)]
            lines: list[tuple[float, bool, str]] = []

            if chars:
                chars_sorted = sorted(
                    chars, key=lambda c: (round(c["top"], 1), c["x0"])
                )
                cur_top = round(chars_sorted[0]["top"], 1)
                cur_chars: list[dict] = [chars_sorted[0]]

                def _flush(cc: list[dict]):
                    text = "".join(c["text"] for c in cc).strip()
                    if not text:
                        return
                    max_sz = max(c.get("size", 0) for c in cc)
                    bold_cnt = sum(
                        1 for c in cc
                        if "bold" in c.get("fontname", "").lower()
                    )
                    is_bold = bold_cnt > len(cc) / 2
                    lines.append((round(max_sz, 1), is_bold, text))

                for ch in chars_sorted[1:]:
                    ch_top = round(ch["top"], 1)
                    if abs(ch_top - cur_top) <= 1.5:
                        cur_chars.append(ch)
                    else:
                        _flush(cur_chars)
                        cur_top = ch_top
                        cur_chars = [ch]
                _flush(cur_chars)

            # ── 줄 → Markdown (폰트 크기 → 제목 수준) ──
            for font_size, is_bold, text in lines:
                if font_size > body_size + 6:
                    raw_parts.append(f"# {text}")
                elif font_size > body_size + 3:
                    raw_parts.append(f"## {text}")
                elif font_size > body_size + 1:
                    raw_parts.append(f"### {text}")
                elif is_bold and font_size >= body_size:
                    raw_parts.append(f"**{text}**")
                else:
                    raw_parts.append(text)

            # ── 표 → Markdown 테이블 ──
            for table in tables:
                rows = table.extract()
                if not rows:
                    continue
                raw_parts.append("")
                for row_idx, row in enumerate(rows):
                    cells = [
                        (cell or "").replace("\n", " ").strip()
                        for cell in row
                    ]
                    raw_parts.append("| " + " | ".join(cells) + " |")
                    if row_idx == 0:
                        raw_parts.append("|" + " --- |" * len(cells))
                raw_parts.append("")

    md_text = "\n\n".join(raw_parts)
    return _postprocess(md_text)


# ---------------------------------------------------------------------------
# 후처리: 아티팩트 정리 · 줄 이어붙이기 · 구조화
# ---------------------------------------------------------------------------

_RE_PAGE_NUM = re.compile(r"^- \d+ -$")
_RE_TOC_DOTS = re.compile(r"[·]{3,}\d*$")
_RE_TOC_HEADING = re.compile(r"^#{1,6}\s*<.*차례>")
_RE_HEADING_SUFFIX = re.compile(
    r"^(#{1,6} .+?)((?:I{1,3}V?|VI{0,3}|[IVX]+)\.\s*)$"
)
_RE_HEADING_NUM_SPACE = re.compile(
    r"^(#{1,6}\s+(?:(?:I{1,3}V?|VI{0,3}|[IVX]+)\.\s*)?)(\d+)\."
)
_RE_BULLET_CHARS = re.compile(r"^[Ÿl]\s*")
_RE_TRAILING_BULLET = re.compile(r"\s*[Ÿl]$")
_RE_CELL_BULLET = re.compile(r"\s*Ÿ\s*")


def _postprocess(md: str) -> str:
    """PDF 추출 결과에서 아티팩트를 정리하고 깨끗한 Markdown으로 다듬는다."""

    raw_lines = [line.strip() for line in md.split("\n")]

    # ── 목차 블록 제거 ──
    cleaned: list[str] = []
    in_toc = False
    for s in raw_lines:
        if _RE_TOC_HEADING.match(s):
            in_toc = True
            continue
        if in_toc:
            if s.startswith("#") and not _RE_TOC_HEADING.match(s):
                in_toc = False
                cleaned.append(s)
            continue
        cleaned.append(s)

    # ── Pass 1: 아티팩트 제거 · 줄 분류 ──
    items: list[tuple[str, str]] = []  # (kind, text)

    for s in cleaned:
        if not s:
            continue
        if _RE_PAGE_NUM.match(s):
            continue
        if _RE_TOC_DOTS.search(s):
            continue

        # 테이블
        if s.startswith("|"):
            s = _RE_CELL_BULLET.sub(" · ", s)
            kind = "sep" if re.match(r"^\|[\s\-:|]+\|$", s) else "table"
            items.append((kind, s))
            continue

        # 제목
        if s.startswith("#"):
            # 제목 끝 로마숫자 → 앞으로 이동
            m = _RE_HEADING_SUFFIX.match(s)
            if m:
                hashes_and_text = m.group(1)
                roman = m.group(2).strip()
                parts = hashes_and_text.split(" ", 1)
                s = f"{parts[0]} {roman} {parts[1]}"
            # 번호 뒤 띄어쓰기
            m2 = _RE_HEADING_NUM_SPACE.match(s)
            if m2:
                after_num = s[m2.end():]
                if after_num and not after_num.startswith(" "):
                    s = s[:m2.end()] + " " + after_num
            items.append(("heading", s))
            continue

        # 줄 끝 글머리 제거
        s = _RE_TRAILING_BULLET.sub("", s).rstrip()
        if not s:
            continue

        # 글머리
        if _RE_BULLET_CHARS.match(s):
            s = "- " + _RE_BULLET_CHARS.sub("", s).lstrip()
            items.append(("bullet", s))
            continue

        # 볼드
        if s.startswith("**"):
            items.append(("bold", s))
            continue

        items.append(("para", s))

    # ── Pass 2: 연속 paragraph 줄 이어붙이기 ──
    _END_PUNCTS = (".", "。", "!", "?", "—", "…")
    _NEW_SENTENCE = re.compile(r"^(?:[0-9]+\.|쟁점|–)")

    merged: list[tuple[str, str]] = []
    for kind, text in items:
        if (kind == "para"
                and merged
                and merged[-1][0] == "para"
                and not merged[-1][1].endswith(_END_PUNCTS)
                and not _NEW_SENTENCE.match(text)):
            merged[-1] = ("para", merged[-1][1] + " " + text)
        else:
            merged.append((kind, text))

    # ── Pass 3: 최종 마크다운 생성 ──
    out_lines: list[str] = []
    prev_kind = ""
    for kind, text in merged:
        if kind == "sep":
            out_lines.append(text)
            prev_kind = kind
            continue
        if kind == "table" and prev_kind in ("table", "sep"):
            out_lines.append(text)
            prev_kind = kind
            continue
        if out_lines:
            out_lines.append("")
        out_lines.append(text)
        prev_kind = kind

    return "\n".join(out_lines).strip() + "\n"
