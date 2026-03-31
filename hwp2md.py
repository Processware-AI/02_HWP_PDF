"""
HWP/HWPX to Markdown 변환 유틸리티

한컴오피스(한글) COM 자동화를 사용하여 HWP/HWPX 파일을 Markdown으로 변환합니다.
COM이 없는 경우 HWPX는 XML 직접 파싱, HWP는 바이너리 파싱으로 폴백합니다.

사용법:
    python hwp2md.py input.hwp                    # 단일 파일 변환
    python hwp2md.py input.hwp -o output.md        # 출력 경로 지정
    python hwp2md.py ./docs/                       # 폴더 내 모든 HWP/HWPX 일괄 변환
    python hwp2md.py ./docs/ -o ./markdown/        # 출력 폴더 지정
    python hwp2md.py input.hwp --engine direct     # 직접 파싱 엔진 강제 사용
"""

import argparse
import os
import re
import struct
import sys
import tempfile
import zlib
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET

from progress import ProgressBar

HWP_EXTENSIONS = {".hwp", ".hwpx"}


# ---------------------------------------------------------------------------
# HTML → Markdown 변환기
# ---------------------------------------------------------------------------

class _HtmlToMarkdown(HTMLParser):
    """간단한 HTML → Markdown 변환기."""

    # 내용을 완전히 무시할 태그
    _SKIP_TAGS = frozenset({"style", "script", "noscript"})

    def __init__(self):
        super().__init__()
        self._parts: list[str] = []
        self._tag_stack: list[str] = []
        self._list_stack: list[tuple[str, int]] = []
        self._skip_depth = 0  # _SKIP_TAGS 중첩 깊이
        self._cell_index = 0
        self._row_index = 0  # 테이블 내 행 번호
        self._in_pre = False
        self._href: str | None = None

    def _is_skipping(self) -> bool:
        return self._skip_depth > 0

    def handle_starttag(self, tag, attrs):
        # skip 태그 진입
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._is_skipping():
            return

        attrs_dict = dict(attrs)
        self._tag_stack.append(tag)

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag[1])
            self._parts.append("\n\n" + "#" * level + " ")
        elif tag == "p":
            self._parts.append("\n\n")
        elif tag == "br":
            self._parts.append("  \n")
        elif tag == "hr":
            self._parts.append("\n\n---\n\n")
        elif tag in ("strong", "b"):
            self._parts.append("**")
        elif tag in ("em", "i"):
            self._parts.append("*")
        elif tag == "code" and not self._in_pre:
            self._parts.append("`")
        elif tag == "pre":
            self._in_pre = True
            self._parts.append("\n\n```\n")
        elif tag == "ul":
            self._list_stack.append(("ul", 0))
            self._parts.append("\n")
        elif tag == "ol":
            self._list_stack.append(("ol", 0))
            self._parts.append("\n")
        elif tag == "li":
            if self._list_stack:
                kind, count = self._list_stack[-1]
                indent = "  " * (len(self._list_stack) - 1)
                if kind == "ul":
                    self._parts.append(f"{indent}- ")
                else:
                    count += 1
                    self._list_stack[-1] = (kind, count)
                    self._parts.append(f"{indent}{count}. ")
        elif tag == "a":
            self._href = attrs_dict.get("href", "")
            self._parts.append("[")
        elif tag == "img":
            alt = attrs_dict.get("alt", "")
            src = attrs_dict.get("src", "")
            self._parts.append(f"![{alt}]({src})")
        elif tag == "table":
            self._parts.append("\n\n")
            self._row_index = 0
        elif tag == "tr":
            self._cell_index = 0
        elif tag in ("td", "th"):
            self._parts.append("| ")
        elif tag == "blockquote":
            self._parts.append("\n\n> ")

    def handle_endtag(self, tag):
        # skip 태그 탈출
        if tag in self._SKIP_TAGS:
            if self._skip_depth > 0:
                self._skip_depth -= 1
            return
        if self._is_skipping():
            return

        if self._tag_stack and self._tag_stack[-1] == tag:
            self._tag_stack.pop()

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._parts.append("\n")
        elif tag in ("strong", "b"):
            self._parts.append("**")
        elif tag in ("em", "i"):
            self._parts.append("*")
        elif tag == "code" and not self._in_pre:
            self._parts.append("`")
        elif tag == "pre":
            self._in_pre = False
            self._parts.append("\n```\n")
        elif tag in ("ul", "ol"):
            if self._list_stack:
                self._list_stack.pop()
            self._parts.append("\n")
        elif tag == "li":
            self._parts.append("\n")
        elif tag == "a":
            self._parts.append(f"]({self._href})")
            self._href = None
        elif tag in ("td", "th"):
            self._parts.append(" ")
            self._cell_index += 1
        elif tag == "tr":
            self._parts.append("|\n")
            self._row_index += 1
            # 첫 번째 행 이후 구분선 삽입
            if self._row_index == 1 and self._cell_index > 0:
                self._parts.append("|" + " --- |" * self._cell_index + "\n")
        elif tag == "thead":
            # thead가 명시적으로 있는 경우에도 구분선 삽입
            if self._cell_index > 0:
                self._parts.append("|" + " --- |" * self._cell_index + "\n")

    def handle_data(self, data):
        if self._is_skipping():
            return
        if self._in_pre:
            self._parts.append(data)
        else:
            self._parts.append(re.sub(r"\s+", " ", data))

    def get_markdown(self) -> str:
        text = "".join(self._parts)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip() + "\n"


def _read_html_with_encoding(path: str) -> str:
    """HTML 파일을 올바른 인코딩으로 읽는다.

    한컴오피스는 HTML을 EUC-KR(CP949)로 저장하는 경우가 많으므로
    charset 메타 태그를 확인하거나 여러 인코딩을 시도한다.
    """
    raw = open(path, "rb").read()

    # 1) charset 메타 태그에서 인코딩 감지
    #    <meta charset="euc-kr"> 또는 <meta ... content="text/html; charset=euc-kr">
    head = raw[:4096].lower()
    m = re.search(rb'charset[="\s]+([a-z0-9_-]+)', head)
    if m:
        declared = m.group(1).decode("ascii", errors="ignore")
        try:
            return raw.decode(declared)
        except (UnicodeDecodeError, LookupError):
            pass

    # 2) UTF-8 시도
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        pass

    # 3) CP949 (EUC-KR 상위호환) 시도
    try:
        return raw.decode("cp949")
    except UnicodeDecodeError:
        pass

    # 4) 최후 수단: UTF-8 + replace
    return raw.decode("utf-8", errors="replace")


def _html_to_markdown(html: str) -> str:
    """HTML 문자열을 Markdown으로 변환한다."""
    parser = _HtmlToMarkdown()
    parser.feed(html)
    return parser.get_markdown()


# ---------------------------------------------------------------------------
# 한컴오피스 COM 변환
# ---------------------------------------------------------------------------

def _convert_with_hancom(input_path: str, output_path: str) -> bool:
    """한컴오피스 COM을 통해 HTML로 저장 후 Markdown으로 변환한다."""
    try:
        import win32com.client
    except ImportError:
        return False

    hwp = None
    try:
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = False
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

        fmt = "HWPX" if input_path.lower().endswith(".hwpx") else "HWP"
        if not hwp.Open(input_path, fmt, "forceopen:true"):
            print(f"  [오류] 파일을 열 수 없습니다: {input_path}")
            return False

        with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
            tmp_html = tmp.name

        try:
            hwp.SaveAs(tmp_html, "HTML")
            if not os.path.isfile(tmp_html):
                return False

            html = _read_html_with_encoding(tmp_html)

            md = _html_to_markdown(html)
            os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md)
            return True
        finally:
            if os.path.isfile(tmp_html):
                os.unlink(tmp_html)

    except Exception as e:
        print(f"  [오류] 한컴오피스 변환 실패: {e}")
        return False
    finally:
        if hwp:
            try:
                hwp.Clear(1)
                hwp.Quit()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# HWPX 직접 파싱 (XML 기반)
# ---------------------------------------------------------------------------

def _convert_hwpx_direct(input_path: str, output_path: str) -> bool:
    """HWPX 파일을 XML 파싱하여 Markdown으로 변환한다."""
    import zipfile

    try:
        with zipfile.ZipFile(input_path, "r") as zf:
            names = zf.namelist()

            # 개요 스타일 맵 구축 (styleIDRef → heading level)
            heading_map: dict[str, int] = {}
            for name in names:
                if name.lower().endswith(".xml") and "head" in name.lower():
                    try:
                        tree = ET.fromstring(zf.read(name))
                        for elem in tree.iter():
                            local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                            if local == "style":
                                sid = elem.get("id", "")
                                ol = elem.get("outlineLevel")
                                if ol and ol.isdigit() and 1 <= int(ol) <= 6:
                                    heading_map[sid] = int(ol)
                    except Exception:
                        pass

            # 본문 섹션 XML 수집
            section_files = sorted(
                n for n in names
                if n.lower().startswith("contents/") and n.lower().endswith(".xml")
                and "head" not in n.lower()
            )
            if not section_files:
                section_files = sorted(
                    n for n in names
                    if "section" in n.lower() and n.endswith(".xml")
                )
            if not section_files:
                print("  [오류] HWPX 본문 섹션을 찾을 수 없습니다.")
                return False

            md_parts: list[str] = []
            for sf in section_files:
                try:
                    root = ET.fromstring(zf.read(sf))
                except ET.ParseError:
                    continue

                for elem in root.iter():
                    local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                    if local != "p":
                        continue

                    # 단락 내 모든 <t> 텍스트 수집
                    texts = []
                    for sub in elem.iter():
                        sl = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                        if sl == "t":
                            if sub.text:
                                texts.append(sub.text)
                            if sub.tail:
                                texts.append(sub.tail)

                    text = "".join(texts).strip()
                    if not text:
                        md_parts.append("")
                        continue

                    style_ref = elem.get("styleIDRef", "")
                    level = heading_map.get(style_ref, 0)
                    if level:
                        md_parts.append(f"{'#' * level} {text}")
                    else:
                        md_parts.append(text)

            if not any(md_parts):
                print("  [오류] 텍스트를 추출할 수 없습니다.")
                return False

            md_text = "\n\n".join(md_parts)
            md_text = re.sub(r"\n{3,}", "\n\n", md_text).strip() + "\n"

            os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md_text)
            return True

    except zipfile.BadZipFile:
        print(f"  [오류] 유효하지 않은 HWPX 파일: {input_path}")
        return False
    except Exception as e:
        print(f"  [오류] HWPX 파싱 실패: {e}")
        return False


# ---------------------------------------------------------------------------
# HWP 바이너리 파싱 (OLE2)
# ---------------------------------------------------------------------------

HWPTAG_PARA_TEXT = 67  # HWPTAG_BEGIN(16) + 51

# 확장 제어 문자 (14바이트 추가 데이터를 가짐)
_HWP_EXTENDED_CHARS = frozenset({
    1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 21, 22, 23,
})


def _extract_text_from_section(data: bytes) -> list[str]:
    """HWP 섹션 바이너리에서 텍스트 레코드를 추출한다."""
    texts = []
    pos = 0
    while pos + 4 <= len(data):
        header = struct.unpack_from("<I", data, pos)[0]
        tag_id = header & 0x3FF
        size = (header >> 20) & 0xFFF
        pos += 4

        if size == 0xFFF:
            if pos + 4 > len(data):
                break
            size = struct.unpack_from("<I", data, pos)[0]
            pos += 4

        if pos + size > len(data):
            break

        if tag_id == HWPTAG_PARA_TEXT:
            text_data = data[pos:pos + size]
            chars: list[str] = []
            i = 0
            while i + 1 < len(text_data):
                code = struct.unpack_from("<H", text_data, i)[0]
                i += 2
                if code == 0:
                    break
                if code in _HWP_EXTENDED_CHARS:
                    i += 14  # 확장 제어 문자 데이터 건너뛰기
                elif code == 10:  # 줄바꿈
                    chars.append("\n")
                elif code == 24:
                    chars.append("-")
                elif code == 30 or code == 31:
                    chars.append(" ")
                elif code >= 32:
                    chars.append(chr(code))

            line = "".join(chars).strip()
            if line:
                texts.append(line)

        pos += size

    return texts


def _convert_hwp_direct(input_path: str, output_path: str) -> bool:
    """HWP 바이너리 파일을 직접 파싱하여 Markdown으로 변환한다."""
    try:
        import olefile
    except ImportError:
        print("  [오류] olefile이 설치되어 있지 않습니다: pip install olefile")
        return False

    try:
        ole = olefile.OleFileIO(input_path)
    except Exception as e:
        print(f"  [오류] HWP 파일을 열 수 없습니다: {e}")
        return False

    try:
        # FileHeader에서 압축 여부 확인
        compressed = True
        if ole.exists("FileHeader"):
            header = ole.openstream("FileHeader").read()
            if len(header) >= 40:
                flags = struct.unpack_from("<I", header, 36)[0]
                compressed = bool(flags & 0x01)

        all_texts: list[str] = []
        section_idx = 0
        while True:
            stream_name = f"BodyText/Section{section_idx}"
            if not ole.exists(stream_name):
                break
            raw = ole.openstream(stream_name).read()
            try:
                section_data = zlib.decompress(raw, -15) if compressed else raw
            except zlib.error:
                section_data = raw

            all_texts.extend(_extract_text_from_section(section_data))
            section_idx += 1

        if not all_texts:
            print("  [오류] 텍스트를 추출할 수 없습니다.")
            return False

        md_text = "\n\n".join(all_texts)
        md_text = re.sub(r"\n{3,}", "\n\n", md_text).strip() + "\n"

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(md_text)
        return True

    except Exception as e:
        print(f"  [오류] HWP 파싱 실패: {e}")
        return False
    finally:
        ole.close()


# ---------------------------------------------------------------------------
# PDF 경유 변환 (HWP → PDF → Markdown)
# ---------------------------------------------------------------------------

def _convert_via_pdf(input_path: str, output_path: str) -> bool:
    """HWP → PDF (한컴 COM) → Markdown (pdfplumber) 경로로 변환한다."""
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        print("  [오류] pdfplumber가 설치되어 있지 않습니다: pip install pdfplumber")
        return False

    # 1단계: HWP → PDF (임시 파일)
    try:
        import win32com.client
    except ImportError:
        print("  [오류] pywin32 미설치 - PDF 경유 변환 불가")
        return False

    hwp = None
    try:
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = False
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

        fmt = "HWPX" if input_path.lower().endswith(".hwpx") else "HWP"
        if not hwp.Open(input_path, fmt, "forceopen:true"):
            print(f"  [오류] 파일을 열 수 없습니다: {input_path}")
            return False

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_pdf = tmp.name

        try:
            hwp.SaveAs(tmp_pdf, "PDF")
            if not os.path.isfile(tmp_pdf):
                return False

            # 2단계: PDF → Markdown
            from pdfparse import pdf_to_markdown
            md = pdf_to_markdown(tmp_pdf)
            os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md)
            return True
        finally:
            if os.path.isfile(tmp_pdf):
                os.unlink(tmp_pdf)

    except Exception as e:
        print(f"  [오류] PDF 경유 변환 실패: {e}")
        return False
    finally:
        if hwp:
            try:
                hwp.Clear(1)
                hwp.Quit()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 통합 변환
# ---------------------------------------------------------------------------

def _prefixed_path(output_path: str, prefix: str) -> str:
    """출력 경로의 파일명 앞에 접두사를 붙인다."""
    p = Path(output_path)
    return str(p.with_name(prefix + p.name))


def convert_file(input_path: str, output_path: str, engine: str = "auto") -> bool:
    """단일 HWP/HWPX 파일을 Markdown으로 변환한다."""
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    is_hwpx = input_path.lower().endswith(".hwpx")

    if engine == "both":
        # COM, OLE, PDF 세 가지 동시 실행, 접두사 붙여서 저장
        results = []
        for prefix, label, fn in [
            ("COM_", "COM", lambda p: _convert_with_hancom(input_path, p)),
            ("OLE_", "OLE", lambda p: (_convert_hwpx_direct(input_path, p) if is_hwpx
                                       else _convert_hwp_direct(input_path, p))),
            ("PDF_", "PDF", lambda p: _convert_via_pdf(input_path, p)),
        ]:
            out = _prefixed_path(output_path, prefix)
            ok = fn(out)
            print(f"  [{label}] {Path(out).name}" if ok else f"  [{label}] 실패")
            results.append(ok)
        return any(results)

    if engine == "hancom":
        return _convert_with_hancom(input_path, output_path)
    elif engine == "direct":
        return _convert_hwpx_direct(input_path, output_path) if is_hwpx else _convert_hwp_direct(input_path, output_path)
    elif engine == "pdf":
        return _convert_via_pdf(input_path, output_path)

    # auto: 한컴 우선, 직접 파싱 폴백
    if _convert_with_hancom(input_path, output_path):
        return True
    print("  [정보] 직접 파싱으로 폴백합니다.")
    return _convert_hwpx_direct(input_path, output_path) if is_hwpx else _convert_hwp_direct(input_path, output_path)


def collect_hwp_files(path: str) -> list[Path]:
    """경로에서 HWP/HWPX 파일 목록을 수집한다."""
    p = Path(path).resolve()
    if p.is_file():
        if p.suffix.lower() in HWP_EXTENSIONS:
            return [p]
        print(f"[오류] 지원하지 않는 파일 형식: {p.suffix}")
        return []
    if p.is_dir():
        return sorted(f for f in p.rglob("*") if f.suffix.lower() in HWP_EXTENSIONS)
    print(f"[오류] 경로를 찾을 수 없습니다: {path}")
    return []


def main():
    parser = argparse.ArgumentParser(
        description="HWP/HWPX 파일을 Markdown으로 변환합니다.",
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="변환할 HWP/HWPX 파일 또는 폴더 경로",
    )
    parser.add_argument(
        "-o", "--output",
        help="출력 Markdown 파일 또는 폴더 경로 (미지정 시 입력 파일과 같은 위치)",
    )
    parser.add_argument(
        "--engine",
        choices=["auto", "hancom", "direct", "pdf", "both"],
        default="auto",
        help="변환 엔진 선택 (auto: 한컴→폴백 / pdf: HWP→PDF→MD / both: COM_·OLE_·PDF_ 동시 출력)",
    )
    args = parser.parse_args()

    input_path = " ".join(args.input)
    files = collect_hwp_files(input_path)
    if not files:
        print("변환할 파일이 없습니다.")
        sys.exit(1)

    input_is_dir = Path(input_path).is_dir()
    failed = []

    pbar = ProgressBar(len(files))
    for file in files:
        pbar.update(file.name)

        if args.output:
            if input_is_dir:
                rel = file.relative_to(Path(input_path).resolve())
                out = Path(args.output) / rel.with_suffix(".md")
            else:
                out = Path(args.output)
        else:
            out = file.with_suffix(".md")

        if not convert_file(str(file), str(out), engine=args.engine):
            failed.append(str(file))

    print(f"\n결과: {len(files) - len(failed)}/{len(files)} 성공")

    if failed:
        print("\n실패한 파일:")
        for f in failed:
            print(f"  - {f}")
        sys.exit(1)


if __name__ == "__main__":
    main()
