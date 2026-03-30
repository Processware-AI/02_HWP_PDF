"""
문서 → PDF 일괄 변환 유틸리티

디렉토리를 재귀 탐색하여 HWP, HWPX, DOC, DOCX, XLS, XLSX, PPT, PPTX, HTML 파일을
모두 PDF로 변환합니다.

사용법:
    python dir2pdf.py ./docs/                # 폴더 내 모든 문서 일괄 변환
    python dir2pdf.py ./docs/ -o ./pdfs/     # 출력 폴더 지정
    python dir2pdf.py report.hwp             # 단일 파일 변환
"""

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

from progress import ProgressBar

ALL_EXTENSIONS = {".hwp", ".hwpx", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".html", ".htm"}
HTML_EXTENSIONS = {".html", ".htm"}


def find_libreoffice() -> str | None:
    """LibreOffice 실행 파일 경로를 찾는다."""
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def convert_with_libreoffice(input_path: str, output_path: str, filter_name: str) -> bool:
    """LibreOffice headless 모드를 사용하여 PDF로 변환한다."""
    import subprocess

    soffice = find_libreoffice()
    if not soffice:
        print("  [오류] LibreOffice를 찾을 수 없습니다.")
        return False

    output_dir = os.path.dirname(output_path) or "."
    try:
        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--convert-to", f"pdf:{filter_name}",
                "--outdir", output_dir,
                input_path,
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )

        if result.returncode != 0:
            print(f"  [오류] LibreOffice 변환 실패: {result.stderr}")
            return False

        expected = os.path.join(output_dir, Path(input_path).stem + ".pdf")
        if expected != output_path and os.path.isfile(expected):
            os.replace(expected, output_path)

        return os.path.isfile(output_path)

    except subprocess.TimeoutExpired:
        print("  [오류] LibreOffice 변환 시간 초과 (120초)")
        return False
    except Exception as e:
        print(f"  [오류] LibreOffice 변환 실패: {e}")
        return False


# ── HWP/HWPX ──

def convert_hwp(input_path: str, output_path: str) -> bool:
    try:
        import win32com.client
    except ImportError:
        return convert_with_libreoffice(input_path, output_path, "writer_pdf_Export")

    hwp = None
    try:
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = False
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        except Exception:
            pass

        fmt = "HWPX" if input_path.lower().endswith(".hwpx") else "HWP"
        if not hwp.Open(input_path, fmt, "forceopen:true"):
            print(f"  [오류] 파일을 열 수 없습니다: {input_path}")
            return False

        hwp.SaveAs(output_path, "PDF")
        return os.path.isfile(output_path)

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


# ── DOC/DOCX ──

def convert_doc(input_path: str, output_path: str) -> bool:
    try:
        import win32com.client
    except ImportError:
        return convert_with_libreoffice(input_path, output_path, "writer_pdf_Export")

    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(input_path, ReadOnly=True)
        doc.SaveAs2(output_path, FileFormat=17)  # wdFormatPDF = 17
        return os.path.isfile(output_path)

    except Exception as e:
        print(f"  [오류] Word 변환 실패: {e}")
        return False
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass


# ── XLS/XLSX ──

def convert_xls(input_path: str, output_path: str) -> bool:
    try:
        import win32com.client
    except ImportError:
        return convert_with_libreoffice(input_path, output_path, "calc_pdf_Export")

    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(input_path, ReadOnly=True)
        wb.Sheets.Select()
        wb.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=output_path,
            Quality=0,
            IncludeDocProperties=True,
            OpenAfterPublish=False,
        )
        return os.path.isfile(output_path)

    except Exception as e:
        print(f"  [오류] Excel 변환 실패: {e}")
        return False
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass


# ── PPT/PPTX ──

def convert_ppt(input_path: str, output_path: str) -> bool:
    try:
        import win32com.client
    except ImportError:
        return convert_with_libreoffice(input_path, output_path, "impress_pdf_Export")

    ppt = None
    presentation = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        presentation = ppt.Presentations.Open(input_path, ReadOnly=True, WithWindow=False)
        presentation.SaveAs(output_path, FileFormat=32)  # ppSaveAsPDF = 32
        return os.path.isfile(output_path)

    except Exception as e:
        print(f"  [오류] PowerPoint 변환 실패: {e}")
        return False
    finally:
        if presentation:
            try:
                presentation.Close()
            except Exception:
                pass
        if ppt:
            try:
                ppt.Quit()
            except Exception:
                pass


# ── HTML ──

def find_browser() -> str | None:
    """Chrome 또는 Edge 실행 파일 경로를 찾는다."""
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def convert_html(input_path: str, output_path: str) -> bool:
    browser_path = find_browser()
    if not browser_path:
        print("  [오류] Chrome 또는 Edge를 찾을 수 없습니다.")
        return False

    file_url = Path(input_path).as_uri()
    try:
        subprocess.run(
            [
                browser_path,
                "--headless",
                "--disable-gpu",
                "--no-sandbox",
                "--disable-extensions",
                f"--print-to-pdf={output_path}",
                "--print-to-pdf-no-header",
                file_url,
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        return os.path.isfile(output_path)

    except subprocess.TimeoutExpired:
        print("  [오류] 변환 시간 초과 (60초)")
        return False
    except Exception as e:
        print(f"  [오류] 브라우저 변환 실패: {e}")
        return False


# ── 변환 디스패치 ──

CONVERTERS = {
    ".hwp": convert_hwp,
    ".hwpx": convert_hwp,
    ".doc": convert_doc,
    ".docx": convert_doc,
    ".xls": convert_xls,
    ".xlsx": convert_xls,
    ".ppt": convert_ppt,
    ".pptx": convert_ppt,
    ".html": convert_html,
    ".htm": convert_html,
}


def convert_file(input_path: str, output_path: str) -> bool:
    """파일 확장자에 따라 적절한 변환기를 선택하여 PDF로 변환한다."""
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    ext = Path(input_path).suffix.lower()
    converter = CONVERTERS.get(ext)
    if not converter:
        return False

    return converter(input_path, output_path)


def collect_files(path: str) -> list[Path]:
    """경로에서 지원되는 문서 파일 목록을 수집한다."""
    p = Path(path).resolve()
    if p.is_file():
        if p.suffix.lower() in ALL_EXTENSIONS:
            return [p]
        else:
            print(f"[오류] 지원하지 않는 파일 형식: {p.suffix}")
            return []
    elif p.is_dir():
        files = sorted(
            f for f in p.rglob("*")
            if f.suffix.lower() in ALL_EXTENSIONS and not f.name.startswith("~$")
        )
        return files
    else:
        print(f"[오류] 경로를 찾을 수 없습니다: {path}")
        return []


def main():
    parser = argparse.ArgumentParser(
        description="HWP/HWPX, DOC/DOCX, XLS/XLSX, PPT/PPTX, HTML 파일을 PDF로 일괄 변환합니다.",
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="변환할 파일 또는 폴더 경로",
    )
    parser.add_argument(
        "-o", "--output",
        help="출력 PDF 파일 또는 폴더 경로 (미지정 시 입력 파일과 같은 위치)",
    )
    args = parser.parse_args()

    # 공백이 포함된 파일명 처리: 인수들을 하나로 합침
    input_path = " ".join(args.input)

    files = collect_files(input_path)
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
                out = Path(args.output) / rel.with_suffix(".pdf")
            else:
                out = Path(args.output)
        else:
            out = file.with_suffix(".pdf")

        if not convert_file(str(file), str(out)):
            failed.append(str(file))
    print(f"\n결과: {len(files) - len(failed)}/{len(files)} 성공")

    if failed:
        print(f"\n실패한 파일:")
        for f in failed:
            print(f"  - {f}")
        sys.exit(1)


if __name__ == "__main__":
    main()
