"""
HTML to PDF 변환 유틸리티

Chrome/Edge 브라우저의 headless 모드를 사용하여 HTML 파일을 PDF로 변환합니다.

사용법:
    python html2pdf.py input.html                      # 단일 파일 변환
    python html2pdf.py input.html -o output.pdf         # 출력 경로 지정
    python html2pdf.py ./docs/                          # 폴더 내 모든 HTML 일괄 변환
    python html2pdf.py ./docs/ -o ./pdfs/               # 출력 폴더 지정
    python html2pdf.py input.html --engine edge         # Edge 엔진 강제 사용
"""

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

from progress import ProgressBar

HTML_EXTENSIONS = {".html", ".htm"}


def find_browser() -> tuple[str, str] | None:
    """Chrome 또는 Edge 실행 파일 경로를 찾는다. (이름, 경로) 반환."""
    candidates = [
        ("chrome", r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        ("chrome", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        ("edge", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        ("edge", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    ]
    for name, path in candidates:
        if os.path.isfile(path):
            return name, path
    return None


def find_chrome() -> str | None:
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def find_edge() -> str | None:
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def convert_with_browser(input_path: str, output_path: str, browser_path: str) -> bool:
    """브라우저 headless 모드를 사용하여 PDF로 변환한다."""
    file_url = Path(input_path).as_uri()

    try:
        result = subprocess.run(
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


def convert_file(input_path: str, output_path: str, engine: str = "auto") -> bool:
    """단일 HTML 파일을 PDF로 변환한다."""
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    if engine == "chrome":
        browser_path = find_chrome()
        if not browser_path:
            print("  [오류] Chrome을 찾을 수 없습니다.")
            return False
        return convert_with_browser(input_path, output_path, browser_path)
    elif engine == "edge":
        browser_path = find_edge()
        if not browser_path:
            print("  [오류] Edge를 찾을 수 없습니다.")
            return False
        return convert_with_browser(input_path, output_path, browser_path)
    else:  # auto
        result = find_browser()
        if not result:
            print("  [오류] Chrome 또는 Edge를 찾을 수 없습니다.")
            return False
        _, browser_path = result
        return convert_with_browser(input_path, output_path, browser_path)


def collect_html_files(path: str) -> list[Path]:
    """경로에서 HTML 파일 목록을 수집한다."""
    p = Path(path).resolve()
    if p.is_file():
        if p.suffix.lower() in HTML_EXTENSIONS:
            return [p]
        else:
            print(f"[오류] 지원하지 않는 파일 형식: {p.suffix}")
            return []
    elif p.is_dir():
        files = sorted(
            f for f in p.rglob("*") if f.suffix.lower() in HTML_EXTENSIONS
        )
        return files
    else:
        print(f"[오류] 경로를 찾을 수 없습니다: {path}")
        return []


def main():
    parser = argparse.ArgumentParser(
        description="HTML 파일을 PDF로 변환합니다.",
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="변환할 HTML 파일 또는 폴더 경로",
    )
    parser.add_argument(
        "-o", "--output",
        help="출력 PDF 파일 또는 폴더 경로 (미지정 시 입력 파일과 같은 위치)",
    )
    parser.add_argument(
        "--engine",
        choices=["auto", "chrome", "edge"],
        default="auto",
        help="변환 엔진 선택 (기본값: auto - Chrome 우선, Edge 폴백)",
    )
    args = parser.parse_args()

    input_path = " ".join(args.input)

    files = collect_html_files(input_path)
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

        if not convert_file(str(file), str(out), engine=args.engine):
            failed.append(str(file))
    print(f"\n결과: {len(files) - len(failed)}/{len(files)} 성공")

    if failed:
        print(f"\n실패한 파일:")
        for f in failed:
            print(f"  - {f}")
        sys.exit(1)


if __name__ == "__main__":
    main()
