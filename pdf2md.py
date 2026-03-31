"""
PDF to Markdown 변환 유틸리티

pdfplumber를 사용하여 PDF 파일을 구조적 Markdown으로 변환합니다.
폰트 크기로 제목을 추론하고, 표를 자동 감지하며, 아티팩트를 정리합니다.

사용법:
    python pdf2md.py input.pdf                    # 단일 파일 변환
    python pdf2md.py input.pdf -o output.md        # 출력 경로 지정
    python pdf2md.py ./docs/                       # 폴더 내 모든 PDF 일괄 변환
    python pdf2md.py ./docs/ -o ./markdown/        # 출력 폴더 지정
"""

import argparse
import os
import sys
from pathlib import Path

from pdfparse import pdf_to_markdown
from progress import ProgressBar


def convert_file(input_path: str, output_path: str) -> bool:
    """단일 PDF 파일을 Markdown으로 변환한다."""
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    try:
        md = pdf_to_markdown(input_path)
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(md)
        return True
    except Exception as e:
        print(f"  [오류] 변환 실패: {e}")
        return False


def collect_pdf_files(path: str) -> list[Path]:
    """경로에서 PDF 파일 목록을 수집한다."""
    p = Path(path).resolve()
    if p.is_file():
        if p.suffix.lower() == ".pdf":
            return [p]
        print(f"[오류] 지원하지 않는 파일 형식: {p.suffix}")
        return []
    if p.is_dir():
        return sorted(f for f in p.rglob("*") if f.suffix.lower() == ".pdf")
    print(f"[오류] 경로를 찾을 수 없습니다: {path}")
    return []


def main():
    parser = argparse.ArgumentParser(
        description="PDF 파일을 Markdown으로 변환합니다.",
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="변환할 PDF 파일 또는 폴더 경로",
    )
    parser.add_argument(
        "-o", "--output",
        help="출력 Markdown 파일 또는 폴더 경로 (미지정 시 입력 파일과 같은 위치)",
    )
    args = parser.parse_args()

    input_path = " ".join(args.input)
    files = collect_pdf_files(input_path)
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

        if not convert_file(str(file), str(out)):
            failed.append(str(file))

    print(f"\n결과: {len(files) - len(failed)}/{len(files)} 성공")

    if failed:
        print("\n실패한 파일:")
        for f in failed:
            print(f"  - {f}")
        sys.exit(1)


if __name__ == "__main__":
    main()
