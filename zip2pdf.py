"""
ZIP → PDF 일괄 변환 유틸리티

ZIP 파일을 압축파일명의 폴더에 해제한 후, 모든 서브디렉토리를 재귀 탐색하여
HWP, HWPX, DOC, DOCX, XLS, XLSX, PPT, PPTX, HTML 파일을 PDF로 일괄 변환합니다.

사용법:
    python zip2pdf.py archive.zip                # ZIP 해제 후 일괄 변환 (같은 위치)
    python zip2pdf.py archive.zip -o ./pdfs/     # 출력 폴더 지정
    python zip2pdf.py ./zips/                    # 폴더 내 모든 ZIP 일괄 처리
"""

import argparse
import os
import sys
import time
import zipfile
from pathlib import Path

from progress import ProgressBar
from dir2pdf import collect_files, convert_file


def extract_zip(zip_path: str) -> str | None:
    """ZIP 파일을 압축파일명의 폴더에 해제한다. 해제된 폴더 경로를 반환한다."""
    zip_path = os.path.abspath(zip_path)
    zip_stem = Path(zip_path).stem
    extract_dir = os.path.join(os.path.dirname(zip_path), zip_stem)

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            # 인코딩 처리: CP949(한글 파일명) 대응
            for info in zf.infolist():
                try:
                    info.filename = info.filename.encode("cp437").decode("cp949")
                except (UnicodeDecodeError, UnicodeEncodeError):
                    pass  # 이미 UTF-8이거나 다른 인코딩인 경우 원본 유지
                info.filename = info.filename.replace("/", os.sep)
                zf.extract(info, extract_dir)

        return extract_dir

    except zipfile.BadZipFile:
        print(f"  [오류] 유효하지 않은 ZIP 파일: {zip_path}")
        return None
    except Exception as e:
        print(f"  [오류] ZIP 해제 실패: {e}")
        return None


def collect_zips(path: str) -> list[Path]:
    """경로에서 ZIP 파일 목록을 수집한다."""
    p = Path(path)
    if p.is_file():
        if p.suffix.lower() == ".zip":
            return [p]
        else:
            print(f"[오류] ZIP 파일이 아닙니다: {p.suffix}")
            return []
    elif p.is_dir():
        return sorted(f for f in p.rglob("*") if f.suffix.lower() == ".zip")
    else:
        print(f"[오류] 경로를 찾을 수 없습니다: {path}")
        return []


def main():
    parser = argparse.ArgumentParser(
        description="ZIP 파일을 해제하고 문서 파일을 PDF로 일괄 변환합니다.",
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="변환할 ZIP 파일 또는 ZIP 파일이 있는 폴더 경로",
    )
    parser.add_argument(
        "-o", "--output",
        help="출력 PDF 폴더 경로 (미지정 시 해제된 폴더와 같은 위치)",
    )
    parser.add_argument(
        "-libre",
        action="store_true",
        help="LibreOffice를 폴백 엔진으로 사용",
    )
    args = parser.parse_args()

    # 공백이 포함된 파일명 처리
    input_path = " ".join(args.input)

    zips = collect_zips(input_path)
    if not zips:
        print("처리할 ZIP 파일이 없습니다.")
        sys.exit(1)

    total_success = 0
    all_failed_files = []

    for zi, zip_file in enumerate(zips, 1):
        print(f"[ZIP {zi}/{len(zips)}] {zip_file.name}")

        # 1. ZIP 해제
        extract_dir = extract_zip(str(zip_file))
        if not extract_dir:
            continue

        # 2. 문서 파일 수집
        files = collect_files(extract_dir)
        if not files:
            print(f"  변환할 문서 파일이 없습니다.")
            continue

        # 3. PDF 변환
        pbar = ProgressBar(len(files))
        for file in files:
            pbar.update(file.name)

            if args.output:
                rel = file.relative_to(Path(extract_dir).resolve())
                out = Path(args.output) / zip_file.stem / rel.with_suffix(".pdf")
            else:
                out = file.with_suffix(".pdf")

            if convert_file(str(file), str(out), libre=args.libre):
                total_success += 1
            else:
                all_failed_files.append(str(file))

    total = total_success + len(all_failed_files)
    print(f"\n결과: {total_success}/{total} 성공")

    if all_failed_files:
        print(f"\n실패한 파일:")
        for f in all_failed_files:
            print(f"  - {f}")
        sys.exit(1)


if __name__ == "__main__":
    main()
