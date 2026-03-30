"""
PPT/PPTX to PDF 변환 유틸리티

Microsoft PowerPoint COM 자동화를 사용하여 PPT/PPTX 파일을 PDF로 변환합니다.
PowerPoint가 없는 경우 LibreOffice를 폴백으로 사용합니다.

사용법:
    python ppt2pdf.py input.pptx                      # 단일 파일 변환
    python ppt2pdf.py input.pptx -o output.pdf         # 출력 경로 지정
    python ppt2pdf.py ./docs/                          # 폴더 내 모든 PPT/PPTX 일괄 변환
    python ppt2pdf.py ./docs/ -o ./pdfs/               # 출력 폴더 지정
    python ppt2pdf.py input.pptx --engine libreoffice  # LibreOffice 엔진 강제 사용
"""

import argparse
import os
import sys
import time
from pathlib import Path

from progress import ProgressBar

PPT_EXTENSIONS = {".ppt", ".pptx"}


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


def convert_with_powerpoint(input_path: str, output_path: str) -> bool:
    """Microsoft PowerPoint COM 자동화를 사용하여 PDF로 변환한다."""
    try:
        import win32com.client
    except ImportError:
        print("  [오류] pywin32가 설치되어 있지 않습니다: pip install pywin32")
        return False

    ppt = None
    presentation = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")

        # ppOpen = 2 (ReadOnly)
        presentation = ppt.Presentations.Open(input_path, ReadOnly=True, WithWindow=False)

        # ppSaveAsPDF = 32
        presentation.SaveAs(output_path, FileFormat=32)

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


def convert_with_libreoffice(input_path: str, output_path: str) -> bool:
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
                "--convert-to", "pdf:impress_pdf_Export",
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

        # LibreOffice는 원본 파일명.pdf로 저장하므로 필요시 이름 변경
        expected = os.path.join(
            output_dir,
            Path(input_path).stem + ".pdf",
        )
        if expected != output_path and os.path.isfile(expected):
            os.replace(expected, output_path)

        return os.path.isfile(output_path)

    except subprocess.TimeoutExpired:
        print("  [오류] LibreOffice 변환 시간 초과 (120초)")
        return False
    except Exception as e:
        print(f"  [오류] LibreOffice 변환 실패: {e}")
        return False


def convert_file(input_path: str, output_path: str, engine: str = "auto") -> bool:
    """단일 파일을 PDF로 변환한다."""
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    # 출력 디렉토리 생성
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    if engine == "powerpoint":
        return convert_with_powerpoint(input_path, output_path)
    elif engine == "libreoffice":
        return convert_with_libreoffice(input_path, output_path)
    else:  # auto
        if convert_with_powerpoint(input_path, output_path):
            return True
        return convert_with_libreoffice(input_path, output_path)


def collect_ppt_files(path: str) -> list[Path]:
    """경로에서 PPT/PPTX 파일 목록을 수집한다."""
    p = Path(path).resolve()
    if p.is_file():
        if p.suffix.lower() in PPT_EXTENSIONS:
            return [p]
        else:
            print(f"[오류] 지원하지 않는 파일 형식: {p.suffix}")
            return []
    elif p.is_dir():
        files = sorted(
            f for f in p.rglob("*")
            if f.suffix.lower() in PPT_EXTENSIONS and not f.name.startswith("~$")
        )
        return files
    else:
        print(f"[오류] 경로를 찾을 수 없습니다: {path}")
        return []


def main():
    parser = argparse.ArgumentParser(
        description="PPT/PPTX 파일을 PDF로 변환합니다.",
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="변환할 PPT/PPTX 파일 또는 폴더 경로",
    )
    parser.add_argument(
        "-o", "--output",
        help="출력 PDF 파일 또는 폴더 경로 (미지정 시 입력 파일과 같은 위치)",
    )
    parser.add_argument(
        "--engine",
        choices=["auto", "powerpoint", "libreoffice"],
        default="auto",
        help="변환 엔진 선택 (기본값: auto - PowerPoint 우선, LibreOffice 폴백)",
    )
    args = parser.parse_args()

    # 공백이 포함된 파일명 처리: 인수들을 하나로 합침
    input_path = " ".join(args.input)

    files = collect_ppt_files(input_path)
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
