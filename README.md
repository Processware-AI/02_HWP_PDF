# 문서 → PDF 변환 유틸리티

Devloped by 오동석 with Claude MAX

HWP/HWPX, DOC/DOCX, XLS/XLSX, PPT/PPTX, HTML 파일을 PDF로 변환하는 CLI 도구 모음입니다.

## 요구 사항

- Windows 10/11
- Python 3.10+
- [pywin32](https://pypi.org/project/pywin32/) (`pip install pywin32`)
- 한컴오피스(한글) 2018 이상 — `hwp2pdf.py` 사용 시
- Microsoft Word — `doc2pdf.py` 사용 시
- Microsoft Excel — `xls2pdf.py` 사용 시
- Microsoft PowerPoint — `ppt2pdf.py` 사용 시
- Chrome 또는 Edge — `html2pdf.py` 사용 시
- (선택) LibreOffice — 위 프로그램 미설치 시 폴백 엔진

---

## zip2pdf.py — ZIP 해제 후 PDF 일괄 변환

ZIP 파일을 압축파일명의 폴더에 해제한 후, 모든 서브디렉토리를 재귀 탐색하여 문서 파일을 PDF로 일괄 변환합니다. 한글 파일명이 포함된 ZIP도 정상 처리됩니다.

### 사용법

```bash
python zip2pdf.py archive.zip                # ZIP 해제 후 일괄 변환 (같은 위치)
python zip2pdf.py archive.zip -o ./pdfs/     # 출력 폴더 지정
python zip2pdf.py ./zips/                    # 폴더 내 모든 ZIP 일괄 처리
```

### 옵션

| 옵션 | 설명 |
|------|------|
| `input` | ZIP 파일 또는 ZIP 파일이 있는 폴더 경로 |
| `-o`, `--output` | 출력 PDF 폴더 경로 (미지정 시 해제된 폴더와 같은 위치) |

### 동작 순서

1. ZIP 파일을 `압축파일명/` 폴더에 해제
2. 해제된 폴더를 재귀 탐색하여 지원 문서 파일 수집
3. 파일 확장자에 따라 적절한 엔진으로 PDF 변환

---

## dir2pdf.py — 모든 문서 → PDF 통합 변환

디렉토리를 재귀 탐색하여 모든 지원 문서 파일을 PDF로 일괄 변환합니다. 파일 확장자에 따라 적절한 변환 엔진(한컴오피스, Word, Excel, PowerPoint)을 자동 선택합니다.

### 사용법

```bash
python dir2pdf.py ./docs/                  # 폴더 내 모든 문서 일괄 변환
python dir2pdf.py ./docs/ -o ./pdfs/       # 출력 폴더 지정
python dir2pdf.py report.hwp               # 단일 파일 변환
```

### 옵션

| 옵션 | 설명 |
|------|------|
| `input` | 변환할 파일 또는 폴더 경로 |
| `-o`, `--output` | 출력 PDF 파일 또는 폴더 경로 (미지정 시 원본과 같은 위치) |

### 지원 파일 형식

| 확장자 | 변환 엔진 |
|--------|-----------|
| `.hwp`, `.hwpx` | 한컴오피스 (폴백: LibreOffice) |
| `.doc`, `.docx` | Microsoft Word (폴백: LibreOffice) |
| `.xls`, `.xlsx` | Microsoft Excel (폴백: LibreOffice) |
| `.ppt`, `.pptx` | Microsoft PowerPoint (폴백: LibreOffice) |

> 임시 파일(`~$`로 시작하는 파일)은 자동으로 제외됩니다.

---

## hwp2pdf.py — HWP/HWPX → PDF

한컴오피스 COM 자동화를 사용하여 변환합니다.

### 사용법

```bash
python hwp2pdf.py input.hwp                      # 단일 파일 변환
python hwp2pdf.py input.hwp -o output.pdf         # 출력 경로 지정
python hwp2pdf.py ./docs/                         # 폴더 일괄 변환 (재귀 탐색)
python hwp2pdf.py ./docs/ -o ./pdfs/              # 출력 폴더 지정
python hwp2pdf.py input.hwp --engine libreoffice  # LibreOffice 엔진 강제 사용
```

### 옵션

| 옵션 | 설명 |
|------|------|
| `input` | 변환할 HWP/HWPX 파일 또는 폴더 경로 |
| `-o`, `--output` | 출력 PDF 파일 또는 폴더 경로 (미지정 시 원본과 같은 위치) |
| `--engine` | 변환 엔진: `auto`(기본), `hancom`, `libreoffice` |

### 변환 엔진

| 엔진 | 설명 | 비고 |
|------|------|------|
| `auto` | 한컴오피스를 먼저 시도, 실패 시 LibreOffice 폴백 | 기본값 |
| `hancom` | 한컴오피스 COM 자동화만 사용 | 최고 품질, 한컴오피스 필수 |
| `libreoffice` | LibreOffice headless 모드만 사용 | 한컴오피스 없이 사용 가능 |

### 지원 파일 형식

| 확장자 | 형식 |
|--------|------|
| `.hwp` | 한글 문서 (바이너리, v5) |
| `.hwpx` | 한글 문서 (XML 기반, OWPML) |

### 보안 모듈 설정

한컴오피스 COM 자동화 시 파일 접근 권한 팝업을 방지하려면 보안 승인 모듈을 등록해야 합니다.

1. [한컴 공식 보안 모듈](https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/%EB%B3%B4%EC%95%88%EB%AA%A8%EB%93%88(Automation).zip)에서 `FilePathCheckerModuleExample.dll`을 다운로드
2. 원하는 위치에 DLL 배치 (예: `C:\Users\<사용자>\HncSecurityModule\`)
3. 레지스트리 등록:
   - 경로: `HKCU\SOFTWARE\HNC\HwpAutomation\Modules`
   - 값 이름: `FilePathCheckerModuleExample`
   - 값 데이터: DLL의 전체 경로 (따옴표 없이)

---

## doc2pdf.py — DOC/DOCX → PDF

Microsoft Word COM 자동화를 사용하여 변환합니다.

### 사용법

```bash
python doc2pdf.py input.docx                      # 단일 파일 변환
python doc2pdf.py input.docx -o output.pdf         # 출력 경로 지정
python doc2pdf.py ./docs/                          # 폴더 일괄 변환 (재귀 탐색)
python doc2pdf.py ./docs/ -o ./pdfs/               # 출력 폴더 지정
python doc2pdf.py input.docx --engine libreoffice  # LibreOffice 엔진 강제 사용
```

### 옵션

| 옵션 | 설명 |
|------|------|
| `input` | 변환할 DOC/DOCX 파일 또는 폴더 경로 |
| `-o`, `--output` | 출력 PDF 파일 또는 폴더 경로 (미지정 시 원본과 같은 위치) |
| `--engine` | 변환 엔진: `auto`(기본), `word`, `libreoffice` |

### 변환 엔진

| 엔진 | 설명 | 비고 |
|------|------|------|
| `auto` | Word를 먼저 시도, 실패 시 LibreOffice 폴백 | 기본값 |
| `word` | Word COM 자동화만 사용 | 최고 품질, Word 필수 |
| `libreoffice` | LibreOffice headless 모드만 사용 | Word 없이 사용 가능 |

### 지원 파일 형식

| 확장자 | 형식 |
|--------|------|
| `.doc` | Word 문서 (바이너리) |
| `.docx` | Word 문서 (XML 기반, OOXML) |

> Word 임시 파일(`~$`로 시작하는 파일)은 자동으로 제외됩니다.

---

## xls2pdf.py — XLS/XLSX → PDF

Microsoft Excel COM 자동화를 사용하여 변환합니다. 모든 시트를 포함하여 하나의 PDF로 내보냅니다.

### 사용법

```bash
python xls2pdf.py input.xlsx                      # 단일 파일 변환
python xls2pdf.py input.xlsx -o output.pdf         # 출력 경로 지정
python xls2pdf.py ./docs/                          # 폴더 일괄 변환 (재귀 탐색)
python xls2pdf.py ./docs/ -o ./pdfs/               # 출력 폴더 지정
python xls2pdf.py input.xlsx --engine libreoffice  # LibreOffice 엔진 강제 사용
```

### 옵션

| 옵션 | 설명 |
|------|------|
| `input` | 변환할 XLS/XLSX 파일 또는 폴더 경로 |
| `-o`, `--output` | 출력 PDF 파일 또는 폴더 경로 (미지정 시 원본과 같은 위치) |
| `--engine` | 변환 엔진: `auto`(기본), `excel`, `libreoffice` |

### 변환 엔진

| 엔진 | 설명 | 비고 |
|------|------|------|
| `auto` | Excel을 먼저 시도, 실패 시 LibreOffice 폴백 | 기본값 |
| `excel` | Excel COM 자동화만 사용 | 최고 품질, Excel 필수 |
| `libreoffice` | LibreOffice headless 모드만 사용 | Excel 없이 사용 가능 |

### 지원 파일 형식

| 확장자 | 형식 |
|--------|------|
| `.xls` | Excel 문서 (바이너리) |
| `.xlsx` | Excel 문서 (XML 기반, OOXML) |

> Excel 임시 파일(`~$`로 시작하는 파일)은 자동으로 제외됩니다.

---

## ppt2pdf.py — PPT/PPTX → PDF

Microsoft PowerPoint COM 자동화를 사용하여 변환합니다.

### 사용법

```bash
python ppt2pdf.py input.pptx                      # 단일 파일 변환
python ppt2pdf.py input.pptx -o output.pdf         # 출력 경로 지정
python ppt2pdf.py ./docs/                          # 폴더 일괄 변환 (재귀 탐색)
python ppt2pdf.py ./docs/ -o ./pdfs/               # 출력 폴더 지정
python ppt2pdf.py input.pptx --engine libreoffice  # LibreOffice 엔진 강제 사용
```

### 옵션

| 옵션 | 설명 |
|------|------|
| `input` | 변환할 PPT/PPTX 파일 또는 폴더 경로 |
| `-o`, `--output` | 출력 PDF 파일 또는 폴더 경로 (미지정 시 원본과 같은 위치) |
| `--engine` | 변환 엔진: `auto`(기본), `powerpoint`, `libreoffice` |

### 변환 엔진

| 엔진 | 설명 | 비고 |
|------|------|------|
| `auto` | PowerPoint를 먼저 시도, 실패 시 LibreOffice 폴백 | 기본값 |
| `powerpoint` | PowerPoint COM 자동화만 사용 | 최고 품질, PowerPoint 필수 |
| `libreoffice` | LibreOffice headless 모드만 사용 | PowerPoint 없이 사용 가능 |

### 지원 파일 형식

| 확장자 | 형식 |
|--------|------|
| `.ppt` | PowerPoint 문서 (바이너리) |
| `.pptx` | PowerPoint 문서 (XML 기반, OOXML) |

> PowerPoint 임시 파일(`~$`로 시작하는 파일)은 자동으로 제외됩니다.

---

## html2pdf.py — HTML → PDF

Chrome/Edge 브라우저의 headless 모드를 사용하여 변환합니다. 추가 라이브러리 설치가 필요 없습니다.

### 사용법

```bash
python html2pdf.py input.html                      # 단일 파일 변환
python html2pdf.py input.html -o output.pdf         # 출력 경로 지정
python html2pdf.py ./docs/                          # 폴더 일괄 변환 (재귀 탐색)
python html2pdf.py ./docs/ -o ./pdfs/               # 출력 폴더 지정
python html2pdf.py input.html --engine edge         # Edge 엔진 강제 사용
```

### 옵션

| 옵션 | 설명 |
|------|------|
| `input` | 변환할 HTML 파일 또는 폴더 경로 |
| `-o`, `--output` | 출력 PDF 파일 또는 폴더 경로 (미지정 시 원본과 같은 위치) |
| `--engine` | 변환 엔진: `auto`(기본), `chrome`, `edge` |

### 변환 엔진

| 엔진 | 설명 | 비고 |
|------|------|------|
| `auto` | Chrome을 먼저 탐색, 없으면 Edge 사용 | 기본값 |
| `chrome` | Chrome headless만 사용 | |
| `edge` | Edge headless만 사용 | |

### 지원 파일 형식

| 확장자 | 형식 |
|--------|------|
| `.html` | HTML 문서 |
| `.htm` | HTML 문서 |

---

## 공통 참고 사항

- 폴더 지정 시 하위 디렉토리를 재귀적으로 탐색합니다.
- 출력 폴더가 없으면 자동 생성됩니다.
- 파일명에 공백이 포함되어도 따옴표 없이 사용할 수 있습니다.

---

## EXE 실행 파일

`dist/` 폴더에 빌드된 EXE 파일이 있습니다. Python 설치 없이 단독 실행 가능합니다.

| 파일 | 설명 |
|------|------|
| `hwp2pdf.exe` | HWP/HWPX → PDF |
| `doc2pdf.exe` | DOC/DOCX → PDF |
| `xls2pdf.exe` | XLS/XLSX → PDF |
| `ppt2pdf.exe` | PPT/PPTX → PDF |
| `html2pdf.exe` | HTML → PDF |
| `dir2pdf.exe` | 모든 문서 → PDF 통합 변환 |
| `zip2pdf.exe` | ZIP 해제 후 PDF 일괄 변환 |

사용법은 Python 스크립트와 동일합니다.

```bash
dir2pdf.exe ./docs/ -o ./pdfs/
zip2pdf.exe archive.zip -o ./pdfs/
hwp2pdf.exe input.hwp -o output.pdf
```

### EXE 빌드 방법

PyInstaller로 직접 빌드할 수 있습니다.

```bash
pip install pyinstaller

python -m PyInstaller --onefile --console hwp2pdf.py --hidden-import progress
python -m PyInstaller --onefile --console doc2pdf.py --hidden-import progress
python -m PyInstaller --onefile --console xls2pdf.py --hidden-import progress
python -m PyInstaller --onefile --console ppt2pdf.py --hidden-import progress
python -m PyInstaller --onefile --console html2pdf.py --hidden-import progress
python -m PyInstaller --onefile --console dir2pdf.py --hidden-import progress
python -m PyInstaller --onefile --console zip2pdf.py --hidden-import progress --hidden-import dir2pdf
```

빌드 결과는 `dist/` 폴더에 생성됩니다.
