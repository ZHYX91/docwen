# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)

Word/Markdown/Excel 양방향 변환을 지원하는 문서·표 변환 도구입니다. 완전 로컬 실행으로 데이터 보안과 신뢰성을 보장합니다.

## 📖 프로젝트 배경

이 소프트웨어는 문서 작업 환경에서 자주 겪는 문제를 해결하기 위해 만들어졌습니다:
- 부서별로 전달되는 문서 형식이 제각각이라 정리/표준화가 필요함
- 문서 유형이 다양하고 유형별로 요구되는 고정 포맷이 다름
- 내부망/구형 PC에서도 동작해야 하므로 오프라인 실행이 필요함

**설계 철학**: 전문 툴(LaTeX, Pandoc 등)만큼의 범용성과 완성도를 목표로 하기보다는, 학습 비용이 거의 없는 “간단하고 바로 쓰는” 변환 도구에 초점을 맞춥니다.

## ✨ 핵심 기능

- **📄 문서 변환** - Word ↔ Markdown 양방향 변환. 수식 변환, 구분선(---/***/___)과 페이지/구역/가로선 매핑을 지원하며, Markdown 표의 명시적 `<` / `^` marker 를 Word의 직사각형 셀 병합으로 복원할 수 있습니다. DOCX/DOC/WPS/RTF/ODT 지원.
- **📊 스프레드시트 변환** - Excel ↔ Markdown 양방향 변환. XLSX/XLS/ET/ODS/CSV/TSV, 병합 셀 내보내기 전략(`fill / empty / marker`), 표 요약 도구와 아래에 설명된 템플릿 플레이스홀더를 지원합니다.
- **📑 PDF/레이아웃 파일** - PDF/XPS/OFD → Markdown 또는 DOCX. PDF 병합/분할 등 지원.
- **🖼️ 이미지 처리** - JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC 변환 및 압축.
- **📥 기타 형식 가져오기** - HTML/MHTML/ENEX/EPUB/PPTX/PPT에서 Markdown으로의 단방향 변환을 지원합니다.
- **🔍 OCR 텍스트 인식** - RapidOCR을 통합하여 이미지와 PDF에서 텍스트를 추출합니다.
- **✏️ 교정(Proofread)** - Word (.docx) 및 Markdown (.md) 파일의 사용자 사전 기반 오탈자/문장부호/기호/민감어 검사. 규칙은 설정 화면에서 편집할 수 있습니다.
- **📝 템플릿 시스템** - 문서/보고서 형식을 템플릿으로 관리.
- **💻 GUI + CLI** - 그래픽 UI와 명령행 모두 제공.
- **🔒 종속성 송신 보호를 갖춘 로컬 처리** - 변환은 온라인 서비스에 의존하지 않습니다. DocWen 실행 중 Python 프로세스 내부 종속성의 DNS 및 IPv4/IPv6 사용을 차단하며, 외부 Office 프로그램은 자체 시스템 네트워크 정책을 따릅니다.
- **🔗 단일 인스턴스 실행** - 프로그램 인스턴스를 자동으로 관리하며, 동반 Obsidian 플러그인과의 연동을 지원합니다.

## 📸 스크린샷

| 일괄 | Markdown |
| --- | --- |
| ![일괄 패널](../assets/screenshots/batch-light.png) | ![메인 창](../assets/screenshots/main-light.png) |

| 문서 | 스프레드시트 |
| --- | --- |
| ![문서 패널](../assets/screenshots/conversion-document-light.png) | ![스프레드시트 패널](../assets/screenshots/conversion-spreadsheet-light.png) |

| 이미지 | 레이아웃 파일 |
| --- | --- |
| ![이미지 패널](../assets/screenshots/conversion-image-light.png) | ![레이아웃 패널](../assets/screenshots/conversion-layout-light.png) |

변경 이력: [CHANGELOG.md](../CHANGELOG.md) 참고

## 🚀 빠른 시작

### 소스에서 설치

**전제 조건**: Python 3.12

**0.9 대상 범위**: 이 소스는 Windows x64와 Ubuntu 24.04 x64 패키지를 빌드합니다. 다른
Linux 배포판과 macOS는 소스/개발 경로이며 Ubuntu 패키지의 지원 범위에 포함되지 않습니다.

**방법 1: uv 사용 (권장)**

[uv](https://docs.astral.sh/uv/getting-started/)를 설치한 후:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

DocWen 0.9 소스, 테스트 및 빌드는 저장소 잠금 파일과 `uv 0.12.0`만 지원합니다. `pip install -e`는 지원되지 않습니다.

### 프로그램 실행

Windows 패키지 버전에서는 `DocWen.exe`를 더블클릭하여 GUI를 실행합니다. 소스에서 설치한 경우:

```bash
docwen-gui  # GUI 모드
docwen      # CLI 모드
```

### macOS 설치 안내

**현재 제한**: macOS에서는 `convert`, `validate`, `number`, `merge`, `split` capability를 현재
사용할 수 없습니다. 아래 내용은 개발 실험을 위한 선택적 의존성만 설명합니다.

**LibreOffice 지원(선택)**

`.doc`, `.xls` 같은 구형 포맷 변환이 필요하다면 LibreOffice를 설치하세요:  
다운로드: https://www.libreoffice.org/download/

**HEIC 이미지 지원(선택)**

HEIC/HEIF 이미지를 처리하려면:

```bash
brew install libheif
pip install pillow-heif
```

### Linux GUI 버전 사전 준비

**지원 패키지 대상**: DocWen 0.9는 Ubuntu 24.04 x64 패키지의 GUI와 CLI를 지원합니다.
이 요구 사항은 다른 배포판이나 아키텍처로 지원 범위를 확장하지 않습니다.

- 데스크톱 환경이 설치되어 있어야 합니다(GNOME, KDE, XFCE 등)
- GUI는 PySide6(Qt6) 기반이며 더 이상 Python Tk에 의존하지 않습니다. 시작 시 시스템 라이브러리가 부족하다는 오류가 나면, 오류 메시지에 표시된 Qt 런타임 의존성(보통 OpenGL/X11 관련)을 설치하세요.
- 헤드리스 서버에서는 GUI 대신 CLI 진입점 `docwen` 사용을 우선하세요. Windows 패키지 빌드에는 `DocWenCLI.exe`도 포함됩니다.

### 빠른 시작 가이드

1.  **Markdown 파일 준비**:

    ```markdown
    ---
    title: Test Document
    ---
    
    ## Test Title
    
    This is the test body content.
    ```

2.  **드래그 앤 드롭 변환**:
    - 프로그램을 실행합니다.
    - `.md` 파일을 창으로 드래그합니다.
    - 템플릿을 선택합니다.
    - "DOCX로 변환"을 클릭합니다.

3.  **결과 확인**:
    - 동일한 디렉터리에 표준화된 Word 문서가 생성됩니다.

**팁**: `samples/` 디렉터리의 샘플 파일을 사용하면 기능을 빠르게 체험할 수 있습니다.

## 🖥️ 그래픽 인터페이스 사용

대부분의 사용자는 그래픽 인터페이스로 이 소프트웨어를 사용합니다. 아래는 상세한 사용 가이드입니다.

### 인터페이스 개요

프로그램은 **적응형 3열 레이아웃**을 사용합니다:

| 영역 | 설명 | 표시 시점 |
| :--- | :--- | :--- |
| **중앙 열(메인 영역)** | 파일 드래그 앤 드롭 영역, 작업 패널, 상태 표시줄 | 항상 표시 |
| **오른쪽 열** | 템플릿 선택기 / 포맷 변환 패널 | 파일 선택 후 자동 확장 |
| **왼쪽 열** | 배치 파일 목록(형식별 그룹) | 배치 모드에서 표시 |

### 기본 작업 흐름

1.  **프로그램 실행**: `DocWen.exe`(Windows 패키지) 더블클릭 또는 `docwen-gui` 실행.
2.  **파일 가져오기**:
    -   방법 1: 파일을 창으로 드래그합니다.
    -   방법 2: 드래그 영역의 "추가" 버튼을 눌러 파일을 선택합니다.
3.  **템플릿 선택**(변환이 필요한 경우): 오른쪽 템플릿 패널이 자동으로 확장되며, 적절한 템플릿을 선택합니다.
4.  **옵션 설정**: 작업 패널에서 필요한 변환/내보내기 옵션을 체크합니다.
5.  **작업 실행**: 해당 기능 버튼(예: "MD 내보내기", "DOCX로 변환")을 클릭합니다.
6.  **결과 확인**: 상태 표시줄에 진행 상황과 결과가 표시되며, 오른쪽의 "출력 열기" 동작을 클릭하면 출력 위치를 열 수 있습니다.

### 단일 파일 모드 vs 배치 모드

프로그램은 2가지 처리 모드를 지원하며, 드래그 앤 드롭 영역의 토글 버튼으로 전환할 수 있습니다:

**단일 파일 모드**(기본값):
-   한 번에 한 파일 처리
-   인터페이스가 단순해 일상 사용에 적합

**배치 모드**:
-   여러 파일을 동시에 가져오기
-   왼쪽 열에 파일 목록을 형식별로 그룹화해 표시
-   일괄 추가/제거/정렬 지원
-   목록에서 파일을 클릭해 현재 작업 대상을 전환

### 작업 패널 기능

작업 패널은 파일 형식에 따라 사용 가능한 기능을 자동으로 조정합니다:

| 파일 형식 | 사용 가능한 작업 |
| :--- | :--- |
| Word 문서 | MD 내보내기, PDF 변환, 텍스트 교정, OCR |
| Markdown | DOCX 변환, PDF 변환, 텍스트 교정 |
| Excel 스프레드시트 | MD 내보내기, PDF 변환, 표 요약 |
| PDF | MD 내보내기, 병합, 분할, OCR |
| 이미지 | 포맷 변환, 압축, OCR |
| HTML/EPUB/PPTX 등 | MD 내보내기 |

### 설정 화면

작업 영역 헤더의 "설정" 버튼을 눌러 설정을 엽니다:

설정은 탭으로 구성됩니다: **일반**, **텍스트**, **교정**, **문서**, **스프레드시트**, **이미지**, **레이아웃**, **링크**, **서식**, **출력**, **내보내기**, **로그**, **기타**.

### 단축 동작

-   **외부 파일 드래그**: 창으로 직접 드래그해 가져오기
-   **출력 위치 열기**: 상태 표시줄 오른쪽의 "출력 열기" 동작을 클릭해 출력 위치를 엽니다.
-   **템플릿 항목 우클릭**: 템플릿 파일 위치 열기
---

## 🔧 명령줄 사용

DocWen은 GUI 외에도 자동화 스크립트, 배치 처리, 외부 연동을 위한 명령줄 인터페이스(CLI)를 제공합니다.

### 자동화 권장 호출 순서

스크립트, Agent, 플러그인 연동에서는 다음 순서를 권장합니다.

1. `inspect <file> [--json]`: 먼저 실제 파일 범주, 형식, 지원 동작을 확인합니다.
2. `schema convert`: `convert` 의 기계 판독 가능한 계약과 조건 규칙을 읽습니다.
3. `convert <file> --to <fmt> --output <path> --dry-run --json`: 결과를 쓰지 않고 탐지, 정규화, 라우팅을 미리 확인합니다.
4. `convert <file> --to <fmt> --output <path> ...`: 확인 후 실제 변환을 실행합니다.

### 자주 쓰는 예시

```bash
# Windows 패키지
DocWenCLI.exe inspect document.docx --json

# 스크립트 / Agent 용 convert 계약 내보내기
DocWenCLI.exe schema convert

# 실제 파일 생성 없이 변환 경로 미리 보기
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr --dry-run --json

# Word 를 Markdown 으로 내보내기 (이미지 추출 + OCR)
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr

# Markdown 을 Word 로 변환 (템플릿 + 제목/본문 병합 모드)
DocWenCLI.exe convert document.md --to docx --output document.docx --template template.docx.6cd486f34e59c79ded078a008b269af37860b63ccb74d8d0ab0080a7229a9ab5 --heading-merge-mode punct_required

# Markdown 출력 시 이미지 모드와 OCR 텍스트 배치 제어
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --image-mode file --ocr --ocr-placement image_md

# 런타임 능력과 의존성 게이트 확인
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# 문서 교정
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# 소스 / uv 설치
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output document.md --dry-run --json
# docwen convert document.docx --to md --output document.md
```

### 자주 쓰는 명령과 옵션

아래 표는 자주 쓰는 명령만 정리합니다. 전체 명령 집합은 `docwen --help`(소스 / uv) 또는 `DocWenCLI --help`(패키지판)를 참고하세요.

| 명령 / 옵션 | 설명 |
| --- | --- |
| `convert <file> --to <fmt> --output <path>` | 변환의 통합 진입점입니다. |
| `convert <file> --to <fmt> --output <path> --dry-run --json` | 실제 변환 없이 탐지, 정규화, 라우팅, 적용 옵션만 미리 확인합니다. |
| `schema convert` | `convert` 의 기계 판독 가능한 계약, 기본값, 조건, 정규 키를 내보냅니다. |
| `validate <file> --check ...` | 문서 교정(`typo/punct/symbol/sensitive/all/none`). CLI envelope에는 `--json`을 사용합니다. `--report`는 선택적 보고서 파일 경로입니다. |
| `inspect <file> [--json]` | 파일 범주/형식, 권장 동작, 확장자와 내용 불일치 경고를 확인합니다. |
| `doctor --json` | 진단 결과와 함께 런타임 능력 요약 및 의존성 게이트를 출력합니다. |
| `resources list formats --json` | 원본 범주별 대상 형식과 의존성 게이트 / 제한 요약을 출력합니다. |
| `resources list templates` | 사용 가능한 템플릿을 나열합니다. |
| `resources list numbering-schemes` | 사용 가능한 번호 체계를 나열합니다. |
| `--template <id>` | `resources list templates`가 반환한 정규 리소스 ID를 그대로 사용합니다. 표시 이름·파일명·경로는 거부됩니다. DOCX ID는 `docx/doc/odt/rtf/wps/pdf`, XLSX ID는 `xlsx/xls/ods/csv`에 적용됩니다. |
| `--extract-img` / `--no-extract-img` / `--ocr` | `convert --to md` 용 이미지 추출 및 OCR 옵션입니다. |
| `--image-mode file|base64` | Markdown 내보내기 시 이미지 출력 방식을 제어합니다. |
| `--ocr-placement image_md|main_md` | OCR 텍스트를 이미지용 Markdown 에 쓸지 메인 Markdown 에 쓸지 제어합니다. |
| `--heading-merge-mode punct_required|always|never` | `convert --to docx` 시 "제목 + 본문" 병합 전략을 제어합니다. |
| `--optimization <id>` | 최적화 프로필을 명시적으로 활성화합니다 (`resources list optimizations` 참고). |
| `batch convert|validate ... --jobs <n> [--continue-on-error]` | 배치 처리 제어 옵션입니다. |
| `--json` / `--quiet` / `--timing` | 스크립트나 플러그인을 위한 구조화 출력, 로그 축소, 시간 정보입니다。 |

`punct_required` 모드의 정확한 기본값은 `。：！？.:!?`입니다. 서식 설정에서 수정할 수 있으며, 값을 비우면 이 모드에서 병합하지 않습니다. 쉼표, 세미콜론, 열거 쉼표, 대시 및 줄임표는 기본값에 포함되지 않습니다.


## 📝 Markdown 문법 규칙

### 제목 수준 매핑

Markdown 제목은 Word 제목과 **1:1**로 대응됩니다:
- 문서 제목과 부제목은 YAML 메타데이터에 둡니다.
- Markdown `# Heading 1` → Word "Heading 1"
- Markdown `## Heading 2` → Word "Heading 2"
- 이와 같은 방식으로 최대 9레벨까지 지원합니다.

**팁**: Markdown의 1레벨 제목(`#`)을 문서 제목으로 쓰고, 2레벨 제목(`##`)부터 본문 제목으로 쓰고 싶다면 Word 템플릿에서 "Heading 1" 스타일을 제목처럼 보이도록(예: 가운데 정렬, 굵게, 큰 글꼴) 조정한 뒤, 설정에서 1레벨 제목 번호를 건너뛰는 번호 매기기 스킴을 선택하면 됩니다.

### 줄바꿈과 단락

**기본 규칙**: 비어 있지 않은 각 줄은 기본적으로 독립된 단락으로 처리됩니다.

**혼합 단락**: 소제목을 본문 텍스트와 같은 단락으로 섞어야 하는 경우(기본 모드: "문장부호 필요"), 아래 조건을 만족해야 합니다:
1.  소제목이 종결 문장부호로 끝납니다(마침표/물음표/느낌표 등 다국어 종결 문장부호 지원).
2.  본문 텍스트가 소제목의 **바로 다음 줄**에 위치합니다.
3.  본문 줄은 특수 Markdown 요소(제목, 코드 블록, 표, 목록, 인용, 수식 블록, 구분선 등)가 아니어야 합니다.

**예시**:
```markdown
## I. Work Requirements.
This meeting requires all units to earnestly implement...
```
위 두 줄은 하나의 단락으로 병합되며, "I. Work Requirements."는 소제목 형식을 유지하고 "This meeting..."은 본문 형식을 유지합니다.

**주의**:
- 소제목과 본문 사이에 빈 줄이 있으면 별도 단락으로 인식됩니다.
- 기본(“문장부호 필요” 모드)에서는 소제목이 종결 문장부호로 끝나지 않으면 빈 줄이 없어도 다음 줄과 병합되지 않습니다.
- 이 동작은 설정 → 서식 → “MarkDown→문서” → “Heading + body merge mode”에서 변경할 수 있습니다.

### 구분선 양방향 변환

Markdown 구분선과 Word의 페이지/구역/가로선 요소 간의 양방향 변환을 지원합니다:

-   **DOCX → MD**: Word의 페이지 나눔, 구역 나눔, 가로선을 Markdown 구분선으로 자동 변환합니다.
-   **MD → DOCX**: Markdown `---`, `***`, `___`를 대응되는 Word 요소로 자동 변환합니다.
-   **구성 가능**: 구체적인 매핑 관계는 설정 화면에서 사용자 지정할 수 있습니다.

### 작업 목록

GFM 작업 목록의 양방향 변환을 지원합니다:

```markdown
- [ ] 할 일
- [x] 완료
```

-   **MD → DOCX**: `☐` / `☑` 텍스트 접두사가 포함된 글머리 기호 목록으로 렌더링됩니다.
-   **DOCX → MD**: `☐` / `☑` / `☒`로 시작하는 목록 항목을 `- [ ]` / `- [x]`로 변환합니다.
-   **글꼴 참고**: `☐`/`☑`는 일부 글꼴에서 표시되지 않을 수 있습니다. 필요한 경우 Word 템플릿에서 "Segoe UI Symbol" 등의 글꼴을 사용하세요.

### 이미지 임베드 및 크기

Obsidian/Wiki 및 표준 Markdown 이미지 임베드를 지원하며, 크기 지정(px)도 가능합니다:

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- 크기 미지정: 원본 크기 사용(페이지/셀 사용 가능 너비를 상한으로 제한)
- 크기 지정: 확대를 허용하되, 사용 가능 너비 상한은 적용
- 이미지 전용 단락: 단락 스타일 “Image”(가운데 정렬, 단일 줄 간격) 사용

### 링크 처리

Markdown -> DOCX 변환 시 클릭 가능한 링크를 지원합니다:

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Markdown 링크와 Wiki 링크는 기본적으로 Word 하이퍼링크로 출력됩니다
- Wiki 링크는 대상 파일을 찾으면 로컬 `file:///` 링크로 해석됩니다
- 꺾쇠 괄호 자동 링크는 `https://...` 와 이메일 `mailto:...` 를 지원합니다
- 일반 URL 자동 링크는 Markdown -> DOCX 요청별로 평가되며 기본값은 꺼짐입니다. `configs/link.toml`의 `[non_embed_links].auto_link_bare_url`로 활성화할 수 있습니다
- Markdown -> XLSX 에서는 DOCX 하이퍼링크 플레이스홀더를 만들지 않고 원래 링크 문법을 유지합니다

## 📖 상세 사용 가이드

### Word → Markdown

1.  `.docx` 파일을 프로그램 창으로 드래그합니다.
2.  프로그램이 문서 구조를 자동 분석합니다.
3.  YAML 메타데이터가 포함된 `.md` 파일을 생성합니다.

**지원 형식**:
-   `.docx` - 표준 Word 문서
-   `.doc` - 자동으로 DOCX로 변환 후 처리
-   `.wps` - WPS 문서를 자동 변환 후 처리

**내보내기 옵션**:

| 옵션 | 설명 |
| :--- | :--- |
| **이미지 추출** | 체크 시 문서의 이미지를 출력 폴더로 추출하고, MD에 이미지 링크를 삽입합니다. |
| **이미지 OCR** | 체크 시 이미지에 OCR을 수행하고, 인식 텍스트를 담은 이미지 `.md` 파일을 생성합니다. |
| **고급 필드 최적화** | 체크 시 더 풍부한 구조화 메타데이터를 추출합니다. 체크하지 않으면 제목과 부제목만 포함하는 간소화 모드를 사용합니다. |
| **소제목 번호 제거** | 체크 시 소제목 앞 번호(예: "一、", "（一）", "1." 등)를 제거하고 순수 제목 텍스트로 변환합니다. |
| **소제목 번호 추가** | 체크 시 제목 레벨에 따라 자동 번호를 추가합니다(번호 스킴은 설정에서 구성 가능). |

참고: DOCX -> MD는 이제 numbering.xml 에서 단락 스타일(pStyle)로 연결된 다단계 번호도 복원합니다. 따라서 Word/WPS의 다단계 목록으로 만든 제목 접두사(예: "一、", "（一）", "1．", "（1）", "①")가 단순 모드와 고급 필드 모드 모두에서 유지되며, "소제목 번호 제거"를 켜도 제목 수준을 올바르게 인식합니다.

### Markdown → Word

1.  YAML 헤더가 포함된 `.md` 파일을 준비합니다.
2.  프로그램 창으로 드래그하고 해당 Word 템플릿을 선택합니다.
3.  프로그램이 템플릿을 자동으로 채우고 문서를 생성합니다.

**변환 옵션**:

| 옵션 | 설명 |
| :--- | :--- |
| **소제목 번호 제거** | 체크 시 소제목 앞 번호를 제거합니다. |
| **소제목 번호 추가** | 체크 시 제목 레벨에 따라 자동 번호를 추가합니다. |

**주의**: 소제목과 본문이 같은 단락으로 섞이는 문장이 있다면, MD 파일에서 줄바꿈 규칙을 엄격히 유지해야 합니다(위 "줄바꿈과 단락" 참고).

### 템플릿 스타일 자동 처리

Markdown → DOCX 변환 시 변환기는 템플릿 스타일을 자동으로 감지하고 처리합니다:

#### 스타일 분류

**단락 스타일(Paragraph Style)**: 단락 전체에 적용됩니다.

| 스타일 | 감지 동작 | 없을 때 주입 | 출처 |
| :--- | :--- | :--- | :--- |
| Heading (1~9) | 단락 스타일 감지 | 템플릿의 Heading 스타일 | Word 기본 |
| Code Block | 단락 스타일 감지 | Consolas 글꼴 + 회색 배경 | 소프트웨어 정의 |
| Quote (1~9) | 단락 스타일 감지 | 회색 배경 + 왼쪽 테두리 | 소프트웨어 정의 |
| Formula Block | 단락 스타일 감지 | 수식 전용 스타일 | 소프트웨어 정의 |
| Separator (1~3) | 단락 스타일 감지 | 아래쪽 테두리 단락 스타일 | 소프트웨어 정의 |

**문자 스타일(Character Style)**: 선택된 텍스트에 적용됩니다.

| 스타일 | 감지 동작 | 없을 때 주입 | 출처 |
| :--- | :--- | :--- | :--- |
| Inline Code | 문자 스타일 감지 | Consolas 글꼴 + 회색 음영 | 소프트웨어 정의 |
| Inline Formula | 문자 스타일 감지 | 수식 전용 스타일 | 소프트웨어 정의 |

**표 스타일(Table Style)**: 표 전체에 적용됩니다.

| 스타일 | 감지 동작 | 없을 때 주입 | 출처 |
| :--- | :--- | :--- | :--- |
| Three-Line Table | 사용자 설정 우선 | 3선 표 스타일 정의 | 소프트웨어 정의 |
| Grid Table | 사용자 설정 우선 | 격자 표 스타일 정의 | 소프트웨어 정의 |

**번호 정의(Numbering Definition)**: 목록 포맷에 사용됩니다.

| 유형 | 감지 동작 | 없을 때 처리 |
| :--- | :--- | :--- |
| List Numbering | 템플릿의 기존 목록 정의 스캔 | decimal/bullet 프리셋 사용 |

#### 스타일 이름 국제화

-   **Word 기본 스타일**(heading 1~9):
    -   스타일 이름은 Word 표준 영문 이름(예: `heading 1`)을 사용합니다.
    -   Word는 시스템 언어에 따라 로컬라이즈된 표시 이름을 자동으로 제공합니다.
-   **소프트웨어 정의 스타일**(Code Block, Quote, Formula, Separator, Table 등):
    -   인터페이스 언어 설정에 맞춰 해당 언어의 스타일 이름을 주입합니다.
    -   예: 한국어 UI에서는 "코드 블록", "인용 1" 등의 이름을 주입할 수 있습니다.
    -   영어 UI에서는 "Code Block", "Quote 1", "Three Line Table" 등으로 주입합니다.
**권장 사항**: 템플릿에서 스타일을 커스터마이즈하면 변환기는 이를 우선 사용하며, 템플릿에 없으면 내장 프리셋을 사용합니다.

### 스프레드시트 파일 처리

1.  **Excel/CSV → Markdown**: `.xlsx` 또는 `.csv` 파일을 드래그하면 Markdown 표로 자동 변환합니다.
2.  **Markdown → Excel**: Markdown 표를 XLSX로 내보낼 수 있습니다. 템플릿은 YAML 필드, 표 열·이미지 플레이스홀더와 병합/보호 셀을 지원합니다.

**지원 형식**:
-   `.xlsx` - 표준 Excel 문서
-   `.xls` - 자동으로 XLSX로 변환 후 처리
-   `.et` - WPS 스프레드시트를 자동 변환 후 처리
-   `.csv` - CSV 텍스트 표
-   `.tsv` - TSV 탭 구분 표


### 텍스트 교정 기능

프로그램은 4가지 사용자 정의 교정 규칙을 제공합니다:

1.  **문장부호 짝 검사** - 괄호/따옴표 등 짝이 맞는지 검사합니다.
2.  **기호 교정** - 한/영 문장부호 혼용 등을 검사합니다.
3.  **오탈자 검사** - 사용자 사전 기반으로 오탈자를 검사합니다.
4.  **민감어 탐지** - 사용자 사전 기반으로 민감어를 탐지합니다.

**사용자 사전**: 설정 화면에서 오탈자/민감어 사전을 시각적으로 편집할 수 있습니다.

**사용 방법**:
1.  교정할 Word 문서 또는 Markdown 파일을 프로그램에 드래그합니다.
2.  필요한 교정 규칙을 체크합니다.
3.  "텍스트 교정" 버튼을 클릭합니다.
4.  결과는 문서의 주석(코멘트)로 표시됩니다. Markdown 파일의 경우 JSON 보고서로 출력됩니다.

추가 설명(Markdown 교정 JSON 보고서):
- 엔진: `text_rules` + Markdown 어댑터 `md_spell`
- 출력 방식: 현재 CLI 교정 경로는 `validate`입니다. CLI envelope에는 `--json`을 사용하세요. `--report`는 선택적 보고서 파일 경로입니다.

## 🛠️ 템플릿 시스템

### 기존 템플릿 사용

프로그램에는 여러 템플릿(다국어 포함)이 기본 제공됩니다. 필요에 따라 선택해 사용할 수 있으며, 템플릿 파일은 `templates/` 디렉터리에 있습니다.

### 사용자 정의 템플릿

1.  Word 또는 WPS로 템플릿 파일을 만듭니다.
2.  기존 템플릿을 참고해 `{{Title}}`, `{{DocumentNumber}}` 등 필요한 위치에 플레이스홀더를 삽입합니다.
3.  템플릿에서 Word 기본 Heading 1 ~ Heading 5 스타일은 수동으로 수정해야 합니다.
4.  템플릿을 `templates/` 디렉터리에 저장합니다.
5.  프로그램을 재시작하면 새 템플릿이 자동으로 로드됩니다.

기존 템플릿을 복사해 수정한 뒤 이름을 변경해 사용하는 방식도 가능합니다.

### 플레이스홀더 사용법

#### Word 템플릿 플레이스홀더

**YAML 필드 플레이스홀더**: 템플릿에서 `{{Field Name}}` 형식을 사용하면, 변환 시 Markdown 파일의 YAML 헤더에서 해당 필드 값으로 치환됩니다.

| 플레이스홀더 | 설명 |
| :--- | :--- |
| `{{Title}}` | 문서 제목(아래 우선순위 참고) |
| `{{Body}}` | Markdown 본문 삽입 위치 |
| 기타 | 임의의 사용자 정의 필드 지원 |

**제목(Title) 가져오기 우선순위**:

| 우선순위 | 출처 | 설명 |
| :--- | :--- | :--- |
| 1 | YAML `Title` | 최우선 |
| 2 | YAML `aliases` | 리스트의 첫 원소 또는 문자열 값 |
| 3 | 파일명 | `.md` 확장자를 제외한 파일명 |

**다국어 지원**: 제목/본문 플레이스홀더는 다국어를 지원합니다. 예: 제목 `{{title}}`, `{{标题}}`, `{{Titel}}` 등, 본문 `{{body}}`, `{{正文}}`, `{{Inhalt}}` 등.

#### Excel 템플릿 플레이스홀더

XLSX 템플릿은 YAML 필드, 세로 `{{↓필드}}`와 가로 `{{→필드}}`, 이미지 플레이스홀더와 병합/보호 셀을 지원합니다.

**1. YAML 필드 플레이스홀더** `{{Field Name}}`

Markdown 파일의 YAML 헤더에 있는 단일 값을 채웁니다:

```markdown
---
ReportName: 2024 Annual Sales Statistics
Unit: Sales Dept
---
```

템플릿의 `{{ReportName}}`, `{{Unit}}`가 해당 값으로 치환됩니다. Title 필드도 동일한 우선순위를 따릅니다.

**2. 열 채우기 플레이스홀더** `{{↓Field Name}}`

Markdown 표에서 데이터를 추출해 플레이스홀더 위치부터 **아래로** 행 단위로 채웁니다:

```markdown
| ProductName | Quantity |
|:--- |:--- |
| Product A | 100 |
| Product B | 200 |
```

`{{↓ProductName}}`는 "Product A"로 치환되며, 다음 행은 "Product B"가 채워집니다.

**3. 행 채우기 플레이스홀더** `{{→Field Name}}`

Markdown 표에서 데이터를 추출해 플레이스홀더 위치부터 **오른쪽으로** 열 단위로 채웁니다:

```markdown
| Month |
|:--- |
| Jan |
| Feb |
| Mar |
```

`{{→Month}}`는 오른쪽으로 "Jan", "Feb", "Mar" 순서로 채워집니다.

**병합 셀 처리**:

- Markdown -> Excel 은 템플릿에 원래 있던 merged ranges 를 계속 유지합니다.
- 연속된 `{{↓Field Name}}` 플레이스홀더로 구성된 알려진 세로형 템플릿 영역에서는 Markdown 표의 명시적 `<` / `^` marker 로부터 직사각형 병합을 복원할 수 있습니다.
- 앞뒤 공백을 제거한 내용이 정확히 `<` 또는 `^` 인 셀만 병합 판정에 참여하며, `\<` 와 `\^` 는 리터럴 텍스트로 유지됩니다.
- 잘못된 직사각형이거나 템플릿의 기존 merged ranges 와 충돌하면 템플릿 구조를 강제로 덮어쓰지 않고 경고를 남긴 뒤 일반 텍스트로 강등합니다.

**다중 표 데이터 병합**: 동일한 헤더 이름을 가진 표가 Markdown에 여러 개 있으면, 데이터를 순서대로 병합해 연속으로 채웁니다.

## 🔌 Obsidian 플러그인

변환기와 함께 동작하는 동반 Obsidian 플러그인은 별도 저장소로 제공됩니다:

### 핵심 기능

-   **🚀 원클릭 실행** - 사이드바 아이콘으로 변환기 빠르게 실행
-   **📂 자동 전달** - 현재 열려 있는 파일 경로를 자동으로 전달
-   **🔄 단일 인스턴스 관리** - 이미 실행 중이면 파일만 전송하고 재시작 불필요
-   **🔒 범위가 제한된 로컬 제어** - 프로세스 이름 탐색이나 명령/상태 파일 없이 형식화된 `status`, `open`, `activate` 요청을 사용

### 동작 원리

DocWen Core의 runtime/control transport는 Windows 명명된 파이프 또는 Linux/macOS의 AF_UNIX
소켓을 사용할 수 있습니다. 파일 잠금은 단일 인스턴스 소유권만 설정하며 제어 명령 전송에는
파일을 사용하지 않습니다. 이는 Core 기능 설명일 뿐입니다. DocWen Assistant 2.0은 Windows
데스크톱 전용이며 Linux/macOS 조합 검수는 없습니다.

1.  **첫 클릭** → 변환기를 실행하고 현재 파일을 전달
2.  **다시 클릭(파일 있음)** → 새 파일로 교체(단일 파일 모드)
3.  **다시 클릭(파일 없음)** → 변환기 창을 활성화

### 설치

DocWen Assistant 2.0은 DocWen Machine Protocol v1과 단일 Artifact Bundle v2 계약을 사용합니다. 소스
버전만으로 게시 여부를 증명할 수 없습니다. 호환되는 게시된 DocWen 릴리스를 명시한 숫자 형식의 릴리스만
설치하세요.

## 🔌 OpenClaw (Plugin + Skill)

OpenClaw 2.0은 DocWen Machine Protocol v1과 단일 Artifact Bundle v2 계약을 사용합니다. 소스 버전만으로
게시 여부를 증명할 수 없습니다. 숫자 형식의 릴리스 페이지를 확인하고 변경 불가능한 릴리스 게이트가 성공한
후에만 설치하세요.

## ❓ 자주 묻는 질문 (FAQ)

### 변환이 실패하면 어떻게 하나요?

-   파일이 다른 프로그램에서 사용 중인지 확인합니다.
-   파일 형식이 올바른지 확인합니다.
-   설정의 "현재 실제 로그 파일 경로"를 확인하거나 시스템 사용자 로그 디렉터리에서 오류 로그를 확인하세요. 패키지 검증에서 `DOCWEN_LOG_DIR`를 사용했다면 해당 덮어쓴 디렉터리를 확인하세요.

### 템플릿이 표시되지 않나요?

-   템플릿 파일이 `templates/` 디렉터리에 있는지 확인합니다.
-   템플릿 파일이 손상되었는지 확인합니다.
-   프로그램을 재시작해 템플릿을 다시 로드합니다.

### 교정 기능이 동작하지 않나요?

-   문서가 `.docx` 또는 `.md` 형식인지 확인합니다.
-   문서에 편집 가능한 텍스트가 있는지 확인합니다.
-   설정에서 교정 규칙이 활성화되어 있는지 확인합니다.

### 출력 형식이 기대와 다르나요?

-   프로그램은 템플릿 스타일을 기반으로 문서를 생성합니다. 출력 형식을 조정하려면 템플릿 파일의 스타일 정의를 직접 수정하세요.
-   템플릿 파일은 `templates/` 디렉터리에 있습니다.
-   템플릿 스타일을 수정하면, 해당 템플릿으로 변환된 모든 문서에 적용됩니다.

### Excel → Markdown 변환 후 수식 셀이 비어 있나요?

이는 정상 동작입니다. 프로그램은 수식 자체가 아니라 셀의 **캐시된 값**을 읽습니다.

**기술적 이유**:
- Excel 파일에서 수식 셀은 수식과 마지막 계산 결과(캐시 값)를 함께 저장합니다.
- 프로그램은 `data_only=True` 모드를 사용해 캐시된 값만 읽습니다.
- 파일이 Excel에서 열려 계산/저장된 적이 없으면 캐시 값이 비어 있을 수 있습니다.

**해결 방법**:
1. Excel에서 파일을 엽니다.
2. 수식 계산이 완료될 때까지 기다립니다.
3. 파일을 저장합니다.
4. 다시 변환합니다.

## 🔒 보안 기능

-   **완전한 로컬 실행**: 처리는 기본적으로 로컬에서 수행되며 온라인 서비스에 의존하지 않습니다.
-   **종속성 송신 보호**: 지원되는 GUI/CLI 진입점은 Python 주 프로세스 전체 수명 동안 CPython 감사 가드를 활성화합니다. 모든 DNS/이름 확인과 AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto`, `sendmsg` 작업을 차단하면서 Windows 명명된 파이프와 Unix 도메인 소켓은 유지합니다.
-   **명확한 경계**: Office/WPS/LibreOffice 및 전용 Office 헬퍼를 포함해 별도로 실행된 프로세스는 관리하지 않습니다. 이는 종속성의 우발적 통신을 막는 심층 방어이며 운영체제 샌드박스가 아닙니다.
-   **데이터 업로드 없음**: 기본적으로 사용자 파일을 외부 서버로 적극 업로드하지 않습니다.
-   **엄격 보안 모드**: 기본 활성화되며, 핵심 보안 검사에 실패하면 프로그램이 종료됩니다. 자세한 내용은 [Troubleshooting](../maintenance/troubleshooting.md) 참고.

## 📜 라이선스

이 프로젝트는 **GNU Affero General Public License v3.0 (AGPL-3.0)** 라이선스를 따릅니다.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   이 프로젝트는 PyMuPDF(AGPL-3.0)를 사용하므로, 전체 프로젝트도 AGPL-3.0을 따릅니다.
- 현재 GUI 는 지원되는 호스트 경로에서 `PySide6-Fluent-Widgets`(QFluentWidgets)를 사용할 수 있습니다. 이 의존성은 `GPLv3 / 상용 라이선스` 이중 라이선스를 사용하며, 이 저장소는 계속 AGPL로 배포됩니다.
-   소프트웨어를 자유롭게 사용/수정/배포할 수 있습니다.
-   소프트웨어를 수정해 네트워크를 통해 서비스를 제공하는 경우, 수정된 소스 코드를 사용자에게 제공해야 합니다.
-   자세한 라이선스 정보는 [LICENSE](../../LICENSE)를 참고하세요.
- 서드파티 구성요소 고지는 [LICENSE_THIRD_PARTY.txt](../../LICENSE_THIRD_PARTY.txt), 배포 요약은 [NOTICE.txt](../../NOTICE.txt)에서 확인할 수 있습니다.
### 연락처

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **작성자 이메일**: zhengyx91@hotmail.com
---

**작성자**: ZhengYX
