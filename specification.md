# HWPX Viewer VSCode Extension Specification

## 1. 개요 (Overview)
본 문서는 한컴오피스 HWPX 파일을 VSCode 내에서 읽고 볼 수 있도록 하는 Custom Editor 기반의 확장 프로그램(Extension) 구조 및 구현 사양을 명세합니다.
제공된 public repository (ex: `hwpx-mcp`, `pyhwpx` 등)를 분석하여 도출된 방법론을 바탕으로 구현됩니다.

## 2. HWPX 파일 구조 분석 (Reverse Engineering Results)
HWPX 포맷은 OWPML(Open Word Processor Markup Language) 표준을 따르며, 실제로는 XML 파일들을 압축한 ZIP 파일 포맷입니다.

* **주요 특징**:
  - `mimetype`: `application/hwp+zip`
  - `Contents/header.xml`: 문서의 메타데이터, 폰트, 스타일, 모양 정보를 저장 (CharShape, ParaShape 등).
  - `Contents/sectionN.xml`: 본문 데이터. Paragraph(`hp:p`), TextRun(`hp:t`), Table(`hp:tbl`), Image 등의 엘리먼트로 구성.
  - `BinData/`: 문서 내에 삽입된 이미지 및 바이너리 데이터를 포함.

* **파싱 전략 (hwpx-mcp 기반)**:
  1. `jszip` 라이브러리를 이용하여 HWPX 파일의 압축을 해제.
  2. `Contents/header.xml`을 파싱하여 스타일 정보 수집.
  3. `Contents/sectionN.xml`을 순차적으로 읽어 `hp:p`, `hp:tbl` 등에서 텍스트와 구조를 추출.
  4. 웹뷰(Webview)에서 렌더링하기 위해 추출된 데이터를 HTML/CSS로 변환.

## 3. Extension 구조 (Architecture)

### 3.1 핵심 파일
| 파일 | 역할 |
|---|---|
| `src/hwpxEditorProvider.ts` | VSCode `CustomReadonlyEditorProvider` 구현. Webview 패널 관리, HTML 렌더링 출력 |
| `src/hwpxParser.ts` | HWPX ZIP 해제 → XML 파싱 → HTML/CSS 변환. 페이지네이션 JS 생성 포함 |
| `src/extension.ts` | Extension 진입점. EditorProvider 등록 |

### 3.2 렌더링 파이프라인
```
HWPX(ZIP) → jszip 해제 → content.hpf 파싱(메타데이터, 이미지맵)
                        → header.xml 파싱(스타일 CSS 생성)
                        → sectionN.xml 파싱(본문 HTML 생성)
                        → 페이지네이션 JS 삽입
                        → Webview에 최종 HTML 전달
```

## 4. 기능 구현 현황

> 기능별 구현 상세는 [`feature_matrix.csv`](./feature_matrix.csv) 참조

### 요약 통계

| 카테고리 | 완료 | 미구현 | 합계 |
|---|---|---|---|
| 문서구조 | 3 | 0 | 3 |
| 페이지레이아웃 | 4 | 0 | 4 |
| 문자스타일 | 5 | 4 | 9 |
| 문단스타일 | 9 | 1 | 10 |
| 표 | 10 | 1 | 11 |
| 이미지 | 1 | 1 | 2 |
| **합계** | **32** | **7** | **39** |

## 5. 핵심 구현 세부사항

### 5.1 가로 방향 판별 로직
```typescript
// HWP에서 landscape 속성이 존재하면(NARROWLY든 WIDELY든) 가로 방향을 의미함
const landscapeVal = String(pagePr['@_landscape'] || '').trim().toLowerCase();
const isLandscape = landscapeVal !== '' && landscapeVal !== '0' 
                    && landscapeVal !== 'false' && landscapeVal !== 'undefined';
if (isLandscape && widthMm < heightMm) { swap(widthMm, heightMm); }
```

### 5.2 페이지네이션 (overflow clip 방식)
```
1. template(전체 콘텐츠가 들어간 div)을 화면 밖에 렌더링하여 scrollHeight 측정
2. contentMaxH = pageHeightPx - paddingTop - paddingBottom
3. numPages = ceil(totalContentH / contentMaxH)
4. 각 페이지: div.hwpx-page > div(clipper, overflow:hidden, height=contentMaxH) 
              > div(inner, position:relative, top=-(i*contentMaxH)px) > 콘텐츠 clone
```

### 5.3 글머리 기호 구현
```
header.xml:  hh:bullets > hh:bullet[id, char]  →  bulletMap = { id: char }
             hh:paraPr > hh:heading[type="BULLET", idRef, level]
CSS 출력:    .para-{id}::before { content: '• '; position: absolute; left: Xpt; }
```

### 5.4 HWPUNIT 변환
- 1mm = 283.465 HWPUNIT
- charPr height: 100 단위 (height=1000 → 10pt)
- margin value: 100 단위 (value=800 → 8pt)

## 6. 의존성 (Dependencies)
| 패키지 | 용도 |
|---|---|
| `vscode` (`@types/vscode`) | VSCode Extension API |
| `jszip` | ZIP 아카이브 압축 해제 |
| `fast-xml-parser` | XML → JSON 변환 |

## 7. 개발 환경 (Development Environment)
본 프로젝트는 시스템 전역 설정을 오염시키지 않기 위해 **Docker 기반의 격리된 컨테이너 환경(DevContainer)**을 사용하여 개발 및 빌드를 진행합니다.
- `Node.js` 및 확장 프로그램 빌드에 필요한 모든 의존성은 컨테이너 내부에서만 설치 및 실행됩니다.
- 빌드: `npm run compile` (tsc -p ./)
- 디버그: VSCode Extension Development Host (F5)
