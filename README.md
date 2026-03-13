# HWPX Viewer for VS Code

한컴 HWPX(한/글 2014+) 문서를 VS Code에서 직접 열어 확인하고, AI 에이전트가 프로그래밍 방식으로 문서를 조작할 수 있도록 HTTP API를 제공하는 확장입니다.

## 설치 및 빌드

```bash
npm install
npm run compile
```

VS Code에서 `F5`를 눌러 Extension Development Host를 실행한 뒤 `.hwpx` 파일을 열면 됩니다.

## 뷰어 기능

### 문서 열기
`.hwpx` 파일을 VS Code에서 열면 자동으로 HWPX Viewer가 활성화됩니다.

### 확대/축소
- `Ctrl + 마우스 스크롤`: 마우스 위치 기준 확대/축소
- `Ctrl + =` / `Ctrl + -`: 확대/축소
- `Ctrl + 0`: 100%로 초기화
- 우하단 줌 컨트롤: `−` / `+` 버튼, 퍼센트 클릭 시 100% 초기화
- 범위: 25% ~ 500%
- 용지 크기와 콘텐츠가 함께 확대/축소됩니다

### 드래그 이동
View 모드에서 마우스 드래그로 문서를 자유롭게 이동할 수 있습니다. 용지 바깥 영역까지 충분한 이동 범위가 확보되어 있습니다.

## 모드

우상단의 모드 토글 바로 전환합니다.

### View (기본)
- 읽기 전용 보기
- 마우스 드래그로 화면 이동

### Edit
- 텍스트를 클릭하여 직접 편집 가능
- 상단 중앙에 서식 도구 모음 표시:

| 도구 | 단축키 | 설명 |
|------|--------|------|
| **B** | `Ctrl+B` | 굵게 |
| *I* | `Ctrl+I` | 기울임 |
| U | `Ctrl+U` | 밑줄 |
| ~~S~~ | | 취소선 |
| 숫자 입력 | | 글꼴 크기 (pt) |
| A + 색상 | | 글자 색상 |
| A + 배경색 | | 배경(하이라이트) 색상 |
| 정렬 x4 | | 왼쪽/가운데/오른쪽/양쪽 정렬 |
| Tx | | 서식 제거 |

### Select
- 요소 위에 마우스를 올리면 파란 점선 외곽선 표시
- 클릭하면 해당 요소의 HWPX XML 경로가 클립보드에 복사됨
- 복사 형식: `<문서경로>#<XML파일>#<XPath>`
- 예: `/home/user/report.hwpx#Contents/section0.xml#/hp:sec/hp:p[2]/hp:run[0]/hp:t`
- AI 에이전트에게 이 경로를 전달하면 정확한 요소를 지정하여 수정할 수 있음

## 렌더링 지원 범위

| 요소 | 상태 |
|------|------|
| 문단 텍스트 / 서식 (bold, italic, underline, 색상 등) | 지원 |
| 표 (셀 병합, 테두리, 배경색 포함) | 지원 |
| 이미지 (BinData base64) | 지원 |
| 페이지 레이아웃 (용지 크기, 여백) | 지원 |
| 페이지 분할 (표 행 분할 포함) | 지원 |
| 글머리 기호 / 번호 매기기 | 지원 |
| 탭 정지 | 지원 |
| 스타일 상속 | 지원 |
| 머리글/바닥글, 각주, 도형, 수식, 다단 등 | 미구현 |

미구현 항목의 전체 목록과 표준 문서 참조는 `UNIMPLEMENTED_FEATURES.md`를 참조하세요.

## 커맨드 팔레트

`Ctrl+Shift+P`로 사용 가능한 커맨드:

| 커맨드 | 설명 |
|--------|------|
| `HWPX: Show API Server Port` | API 서버 주소 표시. "Copy Token" 또는 "Copy URL" 선택 가능 |
| `HWPX: Copy API Help URL` | API 사용법 문서 URL을 클립보드에 복사. LLM에 바로 전달 가능 |

## HTTP API

확장이 활성화되면 `127.0.0.1`의 랜덤 포트에 HTTP API 서버가 자동 시작됩니다. AI 에이전트가 HWPX 문서의 요소를 프로그래밍 방식으로 조회하고 수정할 수 있습니다.

### 인증
모든 API 요청(`/api/help` 제외)에는 토큰이 필요합니다.

```bash
# 헤더 방식
curl -H "Authorization: Bearer <token>" http://127.0.0.1:<port>/api/documents

# 쿼리 방식
curl "http://127.0.0.1:<port>/api/documents?token=<token>"
```

토큰은 커맨드 팔레트의 `HWPX: Show API Server Port` → "Copy Token"으로 복사할 수 있습니다.

### 엔드포인트

#### GET /api/help
API 사용법 문서를 Markdown으로 반환합니다. 토큰 불필요. 현재 열린 문서 목록, 실제 포트 번호, curl 예시가 포함됩니다.

LLM에게 이 URL을 전달하면 스스로 API 사용법을 읽고 문서를 조작할 수 있습니다.

#### GET /api/documents
열려있는 HWPX 문서 목록을 반환합니다.

```json
{ "documents": [{ "path": "/home/user/report.hwpx", "dirty": false }] }
```

#### GET /api/files?doc=\<path\>
HWPX(ZIP) 내부의 XML 파일 목록을 반환합니다. 문서가 1개만 열려있으면 `doc` 생략 가능.

```json
{ "files": ["Contents/content.hpf", "Contents/header.xml", "Contents/section0.xml", "mimetype"] }
```

#### GET /api/xml?file=\<xmlPath\>
XML 파일 원본을 그대로 반환합니다. `Content-Type: application/xml`.

#### GET /api/element?file=\<xmlPath\>&xpath=\<xpath\>
XPath로 특정 요소를 조회합니다. XML과 JSON 모두 반환.

```json
{
  "file": "Contents/section0.xml",
  "xpath": "/hp:sec/hp:p[0]/hp:run[0]/hp:t",
  "xml": "안녕하세요",
  "json": [{ "#text": "안녕하세요" }]
}
```

#### PUT /api/element?file=\<xmlPath\>&xpath=\<xpath\>
요소를 수정합니다. Body는 JSON, 3가지 방식 중 택 1:

```bash
# 텍스트만 수정
curl -X PUT ... -d '{"text": "새 텍스트"}'

# XML로 교체
curl -X PUT ... -d '{"xml": "<hp:t>새 내용</hp:t>"}'

# JSON으로 교체
curl -X PUT ... -d '{"json": [{"#text": "새 내용"}]}'
```

#### POST /api/save?doc=\<path\>
수정사항을 HWPX 파일에 저장합니다. 저장 후 뷰어가 자동으로 새로고침됩니다.

#### POST /api/reload?doc=\<path\>
뷰어를 새로고침합니다. 저장 없이 현재 메모리 상태의 변경사항을 뷰어에 반영합니다.

### XPath 형식
- 슬래시(`/`)로 구분
- 태그명에 네임스페이스 접두사 포함: `hp:`, `hh:`, `hc:` 등
- 인덱스는 `[N]` (0-based). 같은 이름 태그 중 N번째를 지정
- 예: `/hp:sec/hp:p[2]/hp:run[0]/hp:t` → 3번째 문단의 첫 번째 런의 텍스트

### AI 에이전트 연동 흐름

1. 커맨드 팔레트 → `HWPX: Copy API Help URL` → URL을 LLM에 전달
2. LLM이 `/api/help`를 읽어 API 사용법을 파악
3. 토큰을 전달 (`HWPX: Show API Server Port` → "Copy Token")
4. LLM이 `/api/documents` → `/api/files` → `/api/element`로 문서 탐색
5. Select 모드에서 수정할 요소를 클릭 → 경로 복사 → LLM에 전달
6. LLM이 `PUT /api/element`로 수정 → `POST /api/save`로 저장

### 보안
- 서버는 `127.0.0.1`에만 바인딩 (외부 접근 불가)
- 랜덤 토큰 인증 필수
- CORS Origin 제한: `vscode-webview://`, `http://127.0.0.1`, `http://localhost`만 허용

## 테스트

```bash
npm test
```

API 서버의 전체 엔드포인트를 커버하는 단위 테스트가 실행됩니다 (107개 테스트).

## 기술 스택

- **VS Code Custom Editor API** (`CustomReadonlyEditorProvider`)
- **JSZip**: HWPX(ZIP) 파일 파싱
- **fast-xml-parser**: XML 파싱/빌드
- **Node.js http**: API 서버
- **OWPML KS X 6101:2024**: 한/글 문서 표준

## 라이선스

MIT
