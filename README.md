# HWPX Viewer for VS Code

> **Beta** — 이 확장은 현재 베타 버전입니다. 일부 HWPX 문서에서 렌더링이 불완전할 수 있으며, 머리글/바닥글, 각주, 도형, 수식 등은 아직 미구현 상태입니다. 버그 리포트 및 피드백은 [GitHub Issues](https://github.com/solanian/hwpx-for-vscode/issues)에서 환영합니다.

한컴 HWPX(한/글 2014+) 문서를 VS Code에서 직접 열어 확인하고, AI 에이전트가 프로그래밍 방식으로 문서를 조작할 수 있도록 HTTP API를 제공하는 확장입니다.

## Features

### Document Viewer
`.hwpx` 파일을 VS Code에서 열면 자동으로 HWPX Viewer가 활성화됩니다.

- 문단 텍스트, 서식 (bold, italic, underline, 색상 등)
- 표 (셀 병합, 테두리, 배경색)
- 이미지 (BinData base64)
- 페이지 레이아웃 (용지 크기, 여백)
- 페이지 분할 (표 행 분할 포함)
- 글머리 기호 / 번호 매기기
- 탭 정지, 스타일 상속

### Zoom & Navigation
- `Ctrl + 마우스 스크롤`: 마우스 위치 기준 확대/축소
- `Ctrl + =` / `Ctrl + -`: 확대/축소
- `Ctrl + 0`: 100%로 초기화
- 우하단 줌 컨트롤
- 범위: 25% ~ 500%
- View 모드에서 마우스 드래그로 문서 이동

### 3 Modes
우상단의 모드 토글 바로 전환합니다.

| 모드 | 설명 |
|------|------|
| **View** (기본) | 읽기 전용 보기. 마우스 드래그로 화면 이동 |
| **Edit** | 텍스트 클릭하여 직접 편집. 상단 서식 도구 모음 표시 |
| **Select** | 요소 클릭 시 HWPX XML 경로를 클립보드에 복사. AI 에이전트에 전달하여 정확한 요소 지정 가능 |

### Edit Mode Toolbar

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

### HTTP API for AI Agents
확장이 활성화되면 `127.0.0.1`의 랜덤 포트에 HTTP API 서버가 자동 시작됩니다. AI 에이전트(Claude, GPT 등)가 HWPX 문서의 요소를 프로그래밍 방식으로 조회하고 수정할 수 있습니다.

## Commands

`Ctrl+Shift+P`로 사용 가능:

| 커맨드 | 설명 |
|--------|------|
| `HWPX: Show API Server Port` | API 서버 주소 표시. "Copy Token" 또는 "Copy URL" 선택 가능 |
| `HWPX: Copy API Help URL` | API 사용법 문서 URL을 클립보드에 복사. LLM에 바로 전달 가능 |

## Rendering Support

| 요소 | 상태 |
|------|------|
| 문단 텍스트 / 서식 (bold, italic, underline, 색상 등) | Supported |
| 표 (셀 병합, 테두리, 배경색 포함) | Supported |
| 이미지 (BinData base64) | Supported |
| 페이지 레이아웃 (용지 크기, 여백) | Supported |
| 페이지 분할 (표 행 분할 포함) | Supported |
| 글머리 기호 / 번호 매기기 | Supported |
| 탭 정지 | Supported |
| 스타일 상속 | Supported |
| 머리글/바닥글, 각주, 도형, 수식, 다단 등 | Not yet |

## HTTP API Reference

### Authentication
모든 API 요청(`/api/help` 제외)에는 토큰이 필요합니다.

```bash
# 헤더 방식
curl -H "Authorization: Bearer <token>" http://127.0.0.1:<port>/api/documents

# 쿼리 방식
curl "http://127.0.0.1:<port>/api/documents?token=<token>"
```

토큰은 커맨드 팔레트의 `HWPX: Show API Server Port` → "Copy Token"으로 복사할 수 있습니다.

### Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/help` | API 사용법 문서 (Markdown). 토큰 불필요 |
| GET | `/api/documents` | 열려있는 HWPX 문서 목록 |
| GET | `/api/files?doc=<path>` | HWPX(ZIP) 내부 XML 파일 목록 |
| GET | `/api/xml?file=<xmlPath>` | XML 파일 원본 |
| GET | `/api/element?file=<xmlPath>&xpath=<xpath>` | XPath로 요소 조회 (XML + JSON) |
| PUT | `/api/element?file=<xmlPath>&xpath=<xpath>` | 요소 수정 (text/xml/json) |
| POST | `/api/save?doc=<path>` | 파일 저장 + 뷰어 자동 새로고침 |
| POST | `/api/reload?doc=<path>` | 뷰어 새로고침 (저장 없이) |

### PUT /api/element — 수정 방식

```bash
# 텍스트만 수정
curl -X PUT ... -d '{"text": "새 텍스트"}'

# XML로 교체
curl -X PUT ... -d '{"xml": "<hp:t>새 내용</hp:t>"}'

# JSON으로 교체
curl -X PUT ... -d '{"json": [{"#text": "새 내용"}]}'
```

### XPath Format
- 슬래시(`/`)로 구분, 네임스페이스 접두사 포함: `hp:`, `hh:`, `hc:` 등
- 인덱스는 `[N]` (0-based): `/hp:sec/hp:p[2]/hp:run[0]/hp:t`
- Descendant 검색 (`//`): `/hs:sec//hp:t[0]`
- 상대 경로: `hp:p[2]/hp:run[0]/hp:t`
- Wrapper 자동 스킵: `hp:subList` 등 중간 래퍼 요소 자동 건너뜀

### AI Agent Workflow

1. 커맨드 팔레트 → `HWPX: Copy API Help URL` → URL을 LLM에 전달
2. LLM이 `/api/help`를 읽어 API 사용법을 파악
3. 토큰을 전달 (`HWPX: Show API Server Port` → "Copy Token")
4. LLM이 `/api/documents` → `/api/files` → `/api/element`로 문서 탐색
5. Select 모드에서 수정할 요소를 클릭 → 경로 복사 → LLM에 전달
6. LLM이 `PUT /api/element`로 수정 → `POST /api/save`로 저장

### Security
- 서버는 `127.0.0.1`에만 바인딩 (외부 접근 불가)
- 랜덤 토큰 인증 필수
- CORS Origin 제한: `vscode-webview://`, `http://127.0.0.1`, `http://localhost`만 허용

## Development

```bash
npm install
npm run compile
npm test          # 122 tests
```

VS Code에서 `F5`를 눌러 Extension Development Host를 실행한 뒤 `.hwpx` 파일을 열면 됩니다.

## Tech Stack

- **VS Code Custom Editor API** (`CustomReadonlyEditorProvider`)
- **JSZip**: HWPX(ZIP) 파일 파싱
- **fast-xml-parser**: XML 파싱/빌드
- **Node.js http**: API 서버
- **OWPML KS X 6101:2024**: 한/글 문서 표준

## License

[MIT](LICENSE)
