# Changelog

## [0.0.1] - 2025-01-01 (Beta)

### Added
- HWPX 문서 뷰어 (Custom Editor)
  - 문단 텍스트, 서식 (bold, italic, underline, 색상 등)
  - 표 (셀 병합, 테두리, 배경색)
  - 이미지 (BinData base64)
  - 페이지 레이아웃 (용지 크기, 여백)
  - 페이지 분할 (표 행 분할 포함)
  - 글머리 기호 / 번호 매기기
  - 탭 정지
  - 스타일 상속
- View / Edit / Select 3가지 모드
- 확대/축소 (25% ~ 500%)
- 드래그 이동
- HTTP API 서버 (AI 에이전트 연동)
  - 문서 목록 조회, XML 파일 탐색, 요소 조회/수정
  - 토큰 기반 인증
  - POST /api/save (저장), POST /api/reload (새로고침)
- VS Code 커맨드
  - `HWPX: Show API Server Port`
  - `HWPX: Copy API Help URL`
