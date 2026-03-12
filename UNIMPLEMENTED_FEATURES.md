# HWPX 뷰어 미구현/부분 구현 기능 목록

> **작성일**: 2026-03-12
> **표준 기준**: KS X 6101:2024 (OWPML)
> **파서 파일**: `src/hwpxParser.ts`
> **참조 문서 경로**: `openhwp/docs/hwpx/`

---

## 1. Field 관련 기능 (hp:fieldBegin)

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 1.1 | 수식 필드 (FORMULA) | `hp:ctrl > hp:fieldBegin @type="FORMULA"` | 10-body-schema.md §10.7.2.5 | 표 내 SUM/AVG 등 수식 계산 및 결과 표시 | 미구현 | MEDIUM |
| 1.2 | 책갈피 필드 (BOOKMARK) | `hp:ctrl > hp:fieldBegin @type="BOOKMARK"` | 10-body-schema.md §10.7.2.4 | 범위 기반 책갈피, 문서 내 앵커 링크 | 미구현 | HIGH |
| 1.3 | 날짜 필드 (DATE / DOC_DATE) | `hp:ctrl > hp:fieldBegin @type="DATE"` | 10-body-schema.md §10.7.2.6 | 현재 날짜 또는 문서 생성/수정 날짜 삽입 | 미구현 | MEDIUM |
| 1.4 | 요약 필드 (SUMMARY) | `hp:ctrl > hp:fieldBegin @type="SUMMARY"` | 10-body-schema.md §10.7.2.7 | 문서 메타데이터 (Title, Subject, Author 등) 삽입 | 미구현 | LOW |
| 1.5 | 사용자 정보 (USER_INFO) | `hp:ctrl > hp:fieldBegin @type="USER_INFO"` | 10-body-schema.md §10.7.2.8 | 사용자 이름, 소속 등 정보 삽입 | 미구현 | LOW |
| 1.6 | 파일 경로 (PATH) | `hp:ctrl > hp:fieldBegin @type="PATH"` | 10-body-schema.md §10.7.2.9 | 문서 파일 경로 표시 | 미구현 | LOW |
| 1.7 | 교차 참조 (CROSSREF) | `hp:ctrl > hp:fieldBegin @type="CROSSREF"` | 10-body-schema.md §10.7.2.10 | 책갈피, 각주, 그림 등에 대한 교차 참조 링크 | 미구현 | MEDIUM |
| 1.8 | 메일 병합 (MAILMERGE) | `hp:ctrl > hp:fieldBegin @type="MAILMERGE"` | 10-body-schema.md §10.7.2.11 | 외부 데이터 소스 기반 메일 병합 필드 | 미구현 | LOW |
| 1.9 | 메모 (MEMO) | `hp:ctrl > hp:fieldBegin @type="MEMO"` | 10-body-schema.md §10.7.2.12 | 메모 텍스트 삽입 | 미구현 | LOW |
| 1.10 | 교정 부호 (PROOFREADING_MARKS) | `hp:ctrl > hp:fieldBegin @type="PROOFREADING_MARKS"` | 10-body-schema.md §10.7.2.13 | 교정 부호 및 주석 (MarkType, MarkerColor, AuthorName) | 미구현 | LOW |
| 1.11 | 개인정보 보호 (PRIVATE_INFO) | `hp:ctrl > hp:fieldBegin @type="PRIVATE_INFO"` | 10-body-schema.md §10.7.2.14 | AES 암호화 영역 보호 (EncryptMode, MarkChar) | 미구현 | MEDIUM |
| 1.12 | 메타데이터 (METADATA) | `hp:ctrl > hp:fieldBegin @type="METADATA"` | 10-body-schema.md §10.7.2.15 | RDFa 형식 의미론적 메타데이터 (Property, Resource) | 미구현 | LOW |
| 1.13 | 인용 (CITATION) | `hp:ctrl > hp:fieldBegin @type="CITATION"` | 10-body-schema.md §10.7.2.16 | 학술 인용 문헌 참조 (GUID 기반) | 미구현 | LOW |
| 1.14 | 참고문헌 (BIBLIOGRAPHY) | `hp:ctrl > hp:fieldBegin @type="BIBLIOGRAPHY"` | 10-body-schema.md §10.7.2.17 | 참고문헌 목록 (Custom/Bibliography.xml 참조) | 미구현 | LOW |
| 1.15 | 메타 태그 (METATAG) | `hp:ctrl > hp:fieldBegin @type="METATAG"` | 10-body-schema.md §10.7.2.18 | JSON 형식 필드 메타데이터 (예: #전화번호) | 미구현 | LOW |

---

## 2. 각주/미주 고급 기능

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 2.1 | 각주 번호 형식 (autoNumFormat) | `hp:secPr > hp:footNotePr > hp:autoNumFormat` | 10-body-schema.md §10.7.6 | DIGIT/ROMAN/HANGUL 등 번호 형식 선택 | 미구현 | MEDIUM |
| 2.2 | 각주 구분선 (noteLine) | `hp:secPr > hp:footNotePr > hp:noteLine` | 10-body-schema.md §10.7.6 | 각주/미주 구분선 스타일 (길이, 두께, 색상) | 미구현 | MEDIUM |
| 2.3 | 각주 간격 (noteSpacing) | `hp:secPr > hp:footNotePr > hp:noteSpacing` | 10-body-schema.md §10.7.6 | 각주 영역과 본문 사이 간격 | 미구현 | LOW |
| 2.4 | 각주 번호 매김 범위 | `hp:footNotePr @_numbering` | 10-body-schema.md §10.7.6 | CONTINUOUS / ON_SECTION / ON_PAGE 번호 초기화 규칙 | 부분 구현 | MEDIUM |

---

## 3. 도형 고급 기능

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 3.1 | 선 화살표 (headStyle/tailStyle) | `hp:lineShape @_headStyle/@_tailStyle` | 10-body-schema.md §10.10.3 | 선 끝 화살표 모양 (ARROW, STEALTH 등) 및 크기 | 부분 구현 | MEDIUM |
| 3.2 | 선 끝 모양 (endCap) | `hp:lineShape @_endCap` | 10-body-schema.md §10.10.3 | ROUND/FLAT/SQUARE 끝 처리 | 부분 구현 | LOW |
| 3.3 | 도형 투명도 (alpha) | `hp:lineShape @_alpha` | 10-body-schema.md §10.10 | 선/채움 투명도 → CSS opacity | 부분 구현 | LOW |
| 3.4 | 렌더링 변환 행렬 (scaMatrix) | `hp:renderingInfo > hp:scaMatrix` | 10-body-schema.md §10.9.1 | 스케일링 변환 행렬 적용 | 미구현 | LOW |
| 3.5 | 렌더링 변환 행렬 (transMatrix) | `hp:renderingInfo > hp:transMatrix` | 10-body-schema.md §10.9.1 | 이동 변환 행렬 적용 | 미구현 | LOW |
| 3.6 | 도형 효과 (reflection) | `hp:shape > hp:reflection` | 10-body-schema.md §10.9.1 | 반사 효과 | 미구현 | LOW |
| 3.7 | 도형 효과 (glow/softEdge) | `hp:shape > hp:effect` | 10-body-schema.md §10.9.1 | 글로우, 부드러운 가장자리, 그림자 효과 | 미구현 | LOW |
| 3.8 | 연결선 고급 (connectLine) | `hp:connectLine` | 10-body-schema.md §10.10.9 | 곡선 연결(spline), 연결점 처리, 동적 라우팅 | 부분 구현 | LOW |

---

## 4. 이미지 고급 기능

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 4.1 | 이미지 효과 (shadow) | `hp:pic > hp:effect > shadow` | 10-body-schema.md §10.9.6.5 | 그림자 효과 (방향, 거리, 색상, 투명도) | 미구현 | MEDIUM |
| 4.2 | 이미지 효과 (glow) | `hp:pic > hp:effect > glow` | 10-body-schema.md §10.9.6.5 | 글로우 효과 | 미구현 | LOW |
| 4.3 | 이미지 효과 (reflection) | `hp:pic > hp:effect > reflection` | 10-body-schema.md §10.9.6.5 | 반사 효과 | 미구현 | LOW |
| 4.4 | 이미지 효과 (softEdge) | `hp:pic > hp:effect > softEdge` | 10-body-schema.md §10.9.6.5 | 부드러운 가장자리 | 미구현 | LOW |
| 4.5 | 텍스트 래핑 고급 | `hp:pic @_textWrap` | 10-body-schema.md §10.9.6 | SQUARE, TIGHT, THROUGH 등 고급 텍스트 래핑 | 부분 구현 | MEDIUM |
| 4.6 | Z-Order 오버랩 처리 | `hp:pic @_zOrder` | 10-body-schema.md §10.9.6 | 겹치는 객체의 정확한 z-index 렌더링 | 부분 구현 | LOW |

---

## 5. 임베디드 객체

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 5.1 | 차트 (chart) | `hp:run > hp:chart` | 10-body-schema.md §10.5, 17-compatibility.md | 차트 렌더링 (OOXML 차트 형식) — fallback 이미지 표시 | 미구현 | MEDIUM |
| 5.2 | 비디오 (video) | `hp:run > hp:video` | 10-body-schema.md §10.5 | 임베디드 비디오 — placeholder 표시 | 미구현 | LOW |
| 5.3 | OLE 미리보기 이미지 | `hp:ole @_binaryItemIDRef` | 10-body-schema.md §10.9.7 | OLE 객체의 미리보기 이미지 또는 바이너리 렌더링 | 부분 구현 | MEDIUM |

---

## 6. 폼 컨트롤

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 6.1 | 리스트박스 (listBox) | `hp:run > hp:listBox` | 10-body-schema.md §10.11.7 | 다중 선택 목록 상자 | 미구현 | LOW |
| 6.2 | 스크롤바 (scrollBar) | `hp:run > hp:scrollBar` | 10-body-schema.md §10.11.9 | 스크롤바 컨트롤 (최솟값, 최댓값, 현재값) | 미구현 | LOW |

---

## 7. 섹션/레이아웃

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 7.1 | 줄 격자 (grid) | `hp:secPr > hp:grid` | 10-body-schema.md §10.6 | lineGrid/charGrid 격자 설정 | 미구현 | LOW |
| 7.2 | 줄 번호 (lineNumberShape) | `hp:secPr > hp:lineNumberShape` | 10-body-schema.md §10.6 | 줄 번호 표시 (간격, 시작 번호, 형식) | 미구현 | LOW |
| 7.3 | 바탕쪽 (masterPage) | `hp:secPr > hp:masterPage @_idRef` | 10-body-schema.md §10.6.10, 11-masterpage-schema.md | 바탕쪽 XML 파일 파싱 및 적용 | 미구현 | LOW |
| 7.4 | 프레젠테이션 효과 | `hp:secPr > hp:presentation` | 10-body-schema.md §10.6.11 | 화면 전환 효과 (overLeft, blindLeft 등) | 미구현 | LOW |
| 7.5 | 제책 방법 (gutterType) | `hp:pagePr @_gutterType` | 10-body-schema.md §10.6 | LEFT_ONLY / LEFT_RIGHT / TOP_BOTTOM 제본 여백 | 미구현 | LOW |
| 7.6 | 세로쓰기 (textDirection) | `hp:secPr @_textDirection` | 10-body-schema.md §10.6 | VERTICAL 텍스트 방향 (CJK 세로쓰기) | 부분 구현 | MEDIUM |
| 7.7 | 머리글/바닥글 홀짝 구분 | `hp:header/hp:footer @_applyPageType` | 10-body-schema.md §10.7.5 | EVEN/ODD 페이지별 다른 머리글/바닥글 | 부분 구현 | MEDIUM |

---

## 8. 변경 추적

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 8.1 | 삽입 추적 시각화 | `hp:t > hp:insertBegin / hp:insertEnd` | 10-body-schema.md §10.8.5 | 삽입된 텍스트 색상/밑줄 표시 + 작성자/시간 | 부분 구현 | MEDIUM |
| 8.2 | 삭제 추적 시각화 | `hp:t > hp:deleteBegin / hp:deleteEnd` | 10-body-schema.md §10.8.5 | 삭제된 텍스트 취소선 표시 + 작성자/시간 | 부분 구현 | MEDIUM |

---

## 9. 특수 텍스트/콘텐츠

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 9.1 | 수식 렌더링 | `hp:equation > hp:script` | 10-body-schema.md §10.5 | 한/글 수식 문법 → MathML/LaTeX 변환 렌더링 | 부분 구현 | MEDIUM |
| 9.2 | 색인 표시 (indexmark) | `hp:run > hp:indexmark` | 10-body-schema.md §10.7.11 | 찾아보기(Index) 키워드 마킹 | 미구현 | LOW |
| 9.3 | 제목 차례 (titleMark) | `hp:t > hp:titleMark` | 10-body-schema.md §10.8.3 | 자동 목차 생성용 제목 마킹 | 부분 구현 | LOW |
| 9.4 | 숨은 설명 UI | `hp:run > hp:hiddenComment > hp:subList` | 10-body-schema.md §10.7.12 | 숨은 설명 아이콘 + tooltip/popover 렌더링 | 부분 구현 | MEDIUM |
| 9.5 | 페이지 숨기기 (pageHiding) | `hp:run > hp:pageHiding` | 10-body-schema.md §10.7.9 | 머리글/바닥글/바탕쪽/테두리/쪽번호 숨기기 | 미구현 | LOW |
| 9.6 | 쪽 번호 제어 (pageNumCtrl) | `hp:run > hp:pageNumCtrl` | 10-body-schema.md §10.7.8 | 홀수/짝수 쪽 구분 설정 | 미구현 | LOW |

---

## 10. 레이아웃 호환성 옵션

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 10.1 | layoutCompatibility 전체 | `hh:compatibleDocument > hh:layoutCompatibility` | 09-header-schema.md §9.2.5.2 | 약 40개 레이아웃 호환성 플래그 | 미구현 | LOW |

주요 플래그 목록:
- `applyFontWeightToBold` — 글꼴 굵기를 bold에 적용
- `extendLineheightToOffset` — 줄 높이에 오프셋 포함
- `applyFontspaceToLatin` — 라틴 글꼴에 자간 적용
- `baseCharUnitOnEAsian` — 동아시아 글자 기준 단위
- `adjustLineheightToFont` — 글꼴에 맞춘 줄 높이 조정
- `applyParaBorderToOutside` — 문단 테두리 외곽 적용
- `doNotApplyAutoSpaceEAsianEng` — 한영 자동 간격 비적용
- `fixedUnderlineWidth` — 고정 밑줄 두께
- 기타 약 32개 (표준 문서 참조)

---

## 11. 기타 헤더 요소

| # | 기능명 | XML 경로 | 표준 참조 | 설명 | 상태 | 우선순위 |
|---|--------|----------|-----------|------|------|----------|
| 11.1 | 시작 번호 (beginNum) | `hh:beginNum` | 09-header-schema.md §9.2 | page/footnote/endnote/pic/tbl/equation 시작 번호 | 미구현 | LOW |
| 11.2 | 금칙어 (forbiddenWordList) | `hh:forbiddenWordList` | 09-header-schema.md §9.2 | 금칙어 목록 (줄 나눔 규칙) | 미구현 | LOW |
| 11.3 | 문서 옵션 (docOption) | `hh:docOption > linkinfo/licensemark` | 09-header-schema.md §9.2 | 문서 링크/라이선스 정보 | 미구현 | LOW |
| 11.4 | 메모 속성 (memoProperties) | `hh:memoProperties` | 09-header-schema.md §9.2 | 메모 표시 방법 설정 | 미구현 | LOW |

---

## 우선순위별 로드맵

### Phase 1: HIGH — 문서 기본 가독성
- 1.2 책갈피 필드 (BOOKMARK)

### Phase 2: MEDIUM — 고급 렌더링
- 1.1 수식 필드 (FORMULA)
- 1.3 날짜 필드 (DATE/DOC_DATE)
- 1.7 교차 참조 (CROSSREF)
- 1.11 개인정보 보호 (PRIVATE_INFO)
- 2.1~2.2 각주 번호 형식/구분선
- 3.1 선 화살표
- 4.1 이미지 그림자 효과
- 4.5 텍스트 래핑 고급
- 5.1 차트 fallback
- 5.3 OLE 미리보기
- 7.6 세로쓰기
- 7.7 머리글/바닥글 홀짝 구분
- 8.1~8.2 변경 추적 시각화
- 9.1 수식 렌더링
- 9.4 숨은 설명 UI

### Phase 3: LOW — 선택 기능
- 나머지 Field 타입 (1.4~1.6, 1.8~1.10, 1.12~1.15)
- 폼 컨트롤 (6.1~6.2)
- 도형 고급 효과 (3.4~3.8)
- 이미지 효과 (4.2~4.4)
- 섹션/레이아웃 (7.1~7.5)
- 레이아웃 호환성 (10.1)
- 기타 헤더 요소 (11.1~11.4)
- 특수 요소 (9.2~9.3, 9.5~9.6)
