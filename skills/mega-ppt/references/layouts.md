# 레이아웃 카탈로그

메시지의 **형태**로 레이아웃을 고른다. 모든 슬라이드 공통 필드:

| 필드 | 설명 |
|---|---|
| `layout` | 아래 이름 중 하나 (필수) |
| `title` | 액션 타이틀 — 결론 문장, 2줄 이내 |
| `kicker` | 타이틀 위 작은 라벨 (섹션명 등) |
| `source` | 하단 출처 표기 |
| `notes` | 발표자 노트 |
| `highlight` | 강조할 항목의 0-based 인덱스 (레이아웃별 의미 상이) |

## 선택 가이드

| 메시지 형태 | 레이아웃 |
|---|---|
| 덱 시작 | `cover` |
| 챕터 전환 | `section` |
| 핵심 숫자 2–4개 | `kpi` |
| 추세·비교·구성비 데이터 | `chart` |
| A vs B, 전/후, 문제/해결 | `two-column` |
| 단계·로드맵·프로세스 (3–5단계) | `process` |
| 다차원 비교 (옵션×기준) | `table` |
| 요약·결정 요청·권고안 | `bullets` |
| 인용·핵심 인사이트 한 문장 | `quote` |
| 마무리 | `closing` |

## 레이아웃별 필드

### `cover`
`title`, `subtitle`, `author`, `date`

### `section`
`no` ("01"), `title`, `subtitle`

### `bullets`
`bullets`: 문자열 배열 (3–5개 권장)

### `two-column`
`left`, `right`: `{ "heading": str, "bullets": [str] }` · `highlight`: 0(left) / 1(right)

### `kpi`
`items`: `[{ "value": "14%", "label": "설명" }]` (2–4개) · `highlight`: 항목 인덱스 · `takeaway`: 하단 한 줄 해석

### `chart`
```json
"chart": {
  "type": "bar | hbar | stacked | line | pie",
  "categories": ["Q1", "Q2"],
  "series": [{ "name": "매출", "values": [10, 12] }],
  "highlight": 1,
  "number_format": "0%"
}
```
- `highlight`: 단일 시리즈 bar/hbar/pie에서 강조할 카테고리
- `number_format`: 데이터 레이블 포맷 (선택)
- `labels`: 기본 true. false면 레이블 대신 값 축·격자 표시
- `title`: 차트 제목 (선택, 보통 생략 — 슬라이드 타이틀이 대신함)

`takeaway`: 우측 해석 박스 (있으면 차트 폭이 줄어듦). 다중 시리즈는 첫 시리즈가 강조색.

### `process`
`steps`: `[{ "title": str, "body": str }]` (3–5개, body는 `\n`으로 줄바꿈) · `highlight`: 단계 인덱스

### `table`
`columns`: 헤더 배열 · `rows`: 2차원 배열 (최대 ~8행) · `highlight`: 데이터 행 인덱스

### `quote`
`quote`, `by`

### `closing`
`title` (기본 "감사합니다"), `subtitle`

## 새 레이아웃 추가
`scripts/build_deck.py`에 `def name(s, d)` 함수 작성 → `LAYOUTS`에 등록 → 이 문서에 필드와 선택 가이드 행 추가 → `examples/sample.json`에 예시 슬라이드 추가.
