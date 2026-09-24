# 레이아웃 = 헤더 + 패널 격자

슬라이드는 두 종류다.

- **전면 슬라이드** (`layout`): `cover` · `toc` · `divider` · `closing`
- **본문 슬라이드** (`layout` 생략 = `content`): 헤더 + `body` 격자. 레이아웃이 고정되어 있지 않고 **패널을 조합**해 만든다.

## 덱 공통 필드

| 필드 | 설명 |
|---|---|
| `sections` | 섹션명 배열 → 우상단 탭 내비게이션, `toc` 기본 항목 |
| `footer` | 하단 좌측 문구 (과제명·기관명) |
| `theme` | 토큰 덮어쓰기 → [style-guide.md](style-guide.md) |

## 본문 슬라이드

```json
{
  "section": 0,
  "title": "결론 문장 헤드라인, 핵심어는 **강조**",
  "lead": "헤드라인을 보충하는 한 줄 (선택)",
  "source": "출처 (데이터가 있으면 필수)",
  "notes": "발표자 노트",
  "body": [
    { "h": 3, "cols": [ {패널}, {패널, "w": 1.5} ] },
    [ {패널} ]
  ]
}
```

- `body` = 행 배열. 행은 패널 배열이거나 `{ "h": 높이비율, "cols": [...] }`.
- 패널의 `w` = 같은 행 안에서의 너비 비율 (기본 1). `arrow`·`label`은 `w`가 없으면 좁은 고정 폭.
- 모든 패널에 공통으로 쓸 수 있는 필드: `title`(패널 제목), `title_style`(`line`|`bar`), `boxed`(연회색 배경).
- 텍스트 어디서나 `**강조**` = 강조색 볼드. 불릿 항목이 `"- "`로 시작하면 2단계 불릿.

## 패널 카탈로그

| 메시지 형태 | 패널 |
|---|---|
| 서술·근거 나열 | `bullets` |
| 한 줄 결론·시사점 | `callout` |
| 다차원 비교, 과제 개요(키-값), 성능 지표 | `table` |
| 추세·비교·구성비 | `chart` |
| 핵심 수치 2~4개 | `kpi` |
| 병렬 항목 N개 (역할, 기능, 문제점) | `cards` |
| 순서·단계 | `process` |
| 일정·간트 | `timeline` |
| 순환·협업 체계, 허브-스포크 | `cycle` |
| 시스템 구성도·아키텍처 계층 | `stack` |
| 화면·사진·도면 | `image` |
| 행 구분 라벨 (배경/추진내용, 현황/한계) | `label` |
| 흐름 연결 (A → B) | `arrow` |

### `bullets`
`items`: 문자열 배열 · `style`: `"plain"`이면 불릿 없는 문단 · `size`

### `callout`
`text` · `style`: `"dark"` (남색 배경) · `align`

### `table`
`columns`(헤더, 생략 가능) · `rows`(2차원) · `widths`(열 비율) · `rowhead`(첫 열을 항목 열로) · `highlight`(데이터 행 인덱스) · `size` · `stretch`(기본 true)
키-값 개요 표는 `columns` 없이 `rowhead: true`.

### `chart`
`kind`: `bar | hbar | stacked | line | pie` · `categories` · `series: [{name, values}]` · `highlight`(단일 시리즈의 강조 카테고리) · `number_format`(예: `"0%"`, `"#,##0"`) · `unit`(우상단 "(단위: …)") · `labels`(기본 true = 값 레이블 표시, 축 숨김)

### `kpi`
`items: [{value, label}]` · `highlight` · `direction`: `"column"`이면 세로 배치

### `cards`
`items: [{title, items[] | text}]` · `cols`(기본 = 항목 수; 2면 2×N 격자) · `numbered`(기본 true, 번호 배지) · `highlight` · `size`

### `process`
`steps: [{title, label?, items[] | text}]` · `highlight`

### `timeline`
`periods`(열: 월·분기·연차) · `tasks: [{name, span:[시작,끝] | [[a,b],[c,d]], group?, highlight?, note?}]` (1부터 시작, 끝 포함) · `label`(첫 열 제목) · `label_width`

### `cycle`
`center` · `nodes`(3~8개, `\n` 줄바꿈 가능) · `highlight`

### `stack`
`layers: [{label, items[], highlight?}]` (위에서 아래로) · `label_width`

### `image`
`src`(deck.json 기준 상대 경로) · `alt` · `caption`
`src`가 없거나 파일이 없으면 점선 자리표시자를 그린다 → 사용자가 나중에 스크린샷을 넣는다.

### `label`
`text` · `highlight`. 세로로 긴 칸이면 글자를 세로로 쌓는다.

### `arrow`
`direction`: `"right"`(기본) | `"down"`

## 전면 슬라이드

- `cover`: `kicker`(사업명), `title`(`\n` 가능), `subtitle`, `org`, `author`, `date`
- `toc`: `items` (생략하면 `sections`)
- `divider`: `no`("01"), `title`, `items`(하위 목차)
- `closing`: `title`, `subtitle`

## 자주 쓰는 조합

| 목적 | body |
|---|---|
| 배경 → 문제 | `[label, cards]`, `{h:.25, [arrow down]}`, `[label, cards]` |
| 근거 3종 + 결론 | `{h:3, [chart, bullets, table]}`, `{h:.55, [callout]}` |
| 구성도 + 목표 | `[stack w1.3, table]` |
| 체계 + 역할 | `[cycle, cards cols2 w1.5]` |
| 일정 + 단계 목표 | `{h:3, [timeline]}`, `[process]` |
| 화면 + 기능 | `[image w1.3, image, bullets w.9]` |
| 효과 | `[kpi]`, `{h:1.7, [chart, cards cols1]}` |
| 비교 | `[bullets boxed, arrow, bullets boxed]` |

## 새 패널 추가
`scripts/build_deck.py`에 `def p_name(s, box, p)` 작성 → `PANELS`에 등록 → 이 문서 카탈로그에 추가 → `examples/sample.json`에 사용 예 추가 (테스트가 모든 패널이 샘플에 쓰였는지 검사).
