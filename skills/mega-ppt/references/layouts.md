# 레이아웃 = 헤더 + 패널 격자

슬라이드는 두 종류다.

- **전면 슬라이드** (`layout`): `cover` · `toc` · `divider` · `statement` · `closing`
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
| SWOT, 2×2 사분면, 포지셔닝 맵 | `matrix` |
| 계층·역량 수준, 기술 스택 | `pyramid` |
| TAM/SAM/SOM, 전환 단계 | `funnel` |
| 조직도, 추진체계, WBS | `tree` |
| 비전–목표–전략–기반 전략 체계도 | `house` |
| 연혁, 선행 연구, 주요 이정표 | `milestones` |
| 연구진·인력 구성 | `profiles` |
| 핵심 요약, 이슈 목록, 요청 사항 | `numbered` |
| BM, 서비스·데이터 흐름 | `flow` |
| 달성률, 진척도, 역량 수준 | `progress` |
| 트랙 × 단계 로드맵 | `roadmap` |

### `bullets`
`items`: 문자열 배열 · `style`: `"plain"`이면 불릿 없는 문단 · `size`

### `callout`
`text` · `style`: `"dark"` (남색 배경) · `align`

### `table`
`columns`(헤더, 생략 가능) · `rows`(2차원) · `widths`(열 비율) · `rowhead`(첫 열을 항목 열로) · `highlight`(데이터 행 인덱스) · `size` · `stretch`(기본 true)
- `total: true` → 마지막 행을 합계 행으로 표시(굵게, 윗선)
- `merge: [[r1, c1, r2, c2], ...]` → 셀 병합. 좌표는 헤더 행을 0으로 센다. 병합으로 가려지는 셀은 `""`로 둔다
- 셀 값이 `● ◐ ○ ✓ ✗ △`이면 기호로 크게, 색을 입혀 표시한다 (기능 비교표)
키-값 개요 표는 `columns` 없이 `rowhead: true`.

### `chart`
`kind`: `bar | hbar | stacked | line | pie | waterfall` · `categories` · `series: [{name, values}]` · `highlight`(단일 시리즈의 강조 카테고리) · `number_format`(예: `"0%"`, `"#,##0"`) · `unit`(우상단 "(단위: …)") · `labels`(기본 true = 값 레이블 표시, 축 숨김)
- `waterfall`: 단일 시리즈의 증감값. `totals: [인덱스]`는 누계 막대(강조색)로 그린다. 음수는 감소로 표시
- 선 차트에 시리즈가 여럿이면 끝점에만 값을 표시한다

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

### `matrix`
- 사분면: `quadrants: [{title, items[] | text}] ×4` (좌상·우상·좌하·우하 순) · `letters: "SWOT"`(배지) · `highlight`
- 포지셔닝: `points: [{label, x, y, highlight?}]` (x·y는 0~1) · `quadrant_labels` ×4 · `highlight`(사분면)
- 공통: `x_label`, `y_label`(축 제목)

### `pyramid` / `funnel`
`levels: [{title, value?, text?}]` (위에서 아래로) · `highlight`. `text`가 있으면 오른쪽에 점선으로 설명을 연결한다. `pyramid`는 아래로 갈수록 넓고, `funnel`은 좁아진다.

### `tree`
`root: {name, role?, side?: [노드], children: [{name, role?, children?: [str | 노드], items?}]}` · `highlight`(자식 인덱스) · `node_h`
3단까지 그린다. 3단은 2단 노드 아래 세로 목록이 된다. `side`는 루트 옆 점선 노드(자문위원회 등)다.

### `house`
`vision` · `goals[]`(선택) · `pillars: [{title, items[]}]` · `base[]`(기반 띠) · `highlight`(기둥) · `vision_label`, `goals_label`

### `milestones`
`events: [{date, title, text?, highlight?}]` (4~7개 권장). 가로축 위아래로 번갈아 배치한다.

### `profiles`
`people: [{name, role, org?, items[]}]` · `cols` · `highlight`. 사진 대신 이니셜 아바타를 쓴다.

### `numbered`
`items: [{title, text?}]` · `cols`(기본 1) · `highlight`

### `flow`
`nodes: [{title, items[] | text}]` · `edges: [str | {label, back}]` (노드 사이, `back`이 있으면 역방향 화살표 추가) · `highlight`

### `progress`
`items: [{label, value, display?}]` · `max`(기본 100) · `label_width` · `highlight`

### `roadmap`
`phases[]` · `tracks: [{name, cells: [str | [str]] (단계별)}]` · `highlight`(단계 열) · `label_width`. 칩 텍스트가 `**`로 시작하면 강조 칩이 된다.

## 전면 슬라이드

- `cover`: `kicker`(사업명), `title`(`\n` 가능), `subtitle`, `org`, `author`, `date`
- `toc`: `items` (생략하면 `sections`)
- `divider`: `no`("01"), `title`, `items`(하위 목차)
- `statement`: `kicker`, `title`(한두 문장, `**강조**`), `subtitle` — 핵심 주장을 크게 보여주는 전면 슬라이드
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
| 핵심 요약 | `[numbered w2, kpi column]` |
| 환경 분석 | `[matrix letters:"SWOT"]` |
| 경쟁 포지셔닝 | `[matrix points w1.5, bullets]` |
| 시장 규모 | `[funnel, chart]` |
| 예산 | `[table total+merge, chart waterfall]` |
| 중간 점검 | `[progress, numbered]` |

실제 사용 예는 갤러리(https://codeblack-inc.github.io/mega-ppt/)에서 장표별 deck.json으로 볼 수 있다.

## 새 패널 추가
`scripts/build_deck.py`에 `def p_name(s, box, p)` 작성 → `PANELS`에 등록 → 이 문서 카탈로그에 추가 → `examples/sample.json` 또는 `gallery/showcase.json`에 사용 예 추가 (테스트가 두 파일을 합쳐 모든 패널이 쓰였는지 검사).
