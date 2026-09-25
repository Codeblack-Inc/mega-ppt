# 레이아웃 = 헤더 + 패널 격자

슬라이드는 두 종류다.

- **전면 슬라이드** (`layout`): `cover` · `toc` · `divider` · `statement` · `photo` · `split` · `closing`
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
- `banner: {src, focus?, h?, w?, caption?, alt?}` → 상단이 이미지 띠(높이 `h` 기본 1.9in)가 되고, 헤드라인은 왼쪽 잉크 블록(폭 `w` 기본 0.58×슬라이드) 위에 흰 글씨로 올라간다. 반투명 막 없이 이미지는 오른쪽 나머지 칸에만 그린다. 탭 내비게이션은 생략하고 `lead`는 띠 아래 요약 상자로 붙는다. 시설·현장 소개, 섹션 첫 장에 한두 번만.

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
| 영역 안에 여러 패널 중첩 (고밀도 종합 장표) | `group` |
| 핵심 기능·역량·가치 나열 (아이콘) | `icons` |
| 역량·영역의 교집합 | `venn` |
| 비즈니스 모델 | `canvas` |
| 점검표, 자체 평가표(배점·별점) | `checklist` |
| 시·도별 분포 | `map` |
| 시스템 아키텍처, 데이터 흐름, 스윔레인 | `diagram` |
| 스크린샷·사진 여러 장 | `images` |
| 시연 시나리오, 사용자 여정 (화면 3~5단계) | `storyboard` |
| 개선 전·후 화면 + 지표 비교 | `compare` |
| 웹 + 모바일 동시 제공 (노트북·휴대폰 구성) | `devices` |

### `bullets`
`items`: 문자열 배열 · `style`: `"plain"`이면 불릿 없는 문단 · `size` · `gap`
- 항목 전체가 `**…**`이면 □ 소제목, 보통 항목은 ○, `"- "`는 - 하위 항목, `"※"`로 시작하면 회색 주석
- 글자는 칸을 채우도록 커지고(너비 4in 이상 13pt, 미만 12pt까지) 남는 높이는 ○·□ 묶음 사이에만 나눈다. `- ` 하위 항목은 상위 항목에 붙어 있다. 묶음이 하나뿐이면 칸 가운데에 모은다
- `cols: 2` → 가운데에 가장 가까운 □ 소제목에서 두 단으로 나누고, 두 단은 같은 글자 크기를 쓴다

### `callout`
`text` · `label`(왼쪽 남색 라벨 칸, 기본 `"시사점"`, `""`이면 없음) · `style` · `align` · `size`
- 기본: 흰 바탕 + 1.25pt 남색 테두리 + 라벨 칸, 굵은 본문. `**강조**`만 강조색
- `"dark"`: 남색 판에 `⇒`로 시작하는 흰 글자. 덱에 한 번만 쓴다
- `"soft"`: 연회색 바탕, 테두리·라벨 없는 보통 굵기 곁말
- 행 높이는 글자에 맞춰 0.62~1.10in로 자동 제한된다

### `table`
`columns`(헤더, 생략 가능) · `rows`(2차원) · `widths`(열 비율) · `rowhead`(첫 열을 항목 열로: soft 바탕, 굵게, 왼쪽 정렬) · `highlight`(데이터 행 인덱스: 옅은 강조 바탕 + 굵게, 강조색은 그 행의 `**` 셀에만) · `size`(기본 10.5) · `stretch`(기본 true) · `unit`(제목 줄 오른쪽 "(단위: …)", 제목이 없으면 표 위)
- 글자는 칸 높이에 맞춰 12pt까지 커지고(단어가 칸 안에서 끊기지 않는 선까지), 모자라면 8pt까지 줄어든다. 행은 패널 높이까지 늘어나되 평균 0.65in를 넘지 않는다. 표가 패널의 80%를 못 채우면 경고가 난다 → 행·열을 늘리거나 callout 행을 붙이거나 행 `h`를 줄인다
- 헤더는 head 바탕에 굵은 ink 글자, 아래 0.75pt ink 선. 제목(탭)이 있으면 헤더가 탭 밑줄에 바로 붙고, 없으면 위에 1.5pt ink 선을 긋는다. `header: "dark"`이면 예전처럼 ink 바탕에 흰 글자
- 숫자 열(모든 셀이 `12`, `3.5%`, `+10%p`, `1.2×`, `48억`, `5.1만 건`, `42곳`, `-` 같은 값)은 헤더까지 오른쪽 정렬. 14자를 넘는 텍스트 열은 왼쪽, 짧은 텍스트 열과 기호는 가운데
- `total: true` → 마지막 행을 합계 행으로(head 바탕, 굵게, 윗선)
- `highlight_col`: 열 인덱스. 그 열의 헤더는 강조색, 본문은 옅은 강조 바탕(합계 행은 head 그대로). 기호는 다른 열처럼 ink (기능 비교표에서 '본 과제' 열)
- `heat: [열 인덱스]` → 그 열의 `상`·`중`·`하` 셀을 진한 순서로 음영하고 흰 틈으로 칸마다 떼어 놓는다 (위험 대장의 가능성·영향도·잔여 위험)
- `merge: [[r1, c1, r2, c2], ...]` → 셀 병합. 좌표는 헤더 행을 0으로 센다. 병합으로 가려지는 셀은 `""`로 둔다. 세로 병합 셀은 위쪽 정렬(`rowhead` 열은 가운데)
- 셀 값이 `● ◐ ○ ✓ ✗ △ O X`이면 기호로 크게 표시한다. ●✓O는 ink, 나머지는 회색 (기능 비교표). `"**●**"`처럼 감싸면 그 기호 하나만 강조색 (차별 항목 하나)
키-값 개요 표는 `columns` 없이 `rowhead: true`.

### `chart`
`kind`: `bar | hbar | stacked | line | pie | waterfall | combo` · `categories` · `series: [{name, values}]` · `number_format`(예: `"0%"`, `"#,##0"`) · `unit`(제목 줄 오른쪽 "(단위: …)", 제목이 없으면 차트 위) · `labels`(기본 true = 값 레이블 표시, 값 축 숨김) · `min`/`max`(값 축 범위)
- 색: 시리즈 순서대로 ink → 회색 3단(단일 시리즈 막대는 ink2). 강조색은 기본값이 아니다. `highlight`: 단일 시리즈 막대는 강조할 카테고리 인덱스(나머지는 회색), 다중 시리즈와 선은 강조할 시리즈 인덱스. 범위를 벗어나면 경고하고 무시
- 막대: 값 레이블은 막대 끝 밖에 굵게. 시리즈가 여럿이면 범례를 위에 둔다(`hbar`는 시리즈 순서대로 쓴 범례를 차트 위 오른쪽에 직접 그린다)
- `line`: 마지막 점에만 표식과 값 레이블. 시리즈가 2~3개이고 이름이 6자 이하면 이름도 선 끝에 붙이고 범례를 없앤다(아니면 범례). 단일 시리즈는 모든 점에 값. `min`이 없으면 값 축을 데이터 아래쪽 둥근 값에서 시작하고(음수도 포함), 카테고리 축은 늘 맨 아래에 둔다
- `pie`: 구멍 60% 도넛. 조각마다 항목명과 비율을 붙이고 범례는 없다. `highlight` 조각만 강조색. `center`: 구멍 가운데 글자(예: `"5만 건"`)
- `waterfall`: 단일 시리즈의 증감값. `totals: [인덱스]`는 누계 막대(ink)로 그린다. 음수는 감소(연회색, "−n" 레이블). 레이블은 `number_format`을 따른다(`"0.0"`, `"#,##0"`, `"0.0%"`, `'0"억"'`). `highlight`: 강조할 막대 인덱스(예: 헤드라인이 말하는 비목)
- `combo`: 막대 시리즈 + `"type": "line"` 시리즈(보조축). 막대는 아래 절반, 선은 위 띠에 따로 그려 레이블이 겹치지 않는다. 시리즈별 `number_format`. 값이 없는 칸은 `null`. 선 시리즈가 없으면 보통 막대로 그린다. `highlight`: 막대 카테고리 인덱스, 또는 `"line"`이면 선을 강조색으로

### `kpi`
`items: [{value, label, sub?}]` · `highlight` · `direction`: `"column"`이면 세로 배치
- 가로: 한 줄로 붙은 띠. 위 라벨 줄(연회색) + 큰 숫자 + `sub`(기준→목표, 기준연도 등 회색 한 줄)
- 숫자 뒤 3자 이하 단위(`%`, `%p`, `곳`, `억`, `만`, `배`)는 절반 크기로 붙는다. 숫자 크기는 모든 항목이 같다
- `"column"`: 행마다 왼쪽 라벨 칸(40%) + 오른쪽 숫자
- 가로 띠 행은 1.45in(+`sub` 0.3, +제목 0.4)까지만 커지고 남는 높이는 다른 행으로 간다

### `cards`
`items: [{title, items[] | text, icon?, value?, value_label?, note?}]` · `cols`(기본 = 항목 수; 2면 2×N 격자) · `numbered`(기본 true, 남색 번호 칸) · `highlight` · `size`
- 칸마다 연회색 머리 띠(번호 칸 또는 `icon`) + 본문. 본문 글자 크기는 모든 카드가 같다
- `value` + `value_label` → 카드 아래 수치 띠(예: `"30곳"` / `"공공·교육 기관 도입 (2028년)"`). `value`는 숫자+단위만 짧게, 설명은 `value_label`로 (카드 폭의 55%를 넘으면 경고)
- `note` → 맨 아래 `→ 문장` 띠 (대상, 결론 등)
- `cols: 1`이면 제목이 왼쪽 라벨 칸이 되고 본문은 오른쪽에 온다 (파급 효과, SO·WO·ST·WT 전략)
- 강조 카드는 강조색 머리 띠·테두리와 옅은 바탕

### `process`
`steps: [{title, items[] | text, label?}]` · `highlight` · `label_head`(기본 `"목표"`)
짙은 셰브론 아래 단계마다 테두리 칸을 두고, 불릿 크기를 모든 칸이 함께 쓴다. `label`이 있으면 칸 아래 띠에 `"목표  …"`로 단계 목표를 적는다.

### `timeline`
`periods`(열: 월·분기·연차) · `tasks: [{name, span:[시작,끝] | [[a,b],[c,d]], group?, highlight?, note?, owner?, milestone?}]` (1부터 시작, 끝 포함) · `label`(첫 열 제목) · `label_width`
- `groups: [["1차년도", 4], ["2차년도", 4]]` → 기간 헤더 위에 연차 헤더 줄
- `owner` → 과제명 옆 담당 열(`owner_label` 기본 `"담당"`, `owner_width`)
- `milestone: k | [k, …] | [k, "M1 라벨"] | [[k, "M1 라벨"], …]` → k번째 기간 가운데 ◆ (라벨은 오른쪽, 다음 ◆·현재선에 닿으면 왼쪽). `span` 없이 마일스톤만 있는 행도 된다
- `today: 4.2` → 기간 번호(소수 가능) 위치에 붉은 점선과 `today_label`(기본 `"현재"`). 진행 중 보고서에만 쓰고, 착수 전 계획서에는 넣지 않는다
행 높이는 패널 높이를 과제 수로 나눈다(상한 없음). 그룹 행은 전체 폭 음영에 얇은 ink 요약 막대, 과제 막대는 ink2. 그룹 막대는 하위 과제 기간을 모두 덮게 잡는다.

### `cycle`
`center` · `nodes`(3~8개, 2개 이하는 오류 → `flow`를 쓴다. 문자열(`\n` 줄바꿈) 또는 `{title, text}`) · `highlight`
노드를 타원 위에 시계 방향으로 놓고 화살표로 잇는다. `{title, text}` 노드는 제목 아래 담당·주기 같은 설명 줄을 단다.

### `stack`
`layers: [{label, items[], highlight?}]` (위에서 아래로, 빈 `items`는 띠만) · `label_width`
항목은 `"제목\n설명"` 문자열이나 `{title, text}`. 모든 층의 항목이 한 글자 크기를 함께 쓰고, 설명 줄은 작고 연하게 그린다.

### `image`
화면·사진·도면 한 장. `src`(deck.json 기준 상대 경로) · `alt` · `caption`(→ `[그림 n] 캡션` 띠, 덱 전체 번호. `fig:false`면 번호 없음) · `border`(기본 true). 패널 `title`이 있으면 이미지는 탭 선에 바로 붙는다(테두리 패널과 윗선이 맞음).
- **좌표는 모두 원본 이미지 기준 0~1** (`x`, `y`, `region`) → `focus`로 어떻게 잘라도 같은 지점을 가리킨다.
- `focus: [fx, fy, fw, fh]` → 원본에서 보여줄 영역. 칸 비율에 맞춰 넓어지되 찌그러지지 않는다(1.4배 넘게 넓어지면 경고). 스크린샷의 빈 여백을 잘라 내는 기본 수단. 없으면 `crop`(`center` | `top`).
- `fit`: `cover`(기본, 칸을 채움 · 비율이 1.3배 넘게 다르면 칸을 줄이고 80% 미만이면 경고) | `contain`(원본 비율 그대로)
- `frame`: `browser`(상단 ink2 제목 띠, `screen`으로 띠 문구 지정) | `laptop` | `tablet` | `phone` → 기기 프레임은 연회색 그림 칸 위에 92% 크기로 놓인다(노트북은 칸 너비의 94%까지). 기기 화면 비율은 고정이라 `focus`가 1.15배 넘게 넓어지면 경고 → 영역을 화면 비율에 맞추거나 `focus`를 뺀다.
- `bleed: true` → 패널이 슬라이드 여백에 닿은 쪽을 가장자리까지 넓힌다(테두리 없음, 캡션·수치 카드는 여백 안).
- `callouts: [{x, y, text?, box?: [w, h]}]` → 번호 마커(잉크 사각형, 항상 화면 안). `box`는 (x, y)를 중심으로 한 영역을 강조색 2pt 사각형으로 표시(슬라이드당 1개 권장), 마커는 영역 오른쪽 위 바깥 → 화면을 벗어나면 영역 왼쪽·오른쪽 가운데, 왼쪽 위, 아래 모서리 순. `text`(`"제목\n설명"`)가 있으면 오른쪽에 **사양 범례**(No · `legend_title` 기본 "화면 구성")가 패널 높이를 채운다. 범례 행은 제목 + 설명 한 줄이 기준(행이 낮으면 두 줄은 넘친다). `legend_width`(기본 0.36) · `leaders: true` → 마커(영역이면 영역 오른쪽 변)에서 범례 행까지 꺾은 지시선. 선이 화면을 가로지르므로 마커는 행 사이 빈 줄에 둔다.
- `zooms: [{region: [fx, fy, fw, fh], text, highlight?}]` (1~3개) → 오른쪽 열(`legend_width`)에 영역을 확대해 쌓고 연결선으로 잇는다. 본 이미지에 보이는 부분만 표시하고, 보이지 않는 영역은 경고. 각 확대 화면 아래 번호·제목·설명. `highlight`인 확대만 강조색. 영역의 오른쪽 변을 화면 끝에 맞추면 연결선이 화면을 가로지르지 않는다.
- `labels: [{x, y, text, side?}]` → 이미지 양옆 열(`label_width` 기본 0.28)에 지시선 주석. `side: l | r`(`left`/`right`도 됨), 없으면 `x<0.5`가 왼쪽. 주석은 가리키는 지점 높이에 맞춰 놓이고, 한쪽 열이 높이의 85% 미만에 몰리면 위아래로 고르게 벌린다. 휴대폰 목업·프레임 없는 사진에 쓴다.
- `stats: [{value, label}]` (2~4개) → 이미지 하단에 걸친 수치 카드. `highlight`: 강조할 카드 번호(0부터).
`src`가 없거나 파일이 없으면 자리표시자를 그린다.

### `images`
스크린샷·사진 여러 장. `items: [{src, focus?, crop?, caption?, tag?, text?, alt?}]` · `layout` · `gap` · `crop` · `highlight`(태그 강조 번호)
- `layout`: `grid`(기본, `cols` 기본 최대 3) | `feature`(`feature` 번 항목을 크게, 나머지는 옆 열 · `hero_w` 기본 0.6 · `hero_side`: `left`(기본) | `right` → 나머지를 먼저 읽고 대표 화면으로 끝낼 때) | `mosaic`(feature와 같은 배치에 간격 0.08, 캡션은 이미지 아래 잉크 띠) | `flow`(한 줄로 놓고 사이에 화살표, `tag` 기본 "n단계"). 항목이 1개면 grid로 그린다.
- `caption_style`: `strip`(이미지 아래 태그 칸 + `[그림 n]` 캡션 띠, 기본) | `overlay`(이미지 **아래** 잉크 띠에 태그·캡션·`text` — 띠를 먼저 재고 남은 높이에 그림을 넣어 `focus` 영역이 가려지지 않는다. 캡션은 한 줄에 맞춰 줄이고, 설명은 모든 띠가 같은 크기(10.5~11.5pt), 띠는 칸 높이의 절반까지. mosaic 기본). 작은 칸의 그림은 5:1 안팎으로 납작해지므로 `focus`도 그 비율로 잡는다.
- `text`(문자열 또는 불릿 배열) → strip이면 캡션 아래 설명 칸. 설명 칸은 글 높이에 맞추고(같은 줄은 가장 긴 글 기준), 이미지는 나머지 높이를 받되 원본 비율의 1.3배까지만 늘어난다. 그래도 0.25in 넘게 남으면 "images row is … taller" 경고 → 행 `h`를 낮추고 아래에 표·수치 행을 둔다. 칸이 모자라면 "images row too short" 경고.

### `storyboard`
`items: [{title, time?, src?, focus?, alt?, text: [불릿], metric?}]` (3~5개) · `highlight`(단계 번호, 0부터). 맨 위 레일에 번호 칸·단계명·시간 태그, 그 아래 화면 → 불릿 → 지표 띠가 단계마다 같은 줄에 정렬된다. 불릿은 네 단계가 한 크기·한 줄 간격을 쓰고(11→12.5pt, 줄바꿈이 늘지 않는 만큼만 키움), 화면은 설명이 남긴 높이를 모두 쓴다(최대 정사각형). 넓은 화면은 통째로 넣지 말고 `focus`로 핵심 부분만 거의 정사각형(가로:세로 1.0~1.1)으로 잘라 글자가 읽히게 한다. 행이 낮아 화면이 0.8in 미만이 되면 경고한다. 아래에 `callout` 행을 붙이면 좋다.

### `compare`
`before`·`after: {title, src?, focus?, alt?}` · `rows: [[지표, 전 값, 후 값, 변화?]]` · `delta`(`"−80%\n건당 검수 시간"`, 배지) · `highlight`(배지를 강조색으로). 왼쪽은 연한 머리띠, 오른쪽은 잉크 머리띠. 지표 행은 양쪽에 같은 줄로 놓이고, 네 번째 값(변화)은 가운데 홈에 들어간다. 배지는 두 화면 안쪽 모서리에 걸친다.

### `devices`
`laptop`·`phone: {src?, focus?, alt?, label?}` · `phone_side`(`right` 기본 | `left`) · `caption`(`[그림 n]` 캡션). 한 그림 칸(연회색 바탕) 위에 노트북과 휴대폰을 바닥선에 맞춰 겹쳐 놓고, `label`은 각 기기 아래에 적는다. `laptop`이나 `phone` 하나만 주면 그 기기가 그림 칸 가운데를 채운다. 옆 칸 머리와 맞추려면 `title`을 준다. 옆에 채널별 비교 `table`이나 `bullets`를 둔다.

### `label`
`text` · `highlight`. 남색 칸(강조 시 강조색)에 흰 글자, 세로로 긴 칸이면 글자를 세로로 쌓는다.

### `arrow`
`direction`: `"right"`(기본) | `"down"` · `text`(아래 방향일 때만)
- `text`가 있으면 화살표 대신 전체 너비 설명 띠 `▼ 문장`을 그린다 (예: 현황 → 한계의 대응 관계). 띠는 자기 행에 혼자 두거나 `w`를 준다 (고정 폭 칸이면 경고 후 화살표만)
- 행 높이는 0.34in로 자동 제한된다

### `matrix`
- 사분면: `quadrants: [{title, items[] | text}] ×4` (좌상·우상·좌하·우하 순) · `letters: "SWOT"`(머리 띠 왼쪽 남색 글자 칸) · `highlight`
  - `col_labels: ["긍정 요인", "부정 요인"]`, `row_labels: ["내부 역량", "외부 환경"]` → 위·왼쪽 축 띠 (왼쪽 띠는 글자를 세로로 쌓는다)
  - `letters`가 4자보다 짧으면 앞 칸에만 글자 칸을 단다
  - 네 칸의 본문 글자 크기는 같다. 제목에 영문을 겹쳐 쓰지 않는다 (`"강점"`, `"강점 Strength"` ✗)
- 포지셔닝: `points: [{label, x, y, note?, highlight?}]` (x·y는 0~1) · `quadrant_labels` ×4 · `highlight`(사분면)
  - `note` → 점 이름 아래 회색 수치 한 줄 (예: `"61.3% · 1.4×"`). 축 끝에 낮음/높음이 붙는다
  - 이름표는 가운데 구분선을 넘지 않는 쪽에 놓인다. 점을 구분선(0.5±0.03) 위에 두면 경고한다. 0·1은 축과 겹치지 않게 0.04~0.96으로 당긴다
- 공통: `x_label`, `y_label`(축 제목)

### `pyramid` / `funnel`
`levels: [{title, value?, text? | items[], tag?}]` (위에서 아래로) · `highlight`
사다리꼴 층이 이어져 하나의 피라미드(아래로 넓어짐)·퍼널(좁아짐)이 된다. `text`/`items`가 있으면 오른쪽 설명 열에 실선 지시선으로 잇고(불릿 크기 공유), `tag`는 맨 오른쪽 칸(추진 시점, 비중 등)에 넣는다. 태그 칸은 글자 크기(모든 층 공유, 12pt부터)에 맞춰 지시선 가운데에 놓이고, `" · "`에서 두 줄로 나뉜다. 태그만 있으면 도형이 나머지 폭을 쓴다.

### `tree`
`root: {name, role?, side?: [노드], side_left?: [노드], children: [{name, role?, footer?, children?: [str | {name, role}], items?}]}` · `highlight`(자식 인덱스) · `node_h`
3단까지 그린다. 3단은 2단 노드 아래 세로 목록이 되어 남은 높이를 채운다. `highlight`는 2단 머리 칸만 강조색으로 칠하고 그 아래 3단은 ink 테두리. `side`·`side_left`는 루트 오른쪽·왼쪽 점선 노드(자문위원회, PMO), `footer`는 세부과제 열 맨 아래 띠(참여 인력 등).

### `house`
`vision` · `goals[]`(선택) · `pillars: [{title, items[], note?}]` · `base: [str | {label, items[]}]` · `highlight`(기둥)
- 왼쪽 라벨 열: `vision_label`(기본 `"비전"`), `goals_label`(`"목표"`), `pillars_label`(`"추진 전략"`), `base_label`(`"추진 기반"`, 문자열 base일 때), `label_width`
- 기둥 `note` → 기둥 아래 띠 `"→ 성과지표"`. 기둥 본문 불릿은 크기를 공유한다
- `base`가 모두 문자열이면 라벨 칸 하나를 함께 쓰고, `{label, items}`가 섞이면 행마다 라벨 칸을 두며, `items`가 기둥 수와 같으면 기둥 아래에 맞춰 나눈다

### `milestones`
`events: [{date, title, text? | items[], value?, value_label?, highlight?}]` (4~7개 권장)
날짜 띠(ink) 아래 같은 폭의 칸을 이어 붙인 띠. 칸마다 제목·설명(불릿 크기와 문단 간격을 모든 칸이 공유해 줄이 가로로 맞는다)을 두고, `value`는 칸 아래에 굵게, `value_label`(예 `"채점 일치율"`)은 그 옆에 작게 적는다. 칸이 좁으니 항목은 한 줄(10자 안팎)로 쓴다. 패널 높이를 모두 채운다.

### `profiles`
`people: [{name, role, org?, meta?, photo?, items[], stats?}]` · `cols` · `highlight`
- 남색 이름 띠(이름 왼쪽, 역할 오른쪽) → 키-값 행 → 경력 불릿 → `stats` 한 줄
- `meta: ["소속: 가나다 AI 연구소 · 수석연구원", "학위: 박사 (전산학)"]` (없으면 `org`가 `소속` 행)
- `photo`: 증명사진 경로 (위 기준 자르기). 없으면 아무것도 그리지 않는다 (이니셜 아바타 없음)
- `stats`: `"참여율 50% · 논문 14 · 특허 3"`

### `numbered`
`items: [{title, text? | items[], value?, value_label?}]` · `cols`(기본 1) · `highlight`
- 행마다 남색 번호 칸(행 높이 전체) + 제목 + 근거 문장. 제목·본문 글자 크기는 모든 행이 같다
- 본문이 한 문장이면 제목과 본문을 한 덩어리로 행 가운데에 놓는다. 항목이 많아 행이 0.75in보다 낮으면 경고 (`cols: 2` 또는 행 `h`를 키운다)
- `value` + `value_label` → 오른쪽 1.4in 수치 열 (예: `"+3%p"` / `"파라미터 2배 효과"`)

### `flow`
`nodes: [{title, items[] | text}]` · `edges: [str | {label, back}]` (노드 사이, `back`이 있으면 역방향 화살표 추가) · `highlight`
- `zones: [{label, from, to, highlight?}]` → from~to 노드를 점선 경계(내부망, 외부 클러스터 등)로 묶고 왼쪽 위에 라벨 탭을 단다

### `progress`
`items: [{label, value, display?, sub?, target?}]` · `max`(기본 100) · `label_width` · `highlight`(인덱스 또는 인덱스 목록: 계획에 못 미친 항목) · `target`(모든 항목 공통 눈금) · `target_label`(첫 눈금 위에 한 번만 쓰는 글자, 기본 "목표")
- 행이 패널 높이를 나눠 채운다. 값 열은 가장 긴 `display`(예: `"4.8만 / 5만 건"`)에 맞춰 넓어진다. `sub`는 항목명 아래 보조 줄(예: `"4.8만 / 5만 건"`), `target`은 막대 위 ink 눈금

### `group`
`body`: 본문과 같은 격자(행 배열) · `gap`(기본 0.12). 패널 안에 패널을 넣는다. 제목(`title_style: "bar"`)을 단 group 3개를 나란히 두면 공공기관 보고서식 고밀도 장표가 된다. 제목에 목표를 함께 쓰면(`"데이터 구축  |  목표 5만 건"`) 아래 KPI·표가 목표 대비 실적으로 읽힌다. 제목 없는 group을 한쪽 열에 두면 표 + callout + 불릿처럼 서로 다른 패널을 세로로 쌓아 옆 패널과 아래선을 맞출 수 있다.

### `icons`
`items: [{icon, title, text? | items[]}]` · `cols` · `highlight`
- 칸마다 왼쪽 연회색 아이콘 열 + 제목 + 본문(○ 불릿 또는 문단). 가운데 정렬 카드는 쓰지 않는다
- 본문이 한 문장이면 제목과 본문을 한 덩어리로 칸 가운데에 놓는다
- 예전 `align`·`style`은 받아 두기만 하고 무시한다

아이콘 이름은 Lucide 2,118종(`assets/icons.json`의 키). 찾기: `grep -o '"[a-z0-9-]*shield[a-z0-9-]*"' $SKILL/assets/icons.json`. `cards` 항목에도 `icon`을 줄 수 있다.

### `venn`
`sets: [{title, items?}]` (2~3개) · `center`(교집합 라벨) · `highlight`(true → 교집합 칸 강조)
외곽선 원(ink·ink2·grey1)이며 채움은 없다. 집합 제목은 모두 ink, 항목 크기는 세 집합이 함께 쓴다.

### `canvas`
비즈니스 모델 캔버스 9블록: `partners` `activities` `resources` `value` `relationships` `channels` `segments` `costs` `revenue` (각각 불릿 배열) · `highlight`(블록 키 또는 목록 → 강조 테두리) · `widths`(5개 열 폭 비율, 예 `[1,1,1.25,1,1]`) · `split`(위 블록 높이 비율, 기본 0.66)
블록 본문은 글자 크기와 문단 간격을 아홉 블록이 함께 쓴다.

### `checklist`
- 점검표: `items: [{label, status: done|partial|todo|fail, status_label?, note?}]` → 상태 아이콘(완료 ink, 진행·미착수 회색 단계)과 비고
- 평가표: `items: [{label, score, max?(5), weight?, note?}]` → 점수를 `4 / 5`와 가는 막대로 표시. `weight`가 있으면 배점 열과 종합 점수 행(Σ 배점 × 점수 / 만점)을 자동 계산
- 공통: `columns`(헤더 덮어쓰기) · `widths` · `summary`(종합 행 문구 덮어쓰기) · `highlight`(강조할 항목 인덱스: 옅은 강조 바탕 + 강조색)
- 헤더·선은 `table`과 같고, 행이 패널 높이를 나눠 채운다. 항목이 적어 행이 너무 높으면 group으로 아래에 불릿·callout을 쌓는다

### `map`
`values: {시·도: 숫자}` (서울, 부산, 대구, 인천, 광주, 대전, 울산, 세종, 경기, 강원, 충북, 충남, 전북, 전남, 경북, 경남, 제주) · `unit` · `format`(기본 `"{:,}"`) · `highlight`(강조색 테두리로 강조할 시·도 목록, 슬라이드의 강조 하나로 친다). 17개 시·도 사각 타일맵에 값의 크기를 ink 농도(head → 회색 → ink2)로, 값은 타일 오른쪽 아래에 굵게 표시한다. 타일은 패널 높이를 6행으로 나누고 가로로 조금 넓게 채운다. 값 합계와 출처의 n이 다르면 출처에 제외 사유를 쓴다.

### `diagram`
격자(`cols` × `rows`) 위에 자유 배치:
- `zones: [{label, col, row, w, h, highlight?}]` — 점선 영역(내부망, 클라우드 등), 왼쪽 위 라벨 탭
- `nodes: [{id, label, sub?, icon?, col, row, w?, h?, height?, style: paper|accent|ink|soft}]` — 사각 노드, 라벨 크기는 모든 노드가 함께 쓴다. `icon`은 핵심 노드에만
- `edges: [{from, to, label?, style: "dashed"?, both?, highlight?}]` — 겹치는 행/열이면 직선, 아니면 ㄱ자로 자동 연결. 가로선이 라벨보다 짧으면 라벨을 그리지 않고 경고한다(노드 `sub`에 적거나 간격을 넓힌다)
- `lanes: [행위자...]` → 스윔레인. 레인마다 행 하나, `lane_width`
- `numbered: true` → 노드마다 왼쪽에 번호 칸(노드 순서)

### `roadmap`
`phases[]` · `tracks: [{name, cells: [str | [str]] (단계별)}]` · `highlight`(단계 열) · `highlight_chip: [트랙, 단계, 항목]`(강조색 칩 하나) · `label_width`
칩 높이와 줄 위치는 트랙마다 같아서(가장 많은 칸 기준) j번째 칩이 단계를 가로질러 맞는다. 칩이 적은 칸은 일찍 끝난다. `"**…**"` 칩은 짙은 칩, `"M3 | 내용"`은 `|` 앞을 왼쪽 월 칸에 넣는다. 단계 게이트(통과 기준)는 트랙 하나로 둔다.

## 전면 슬라이드

- `cover`: `kicker`(사업명, 제목 바로 위 테두리 태그), `title`(`\n` 가능, 34→46pt · 사진이 있으면 →40pt, 한 줄이어도 경고 없음), `subtitle`, `org`(부제 아래 기관 블록 `"주관기관: A · 공동기관: B · C"` — ` · `와 `:`로 나누고, 라벨이 없으면 '기관'), `image`(오른쪽 사진 칸), `crop`, `focus`. 하단 남색 띠에 덱 `footer`. 연구기간·연구비·제출일 같은 정보표는 두지 않는다
- `toc`: `title`(기본 "목 차"), `items` — 문자열 또는 `{title(또는 name), items}` (생략하면 `sections`). 문자열 항목은 같은 번호(`no`)의 divider `items`를 하위 목차로 가져온다. 로마 숫자 칸·하위 목차·점선 리더·쪽 범위. 쪽 범위는 섹션의 첫 연속 구간(해당 `no`의 divider 포함)에서 자동 계산. 하위 목차가 없으면 행 높이 0.95. 6개를 넘으면 2단
- `divider`: `no`("01" → Ⅰ, 왼쪽 남색 칸의 번호 칸 · 숫자가 아니면 그대로, "00"이면 생략), `title`, `items`(하위 목차, 제목 바로 아래). 사진이 없으면 오른쪽에 전체 섹션 목차(섹션명만, 현재 섹션 강조)와 쪽 번호. `image`: 칸과 비율이 비슷하면 오른쪽을 막 없이 채우고(`crop`, `focus`), 25% 넘게 다르면 자르지 않고 회색 그림 칸에 테두리와 함께 온전히 싣고 아래에 한 줄 섹션 띠(번호·섹션명·쪽 범위)를 둔다. `fit`(`cover`|`contain`)으로 강제, `caption`은 '[그림 n]' 캡션
- `statement`: `kicker`(이중 테두리에 걸친 탭, 기본 "핵심 메시지"), `title`(한두 문장, `**강조**`, 26→34pt), `subtitle`(짧은 괘선 아래) — 주장·괘선·부제를 한 덩어리로 테두리 안 세로 가운데에 놓는다. `points`(≤4, `{value, label}` 또는 문장 → '근거 n' 칸, 값 글자 크기는 칸끼리 같다) — 핵심 주장을 크게 보여주는 전면 슬라이드
- `photo`: `image`, `focus`, `kicker`, `title`, `subtitle`, `points: [{value, label}]`(최대 4, 핵심 수치 행), `caption`, `style`(`split` 기본: 잉크 단 + 사진 | `band`: 위 사진 + 아래 잉크 띠, 수치는 오른쪽), `align`(`left`|`right` 잉크 단 위치) — 전면 사진·화면 장표. 막(veil) 없이 사진은 남은 칸만 채우고, 제목(최대 2줄)·부제·수치는 실제 높이로 쌓아 가운데 정렬한다.
- `split`: 한쪽(`side` `left` 기본 | `right`, 폭 `ratio` 0.36~0.5, 기본 0.42)은 이미지, 다른 쪽은 일반 본문(`section`, `kicker`, `title`, `lead`, `body`, `source`)이 들어간다. `image`, `focus`, `caption`, `stats: [{value, label}]`. `mat`(`ink`|`soft`)을 주면 단을 색으로 칠하고 그 위에 `frame`(`browser`|`laptop`|`phone`, `screen`=브라우저 제목) 화면과 캡션, 수치 행을 쌓는다. 화면은 `focus` 비율 그대로 그리고, 세로 공간이 모자라면 잘라 넓히지 않고 폭을 줄여 가운데 둔다. `mat`이 없으면 이미지가 위·아래·바깥 끝까지 꽉 차고, 수치는 이미지 아래 잉크 띠에 들어간다. 밝은 가로 화면은 `mat`, 사진·어두운 화면은 `mat` 없이.
- `closing`: `title`(기본 "감사합니다"), `subtitle`(맺음 문장 한 줄). 하단 띠는 cover와 같다. 목표 수치·연락처 표는 두지 않는다

## 자주 쓰는 조합

| 목적 | body |
|---|---|
| 배경 → 문제 | `[label, cards]`, `{h:.25, [arrow down]}`, `[label, cards]` |
| 근거 3종 + 결론 | `{h:3, [chart, bullets, table]}`, `{h:.55, [callout]}` |
| 구성도 + 목표 | `[stack w1.3, table]` |
| 체계 + 역할 | `[cycle, cards cols2 w1.5]` |
| 일정 + 단계 목표 | `{h:3, [timeline]}`, `[process]` |
| 화면 + 기능 | `{h:3, [image frame:browser focus callouts leaders]}`, `{h:1.35, [table]}` |
| 효과 | `[kpi]`, `{h:1.7, [chart, cards cols1]}` |
| 비교 | `{h:.5, [callout label]}`, `[table rowhead merge]` (구분 · AS-IS · TO-BE, 정량 효과 행 강조) |
| 핵심 요약 | `[numbered w2, kpi column]` |
| 환경 분석 | `[matrix letters:"SWOT" row_labels col_labels w2, cards cols1]` (SO·WO·ST·WT 전략) |
| 경쟁 포지셔닝 | `[matrix points w1.5, group [bullets], [table]]` |
| 시장 규모 | `[funnel, chart]` |
| 예산 | `[table total+merge, chart waterfall]` |
| 중간 점검 | `[progress, numbered]` |
| 화면 설명 | `[image frame:laptop zooms]` · `[image bleed stats w1.65, bullets]` |
| 앱 소개 | `[image frame:phone labels]` |
| 작업 흐름 화면 | `{h:3.2, [images layout:flow]}`, `{h:1, [table rowhead]}`, `{h:.6, [callout]}` |
| 시연 화면 모음 | `[images layout:mosaic hero_side:right]` |
| 화면 + 논지 (반분할) | `layout:split mat:ink frame:browser stats` + `{h:2.4, [table]}`, `{h:1.7, [bullets]}` |
| 시설·현장 소개 | `banner` + `[kpi]`, `{h:2.2, [table w1.5, bullets]}` |
| 시연 시나리오 | `{h:3.9, [storyboard]}`, `{h:.6, [callout]}` |
| 개선 전후 | `{h:4, [compare delta]}`, `{h:.6, [callout]}` |
| 웹·앱 동시 제공 | `{h:4, [devices w1.7 title, table]}`, `{h:.6, [callout]}` |
| 아키텍처 | `[diagram zones+nodes+edges]` |
| 업무 프로세스 | `[diagram lanes]` |
| 고밀도 종합 | `[group bar, group bar, group bar]` (각 group에 kpi·chart·table) |

실제 사용 예는 갤러리(https://codeblack-inc.github.io/mega-ppt/)에서 장표별 deck.json으로 볼 수 있다.

## 새 패널 추가
`scripts/build_deck.py`에 `def p_name(s, box, p)` 작성 → `PANELS`에 등록 → 이 문서 카탈로그에 추가 → `examples/sample.json` 또는 `gallery/showcase.json`에 사용 예 추가 (테스트가 두 파일을 합쳐 모든 패널이 쓰였는지 검사).
