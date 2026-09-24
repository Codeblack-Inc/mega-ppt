---
name: mega-ppt
description: 전문적인 프레젠테이션(.pptx)을 생성한다. 사용자가 PPT, 슬라이드, 발표자료, 장표, 덱, 사업계획서, 제안서, 과제 발표자료, 보고서 슬라이드를 만들어 달라고 하거나 문서·데이터를 발표자료로 바꿔 달라고 할 때 사용. 스토리라인 설계 → 슬라이드마다 패널(표·차트·카드·프로세스·간트·구성도 등) 조합 → deck.json 작성 → python-pptx로 편집 가능한 PowerPoint 생성 → 렌더링 이미지로 시각 검수. 한국어 우선.
license: MIT
metadata:
  version: "0.4.1"
---

# mega-ppt

한국 사업계획서·과제 발표 수준의 **밀도 높은** 장표를 만든다. 슬라이드마다 결론 문장 헤드라인 하나를 두고, 그 결론을 증명하는 패널 2~4개를 격자로 조합한다.

이 파일이 있는 디렉터리를 `$SKILL`이라 한다. 스크립트는 `uv run`으로 실행한다(의존성 자동 설치). `uv`가 없으면 `pip install python-pptx pymupdf` 후 `python3`로 실행한다.

## 워크플로

### 1. 브리핑 (모르면 묻는다, 최대 3개)
- **청중과 목적**: 심사위원, 경영진, 고객 중 누구에게 무엇을 설득하려는가
- **분량과 밀도**: 장수, 그리고 발표용(간결)인지 제출용(상세)인지 → [`references/density.md`](references/density.md)
- **소스**: 첨부 문서와 데이터를 먼저 읽는다. 숫자는 소스에서만 가져오고 지어내지 않는다. 없는 수치는 `[확인 필요]`로 둔다.

### 2. 스토리라인 → [`references/storyline.md`](references/storyline.md)
슬라이드를 만들기 전에 **헤드라인만** 순서대로 나열해 사용자에게 보여주고 확인을 받는다. 헤드라인만 이어 읽어도 논리가 이어져야 한다.

### 3. 패널 설계 → [`references/layouts.md`](references/layouts.md)
슬라이드마다 "이 헤드라인을 증명하려면 무엇을 보여야 하나?"를 묻고 패널을 고른다.
- 숫자 추세는 `chart`, 다차원 비교는 `table`, 병렬 항목은 `cards`, 순서는 `process`, 일정은 `timeline`·`roadmap`, 시스템 구조는 `stack`, 조직은 `tree`, 전략 체계는 `house`, 환경 분석은 `matrix`, 시장 규모는 `funnel`을 쓴다. 전체 32종은 layouts.md에 있다.
- 같은 body 구성이 3장 연속되지 않게 한다.
- 강조(`highlight`, `**강조**`)는 슬라이드당 1~2곳만 쓴다.
- 사용자가 준 스크린샷·사진은 적극적으로 쓴다: 앱 화면은 `image`에 `frame`과 `callouts`, 여러 장은 `images`, 표지·구분 슬라이드는 `image` 배경, 전면 사진은 `photo`.
- 아이콘(`icons`, `cards`·`diagram`의 `icon`)은 Lucide 이름을 쓴다. 이름이 없으면 경고가 나오니 `assets/icons.json`에서 찾는다.

### 4. deck.json 작성 → 빌드
```bash
uv run $SKILL/scripts/build_deck.py deck.json -o deck.pptx
```
예시: [`examples/sample.json`](examples/sample.json)(모든 패널 사용). 테마: [`references/style-guide.md`](references/style-guide.md).
**경고(WARN)가 0개가 될 때까지** 고친다. 넘치면 문장을 줄이고, 비면 내용을 보강하거나 행 높이를 줄인다.

### 5. 시각 검수 (생략 금지)
```bash
uv run $SKILL/scripts/render.py deck.pptx
```
생성된 `deck/slide-NN.png`를 **모두 직접 열어 보고** 다음을 확인한다.
- 텍스트 잘림, 겹침, 단어 중간 줄바꿈
- 패널 사이 빈 공간과 높이 불균형
- 차트와 표가 읽히는 크기인가

문제가 있으면 deck.json을 고치고 다시 빌드·검수한다. LibreOffice가 없으면 사용자에게 알리고 이 단계만 건너뛴다.

## 원칙
- **헤드라인은 결론 문장**으로 쓴다. "시장 현황"(X), "시장은 연 12% 성장하지만 점유율은 정체"(O).
- **본문은 헤드라인을 증명**하는 역할만 한다. 증명하지 않는 패널은 뺀다.
- **숫자에는 출처와 단위**를 단다.
- **이미지는 지어내지 않는다.** 스크린샷이나 사진이 필요하면 `image` 자리표시자를 두고, 무엇을 넣으면 되는지 사용자에게 알려준다.
- **발표자 노트**: 장표에 못 넣은 설명은 `notes`에 넣는다.
