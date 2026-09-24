---
name: mega-ppt
description: 전문적인 프레젠테이션(.pptx)을 생성한다. 사용자가 PPT, 슬라이드, 발표자료, 덱, 보고서 장표, 제안서, 피치덱을 만들어 달라고 하거나 문서·데이터를 발표자료로 바꿔 달라고 할 때 사용. 스토리라인 설계 → 슬라이드별 레이아웃 선택 → deck.json 작성 → python-pptx로 편집 가능한 PowerPoint 생성 → 렌더링 이미지로 시각 검수. 한국어 우선.
license: MIT
metadata:
  version: "0.1"
---

# Mega PPT

컨설팅급 PPT를 만든다. 핵심은 **내용이 레이아웃을 고르게 하는 것** — 슬라이드마다 메시지 하나, 그 메시지를 가장 잘 보여주는 레이아웃 하나.

이 파일이 있는 디렉터리를 `$SKILL`이라 한다. 스크립트는 `uv run`으로 실행한다(의존성 자동 설치). `uv`가 없으면 `pip install python-pptx pymupdf` 후 `python3`로 실행.

## 워크플로

### 1. 브리핑 (모르면 묻는다, 최대 3개)
- **청중과 목적**: 누구에게, 무엇을 결정/이해시키려는가
- **분량**: 기본 8–12장
- **소스**: 첨부 문서·데이터가 있으면 먼저 읽는다. 숫자는 소스에서만 가져온다. 지어내지 않는다.

### 2. 스토리라인 → [`references/storyline.md`](references/storyline.md)
슬라이드를 만들기 전에 **액션 타이틀만** 순서대로 나열해 사용자에게 보여준다. 타이틀만 읽어도 전체 논리가 이어져야 한다. 확인 받고 다음 단계.

### 3. 레이아웃 매칭 → [`references/layouts.md`](references/layouts.md)
각 슬라이드 메시지의 형태(숫자 강조? 비교? 순서? 추세?)를 보고 레이아웃을 고른다. 같은 레이아웃이 3장 연속이면 다시 고민한다.

### 4. deck.json 작성 → 빌드
```bash
uv run $SKILL/scripts/build_deck.py deck.json -o deck.pptx
```
예시: [`examples/sample.json`](examples/sample.json). 테마 커스터마이즈: [`references/style-guide.md`](references/style-guide.md).

### 5. 시각 검수 (생략 금지)
```bash
uv run $SKILL/scripts/render.py deck.pptx
```
생성된 `deck/slide-NN.png`를 **모두 직접 열어 보고** 확인:
- 텍스트 넘침·잘림·겹침 → 문장을 줄인다 (폰트를 줄이지 않는다)
- 강조색이 슬라이드당 1곳인가
- 타이틀이 2줄을 넘지 않는가

문제가 있으면 deck.json 수정 → 재빌드 → 재검수. LibreOffice가 없으면 사용자에게 알리고 이 단계만 건너뛴다.

## 원칙
- **액션 타이틀**: "시장 현황"(X) → "시장은 연 12% 성장하지만 점유율은 정체"(O). 결론을 문장으로.
- **강조는 하나**: `highlight`는 독자가 가장 먼저 봐야 할 1곳에만.
- **숫자에는 출처**: 데이터 슬라이드는 `source` 필수.
- **덜어내기**: 불릿 5개 초과, 한 불릿 2줄 초과면 쪼개거나 줄인다.
- **발표자 노트**: 슬라이드에 못 넣은 설명은 `notes`로.
