# 디자인 시스템

16:9 (13.333 × 7.5 in), 흰 배경. 그림자·그라데이션·3D는 쓰지 않는다. 밀도 높은 한국식 장표를 전제로, 색은 줄이고 격자와 선으로 정돈한다.

## 토큰 (기본값 — `build_deck.py`의 `THEME`)

| 토큰 | 기본값 | 용도 |
|---|---|---|
| `font` | Pretendard | 전체 서체 (라틴+한글) |
| `ink` | `#14213D` | 헤드라인, 표 헤더, 라벨·구분 슬라이드 배경 |
| `text` | `#252C3A` | 본문 |
| `muted` | `#5B6475` | 리드, 보조 텍스트, 캡션 |
| `rule` | `#D5DAE3` | 구분선, 카드 테두리 |
| `soft` | `#F3F5F9` | 카드 헤더, 표 항목 열, KPI 배경 |
| `paper` | `#FFFFFF` | 배경 |
| `accent` | `#2152E0` | **강조: 슬라이드당 1~2곳** |
| `title_style` | `line` | 패널 제목 스타일: `line`(마커와 밑줄) 또는 `bar`(남색 띠) |

deck.json의 `theme`으로 일부만 덮어쓴다:
```json
"theme": { "accent": "#E4572E", "font": "Noto Sans KR" }
```

브랜드 색을 받으면 `accent`만 바꾸는 것을 우선한다. 나머지는 중립색 유지.

## 폰트
Pretendard가 없는 PC에서는 PowerPoint가 대체 폰트로 표시한다. 설치: `brew install --cask font-pretendard` 또는 https://github.com/orioncactus/pretendard

## 타입 스케일
헤드라인 20pt bold(최소 16) · 리드 11pt · 패널 제목 11.5pt bold · 본문 10.5pt(자동 축소 최소 8) · 표 9.5pt · 캡션·출처 8~9pt · KPI 숫자 최대 30pt

글자 크기를 직접 지정하지 않는다. 박스에 맞춰 엔진이 자동으로 줄인다. 최소 크기로도 넘치면 경고가 나오니, 그때는 내용을 줄인다.
