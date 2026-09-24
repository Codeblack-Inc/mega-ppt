# 디자인 시스템

16:9 (13.333 × 7.5 in), 흰 배경, 에디토리얼 스타일. 그림자·그라데이션·3D 없음.

## 토큰 (기본값 — `build_deck.py`의 `THEME`)

| 토큰 | 기본값 | 용도 |
|---|---|---|
| `font` | Pretendard | 전체 서체 (라틴+한글) |
| `ink` | `#1B1F2A` | 본문, 섹션 배경 |
| `muted` | `#6B7280` | 보조 텍스트, 라벨 |
| `rule` | `#E5E7EB` | 구분선 |
| `soft` | `#F3F4F6` | 카드 배경 |
| `paper` | `#FFFFFF` | 배경 |
| `accent` | `#1E4FD8` | **강조 — 슬라이드당 1곳** |

deck.json의 `theme`으로 일부만 덮어쓴다:
```json
"theme": { "accent": "#E4572E", "font": "Noto Sans KR" }
```

브랜드 색을 받으면 `accent`만 바꾸는 것을 우선한다. 나머지는 중립색 유지.

## 폰트
Pretendard가 없는 PC에서는 PowerPoint가 대체 폰트로 표시한다. 설치: `brew install --cask font-pretendard` 또는 https://github.com/orioncactus/pretendard

## 타입 스케일
타이틀 26pt bold · 본문 15–18pt · KPI 숫자 44pt · 라벨 14pt · 출처/페이지 9pt
