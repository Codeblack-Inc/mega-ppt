# Mega PPT

AI가 스토리라인을 설계하고, 레이아웃 카탈로그에서 슬라이드별 레이아웃을 골라, 편집 가능한 `.pptx`를 만드는 Claude Code / Codex 플러그인.

- **스토리라인**: 피라미드 원칙 + 액션 타이틀
- **레이아웃 카탈로그**: cover · section · bullets · two-column · kpi · chart · process · table · quote · closing
- **디자인 시스템**: 중립색 + 강조색 1개, Pretendard, 테마 토큰으로 브랜드 적용
- **차트**: 네이티브 PowerPoint 차트 (bar / hbar / stacked / line / pie) — 열어서 데이터 수정 가능
- **시각 검수**: LibreOffice로 렌더링해 AI가 직접 보고 고침

## 설치

**Claude Code**
```text
/plugin marketplace add dr-coton/mega-ppt
/plugin install mega-ppt@mega-ppt
```

**Codex**
```bash
codex plugin marketplace add dr-coton/mega-ppt
codex plugin add mega-ppt@mega-ppt
```

**로컬 개발(심볼릭 링크)**
```bash
ln -s "$PWD/skills/mega-ppt" ~/.claude/skills/mega-ppt
ln -s "$PWD/skills/mega-ppt" ~/.codex/skills/mega-ppt
```

필요 도구: [`uv`](https://docs.astral.sh/uv/) (의존성 자동 설치), 시각 검수용 LibreOffice (`brew install --cask libreoffice`), 권장 폰트 Pretendard.

## 사용
> 이 보고서로 경영진 보고용 PPT 10장 만들어줘

직접 빌드:
```bash
uv run skills/mega-ppt/scripts/build_deck.py skills/mega-ppt/examples/sample.json -o sample.pptx
uv run skills/mega-ppt/scripts/render.py sample.pptx
```

## 구조
```
.claude-plugin/          Claude Code 플러그인 + 마켓플레이스 매니페스트
.codex-plugin/           Codex 플러그인 매니페스트
.agents/plugins/         Codex 마켓플레이스
skills/mega-ppt/
  SKILL.md               워크플로 (브리핑 → 스토리라인 → 레이아웃 → 빌드 → 검수)
  references/            storyline.md · layouts.md · style-guide.md
  scripts/build_deck.py  deck.json → .pptx (레이아웃 엔진)
  scripts/render.py      .pptx → slide PNG
  examples/sample.json   전 레이아웃 예시 덱
tests/test_build.py      스모크 테스트
```

## License
MIT
