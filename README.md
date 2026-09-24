# Mega PPT

**[장표 갤러리 보기 →](https://codeblack-inc.github.io/mega-ppt/)**

AI가 스토리라인을 설계하고, 슬라이드마다 패널(표·차트·카드·간트·구성도 등)을 조합해 편집 가능한 `.pptx`를 만드는 Claude Code / Codex 플러그인.

- **스토리라인**: 피라미드 원칙 + 액션 타이틀
- **헤더 + 패널 격자**: 슬라이드마다 패널 24종(표·차트·KPI·카드·프로세스·간트·순환도·구성도·SWOT·포지셔닝·피라미드·퍼널·조직도·전략 체계도·마일스톤·프로필·로드맵·흐름도 등)을 조합하므로 레이아웃이 고정되지 않는다
- **한국식 장표 밀도**: 사업계획서·공공기관 보고서 수준(장당 300~800자)을 기준으로, 넘침과 빈 공간을 자동으로 감지해 경고한다
- **디자인 시스템**: 중립색과 강조색 1개, Pretendard, 테마 토큰으로 브랜드를 적용한다
- **편집 가능**: 표와 차트가 모두 PowerPoint 네이티브라서 열어서 바로 수정할 수 있다
- **시각 검수**: LibreOffice로 렌더링한 이미지를 AI가 직접 보고 고친다

## 설치

**Claude Code**
```text
/plugin marketplace add Codeblack-Inc/mega-ppt
/plugin install mega-ppt@mega-ppt
```

**Codex**
```bash
codex plugin marketplace add Codeblack-Inc/mega-ppt
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
  references/            storyline.md · layouts.md(패널 카탈로그) · density.md · style-guide.md
  scripts/build_deck.py  deck.json → .pptx (레이아웃 엔진)
  scripts/render.py      .pptx → slide PNG
  examples/sample.json   모든 패널을 쓰는 12장 예시 덱
tests/test_build.py      스모크 테스트
```

## License
MIT
