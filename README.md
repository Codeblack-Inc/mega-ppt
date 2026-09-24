<h1><img src="gallery/mega-ppt.svg" alt="mega-ppt" width="220" /></h1>

[장표 갤러리](https://codeblack-inc.github.io/mega-ppt/) · [mega 제품군](https://codeblack-inc.github.io/mega-bi/) · [브랜드 가이드와 로고](https://github.com/Codeblack-Inc/mega-bi)

AI가 스토리라인을 설계하고, 슬라이드마다 패널(표·차트·카드·간트·구성도 등)을 조합해 편집 가능한 `.pptx`를 만드는 Claude Code / Codex 플러그인.

mega-ppt는 [mega 오픈소스 제품군](https://codeblack-inc.github.io/mega-bi/)의 발표자료 도구다. 갤러리에는 공식 로고와 심벌을 사용하고, 기본 장표에는 [mega BI](https://github.com/Codeblack-Inc/mega-bi)의 잉크·페이퍼·coral 계열을 적용한다. 공식 coral은 `#F37055`이며, 장표의 작은 글자와 색면에는 대비를 위해 진한 coral `#B5452D`를 쓴다. 다른 색이 필요하면 `deck.json`의 `theme`에서 바꿀 수 있다.

- **스토리라인**: 피라미드 원칙 + 액션 타이틀
- **헤더 + 패널 격자**: 패널 35종(표·차트·콤보·폭포·KPI·카드·프로세스·간트·로드맵·SWOT·포지셔닝·벤·피라미드·퍼널·조직도·전략 체계도·BMC·아키텍처·스윔레인·타일맵·체크리스트·아이콘·이미지·스토리보드·전후 비교·기기 구성 등)을 중첩 격자로 조합한다
- **보고서체 디자인**: 핵심 요약 상자, 잉크 탭 제목, □○- 불릿 위계, 테두리 격자처럼 실제 사업계획서에서 쓰는 형식을 따른다. 글자는 칸에 맞춰 커지고 옆 칸과 같은 크기로 맞춰져서 장표가 비어 보이지 않는다
- **한국식 장표 밀도**: 사업계획서·공공기관 보고서 수준(장당 300~600자)을 기준으로, 넘침·빈 칸·강조 과다를 자동으로 감지해 경고한다
- **이미지 장표**: 반분할(split)·이미지 배너·전면 사진 장표, 확대 삽입(zoom)·지시선 주석·번호 범례, 모자이크·흐름 화면 모음, 시연 스토리보드, 전후 비교, 노트북+휴대폰 구성, 영역 지정 자르기(focus)
- **디자인 시스템**: mega 기본 테마와 추가 테마 3종, Pretendard, 테마 토큰으로 브랜드를 적용한다
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
  examples/sample.json   예시 덱 (gallery/showcase.json과 합쳐 모든 패널·레이아웃 사용)
tests/test_build.py      스모크 테스트
```

## License
MIT · 아이콘: [Lucide](https://lucide.dev) (ISC, `skills/mega-ppt/assets/LICENSE-lucide`)
