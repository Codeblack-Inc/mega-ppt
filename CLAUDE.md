# mega-ppt

Claude Code + Codex 플러그인. 공용 스킬은 `skills/mega-ppt/` 하나 — 두 호스트가 같은 파일을 쓴다.

- 패널 추가: `scripts/build_deck.py`의 `PANELS`에 `p_name(s, box, spec)` 등록 → `references/layouts.md` 카탈로그 → `examples/sample.json`에 사용 예 (테스트가 모든 패널/전면 레이아웃이 샘플에 쓰였는지 검사)
- 품질 기준은 `references/density.md` (실제 한국 사업계획서 206장 분석)
- 테스트: `uv run tests/test_build.py`
- 시각 확인: `uv run skills/mega-ppt/scripts/build_deck.py skills/mega-ppt/examples/sample.json -o /tmp/s.pptx && uv run skills/mega-ppt/scripts/render.py /tmp/s.pptx`
- 버전 올릴 때 `.claude-plugin/plugin.json`, `.codex-plugin/plugin.json`, SKILL.md `metadata.version` 함께 수정
- 갤러리: `uv run gallery/build.py` → `site/` (sample.json × 테마 4종, LibreOffice 필요). main에 push하면 `.github/workflows/pages.yml`이 Pretendard를 설치하고 빌드·배포. 슬라이드별 `gallery` 키(name/desc)는 엔진이 무시하고 갤러리만 사용
- 예시 콘텐츠는 가상의 LLM 연구 과제만 쓴다 (실제 고객·과제 자료 금지)
- 목업 이미지: `gallery/mocks/*.html` → Chrome 헤드리스로 캡처해 `skills/mega-ppt/examples/assets/`에 커밋 (CI는 재생성하지 않음)
- 아이콘: `skills/mega-ppt/assets/icons.json` (lucide-static에서 추출, ISC)
