# mega-ppt

Claude Code + Codex 플러그인. 공용 스킬은 `skills/mega-ppt/` 하나 — 두 호스트가 같은 파일을 쓴다.

- 레이아웃 추가: `scripts/build_deck.py`의 `LAYOUTS`에 함수 등록 → `references/layouts.md` 문서화 → `examples/sample.json`에 예시 추가 (테스트가 sample.json이 모든 레이아웃을 쓰는지 검사함)
- 테스트: `uv run tests/test_build.py`
- 시각 확인: `uv run skills/mega-ppt/scripts/build_deck.py skills/mega-ppt/examples/sample.json -o /tmp/s.pptx && uv run skills/mega-ppt/scripts/render.py /tmp/s.pptx`
- 버전 올릴 때 `.claude-plugin/plugin.json`, `.codex-plugin/plugin.json`, SKILL.md `metadata.version` 함께 수정
