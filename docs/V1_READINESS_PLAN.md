# v1.0 Readiness + v1.1 Roadmap Plan

> 작성일: 2026-04-15 (v0.0.26 기준)
> 목적: 코드베이스 전수 감사 결과 남은 블로커 정리 + v1.0 release + v1.1 로드맵

---

## 현재 상태 요약

### 이미 완료 (v0.0.11 ~ v0.0.26, 16개 릴리즈)

✅ 기초 — Fields/Bookmarks/Hyperlinks accessor, charshape/parashape property
✅ 18개 accessor 시스템 + 8개 context manager
✅ 11개 preset (승승아빠 매크로 이식)
✅ P0 버그 fix 5개 (replace_font / default_font / charshape 충돌 / Run None / view.zoom)
✅ P1 버그 fix (find_text / set_charshape facename fan-out)
✅ silent fail 14곳 logger.debug 화 + Engine/open/save INFO
✅ BorderType / HatchStyle / Color 등 enum 통일
✅ 객체별 25개 API 문서 + sidebar 트리
✅ tests 1388/1388 통과 (23개 도메인 파일)

---

## v1.0 블로커 (릴리즈 전 필수)

### B-1: CHANGELOG 정리 + v1.0 announcement
**파일**: `CHANGELOG.md`
- `[Unreleased]` 섹션에 breaking / deprecated / migration 정리
- `[1.0.0]` 섹션 신설 — v0.0.11~0.0.26 의 큰 흐름 요약
- "Deprecation timeline" 섹션 — fields 10개는 v1.1 에서 제거 명시
- 모든 변경의 사용자 영향 설명

### B-2: `pyproject.toml` v1.0 업그레이드
**파일**: `pyproject.toml`
- `version = "1.0.0"`
- `Development Status :: 5 - Production/Stable` (현재 2 - Pre-Alpha)
- `[project.optional-dependencies]` 추가:
  ```toml
  [project.optional-dependencies]
  dev = ["pytest>=7.0", "ruff>=0.1", "PyMuPDF>=1.23", "pandas"]
  test = ["pytest>=7.0", "PyMuPDF>=1.23"]
  ```

### B-3: CI test workflow
**파일**: `.github/workflows/test.yaml` (신설)
- ubuntu-latest 에서 pytest mock 테스트만 실행 (1388개)
- `python -m pytest tests/ -v --ignore=tests/test_all_actions.py`
- smoke test 는 Windows self-hosted runner 있을 때만 (조건 처리)

### B-4: Migration guide 작성
**파일**: `docs/MIGRATION_v1.md` (신설)
- legacy 17개 → 신 API 대응표
- `set_charshape(facename=)` ↔ `facename_hangul/latin/...` 관계
- ENUM 네이밍 이동 (`BorderType=1` → `"solid"`)
- Color 상수 사용법
- 튜토리얼 9 의 mail merge 코드를 신 API 로 재작성 예시

### B-5: 튜토리얼 API 갱신 (7 튜토리얼, 43회 `set_charshape`)
**파일들**: `nbs/01_tutorials/02, 04, 05, 06, 07, 09, 12.ipynb`
- `set_charshape(bold=True, text_color="red")` 유지 (이건 정상 API)
- 하지만 `Color.RED` / `"yellow"` 같은 신 방법 추가 예시
- `app.charshape = {"bold": True}` property 스타일도 혼합
- `replace_font` 예제는 `old` 인자 명시 (이전 튜토리얼 12 버그 있음)

### B-6: 내부 helper 중복 정리
**파일**: `hwpapi/core/app.py:1625` 의 중첩 `_cell_addr` 함수 제거
- `presets/__init__.py` 의 `_cell_addr` 이미 `functions.cell_addr` 의 wrapper
- `core/app.py` 에서도 `from hwpapi.functions import cell_addr` 사용

### B-7: smoke test 확장 (미검증 8개 영역)
**파일**: `tests/smoke_features.py`
- `feature_19_styles`: app.styles.apply/export 검증
- `feature_20_controls_find`: app.controls.find 로 표 찾기
- `feature_21_images_grayscale`: app.images.grayscale_all
- `feature_22_convert_wrap`: app.convert.wrap_by_word / wrap_by_char
- `feature_23_preset_page_border`: app.preset.page_border
- `feature_24_template_roundtrip`: template.save → apply 검증
- `feature_25_lint_accuracy`: 알려진 문제 샘플로 lint 감지
- `feature_26_debug_tools`: debug.timing / debug.trace 실제 동작

### B-8: README / index.qmd v1.0 마감
**파일**: `README.md`, `nbs/index.qmd`
- README 상단에 `![PyPI](badge)` + `![Tests](badge)` + `![Docs](badge)`
- "v1.0 announcement" 섹션 — 주요 achievement 정리
- index.qmd 랜딩에 "v1.0 Ready" 표시

---

## v1.1 개선 로드맵 (v1.0 announce 후)

### R-1: App 클래스 분할 (103 methods → 5 mixins)
**파일**: `hwpapi/core/app.py` (3,272줄) 를 `hwpapi/core/ops/` 패키지로
- `_LifecycleMixin` (118-520): `__init__`, properties, `select_all`, clipboard
- `_EditingMixin` (536-710): insert_text, insert_table, insert_hyperlink 등
- `_NavigationMixin` (731-918): goto_page, highlight
- `_FormattingMixin` (2013-2494): charshape, parashape, find_text
- `_TableMixin` (3066+): set_cell_*, cell 관련

App 이 mixins 상속만 바꾸므로 public API 변경 없음 (non-breaking).

### R-2: HwpError 예외 계열
**파일**: `hwpapi/exceptions.py` (신설)
```python
class HwpError(Exception): ...
class HwpFieldNotFound(HwpError, KeyError): ...
class HwpFileLocked(HwpError): ...
class HwpDialogTimeout(HwpError): ...
class HwpVersionMismatch(HwpError): ...
```
- `field_exists=False` 인데 `fields["name"].value` 접근 시 `HwpFieldNotFound` raise
- 현재 `app.fields[name]` 은 `KeyError` raise — 이미 표준이지만 `HwpFieldNotFound(KeyError)` 로 감쌈

### R-3: `python -m hwpapi.migrate` CLI
**파일**: `hwpapi/migrate.py` (신설), `hwpapi/__main__.py` (신설)
- AST 변환으로 legacy API → 신 API 자동 변환
- `app.create_field("x")` → `app.fields.add("x")`
- `app.set_field("x", v)` → `app.fields["x"] = v`
- 17개 legacy 모두 변환 가능
- `--dry-run` / `--diff` / `--write` 옵션

### R-4: Context manager 네이밍 통일
**파일**: `hwpapi/core/app.py`
- `silenced` → 유지 (이미 Pythonic)
- `charshape_scope` → `using_charshape` (verb form, `using_X` 패턴)
- `parashape_scope` → `using_parashape`
- `use_document` → 유지 또는 `using_document`
- 기존 이름은 deprecated alias 로 유지

### R-5: 색상 유틸 단일화 — `hwpapi/color.py`
**파일**: `hwpapi/color.py` (신설), `hwpapi/functions.py` (4개 함수 이전)
- `convert_to_hwp_color`, `convert_hwp_color_to_hex`, `get_rgb_tuple`, `hex_to_rgb` 모두 `hwpapi.color` 로
- `from hwpapi import color as C; C.red()` 등 깔끔한 namespace

### R-6: Type hints + mypy
**파일**: `pyproject.toml`, `mypy.ini` (신설)
- `[tool.mypy]` 설정 — strict 모드, but `hwpapi.parametersets.sets.*` 는 ignore (auto-generated 성격)
- `hwpapi/core/*.py` 먼저 typed — 가장 공개적 API
- `from __future__ import annotations` 이미 쓰고 있음

### R-7: pre-commit + ruff 설정
**파일**: `.pre-commit-config.yaml` (신설), `pyproject.toml [tool.ruff]`
- ruff — formatter + linter (black + flake8 대체)
- pre-commit — git commit 전 자동 실행
- 기존 코드 style 존중 (광범위한 format 변경 피함)

### R-8: Windows self-hosted CI runner (smoke test 자동화)
**파일**: `.github/workflows/smoke.yaml` (신설)
- Windows + HWP 가 설치된 self-hosted runner 에서
- PR 시 smoke_features.py / smoke_scenarios.py 자동 실행
- PDF 차이 감지 (golden snapshot 비교)
- 시각 regression 잡는 유일한 방법

---

## 구현 순서 (권장)

### Wave 1 (v1.0 블로커, 2~3 commit)
1. B-6 내부 helper 정리 (quickest)
2. B-2 pyproject.toml v1.0
3. B-1 CHANGELOG 정리
4. B-4 MIGRATION_v1.md
5. B-8 README / index v1.0 마감
6. B-3 CI test workflow

### Wave 2 (v1.0 실증, HWP 필요)
7. B-7 smoke test 확장 (8 시나리오)
8. B-5 튜토리얼 API 갱신
9. `python tests/generate_doc_artifacts.py` 전체 재생성 검증

### Wave 3 (v1.0 tag + PyPI)
10. `git tag v1.0.0 -m "v1.0 — Stable release"`
11. PyPI Trusted Publisher 자동 업로드
12. GitHub Release 공지

### Wave 4 (v1.1 개선, post-1.0)
13. R-1 App mixin 분할
14. R-2 HwpError
15. R-3 migrate CLI
16. R-4~R-8 infrastructure 업그레이드

---

## 검증 기준

각 Wave 완료 시:
- `python -m pytest tests/ -q --ignore=tests/test_all_actions.py --ignore=tests/test_all_parametersets.py` → 1388+ 통과
- `python tests/generate_doc_artifacts.py` → 13+개 demo 성공
- `quarto render nbs --no-execute` → 사이트 빌드 성공
- git push + tag

## v1.0 announce 체크리스트

- [ ] B-1 CHANGELOG 정리
- [ ] B-2 pyproject.toml v1.0.0 + dev extras
- [ ] B-3 .github/workflows/test.yaml
- [ ] B-4 docs/MIGRATION_v1.md
- [ ] B-5 튜토리얼 7개 갱신
- [ ] B-6 `_cell_addr` 중첩 제거
- [ ] B-7 smoke test +8 시나리오 (실제 HWP 검증)
- [ ] B-8 README / index v1.0 마감
- [ ] tag v1.0.0 + PyPI publish + GitHub Release

---

## 참고 문서

- `docs/V1_CONSISTENCY_PLAN.md` — 원래 V1 설계 철학
- `docs/DOCUMENT_PRESETS_PLAN.md` — Preset 3-layer 아키텍처
- `docs/REFACTORING_PLAN.md` — 기존 refactor 로드맵
- `docs/PYHWPX_COMPARISON.md` — 외부 라이브러리 기능 비교
