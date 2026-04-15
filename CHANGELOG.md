# Changelog

모든 주요 변경사항이 이 문서에 기록됩니다.
포맷은 [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), 버전은
[Semantic Versioning](https://semver.org/spec/v2.0.0.html) 을 따릅니다.

## [Unreleased] — Phase 1+2 리팩토링

### Added
- `hwpapi/core/document.py` — Document/Documents 클래스를 별도 모듈로 분리
  (app.py 2,322줄 → 1,930줄)
- `Color` 클래스 + `UNSET` 센티넬 — `hwpapi.parametersets` 에서 import
- 21개 새 단위 테스트 (`tests/test_color_semantics.py`)
- `CHANGELOG.md` 도입
- `docs/REFACTORING_PLAN.md` — 전체 리팩토링 로드맵 (524줄)

### Changed
- **P0-1 버그 수정**: `_Action` 이 문서별 lazy-cache 로 동작
  (`@property act`, `@property pset` 로 재설계). 이제
  `app.actions.X.run()` 이 여러 문서에서 올바르게 각 문서에 적용됩니다.
- **P0-2**: 28개 Color 관련 PropertyDescriptor 를 `ColorProperty` 로 승격.
  TextColor / ShadeColor / UnderlineColor / ShadowColor / DiagonalColor /
  Border*Color 등.
- `ColorProperty` 쓰기 의미론 명시화:
  - `prop = UNSET` → 변경 없음 (no-op)
  - `prop = None` → 필드 제거
  - `prop = "#FF0000" / Color / int` → 정규화 후 저장
- `ColorProperty` 읽기는 이제 항상 `Color` 인스턴스 반환
  (`__str__`, `__eq__` 로 기존 문자열 비교 호환)
- `App.insert_text` → `_Action` 수정 후 캐시된 action 경로로 복귀

### Removed
- `hwpapi/core/app.py` 내 Document/Documents 정의 (document.py 로 이동)

## [0.0.3] — 2026-04-15

### Added — 다중 문서 관리 API
- `Document` 클래스 (IXHwpDocument 래퍼)
- `Documents` 컬렉션 (`app.documents`)
- `app.use_document(...)` context manager
- **xlwings-style Document proxy** — `Document.__getattr__` 로 App 의
  모든 메소드/속성을 자동 프록시 (활성화 → 호출 → 복원)

### Added — 서식 자동 리셋 헬퍼
- `app.styled_text(text, **fmt)` — 단일 구 스타일
- `app.charshape_scope(**fmt)` — CharShape 블록
- `app.parashape_scope(**fmt)` — ParaShape 블록
- 모두 snapshot + restore 패턴 — shade_color 같은 까다로운 속성도 정확히 복원

### Added — 표준 word-processor 메소드
- Properties: `text` (get/set), `visible`, `version`, `page_count`,
  `current_page`, `selection`, `saved`, `name`
- Clipboard: `select_all`, `clear`, `undo`, `redo`, `copy`, `paste`, `cut`, `delete`
- Breaks: `insert_page_break`, `insert_line_break`,
  `insert_paragraph_break`, `insert_tab`
- High-level: `insert_heading`, `insert_table`, `insert_hyperlink`,
  `insert_bookmark`, `new_document`

### Added — 문서 웹사이트
- Quarto 기반 완전 재설계 (`https://jundamin.github.io/hwpapi`)
- 8개 튜토리얼 (기초 4 + 사용사례 4)
- 전체 API 레퍼런스 (6,676줄 자동생성)

### Added — CI/CD
- PyPI Trusted Publisher — GitHub Release 발행 시 자동 배포
- Modern GitHub Pages deploy (Node 24 호환, 기존 fastai 워크플로우 실패 해결)

### Added — E2E 테스트 (31개 시나리오)
- `tests/smoke_scenarios.py` (13개)
- `tests/smoke_features.py` (18개)
- 모두 read-back assertion 으로 실제 결과 검증

### Changed
- `parametersets.py` (4,198줄) 를 15파일 패키지로 분할
- `core.py` → `core/engine.py` + `core/app.py`
- `classes.py` → `classes/accessors.py` + `classes/shapes.py`
- 기본 로그 레벨 `INFO` → `WARNING` (production 친화적)

### Fixed
- HTML 튜토리얼 preview 를 실제 HWP → PDF → PNG 렌더링으로 교체
- `app.get_text()` 기본 범위가 현재 줄뿐 → scan 매개변수로 전체 문서 가능

## [0.0.2.5] 및 이전

초기 nbdev-based 구조에서 출발한 버전. 상세 기록은
`REFACTORING_SUMMARY.md` 및 `PSET_MIGRATION_SUMMARY.md` 참고.

---

## 버전 관리 정책

- **v0.0.x** — 초기 베타, 구조 정립 단계
- **v0.1.x** — Phase 2 구조 재편 (app.py 분할, CharShape 중복 제거,
  accessor 재그룹)
- **v1.0** — 호환성 깨는 API 표준화 (Fluent 반환, deprecated 제거).
  Migration guide 제공.

마이그레이션 전략은 [`docs/REFACTORING_PLAN.md`](docs/REFACTORING_PLAN.md)
참고.

[Unreleased]: https://github.com/JunDamin/hwpapi/compare/v0.0.3...HEAD
[0.0.3]: https://github.com/JunDamin/hwpapi/releases/tag/v0.0.3
