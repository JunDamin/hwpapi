# Changelog

모든 주요 변경사항이 이 문서에 기록됩니다.
포맷은 [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), 버전은
[Semantic Versioning](https://semver.org/spec/v2.0.0.html) 을 따릅니다.

## [Unreleased]

*(준비 중인 변경사항 없음 — 다음 릴리즈 예정)*

## [0.0.9] — 2026-04-15 — Styles parser · MoveAccessor sub-groups

### Added

- **MoveAccessor sub-grouping** — 38개의 flat 메소드를 의미 단위로 묶은
  sub-accessor 도입 (기존 flat API 는 호환 유지):
  - `app.move.doc.top() / bottom()`
  - `app.move.line.start() / end() / next() / prev()`
  - `app.move.word.start() / end() / next() / prev()`
  - `app.move.para.start() / end() / next() / prev()`
  - `app.move.char.next() / prev()`
  - `app.move.page.top() / bottom() / next() / prev()`
  - `app.move.cell.left() / right() / up() / down() / start() / end() / top() / bottom()`

### Fixed

- **`app.styles` parser 완전 작동** — pyhwpx 의 `FileSaveBlock_S` +
  `HFileOpenSave` pset 패턴을 채용해 HWPML2X export 후 `<STYLE Name>`
  태그 파싱. 22개 기본 스타일 (바탕글, 본문, 개요 1~8 등) 정확히 인식.
- SaveBlockAs/SaveBlockAction 존재 가정 제거 (HWP COM 에 없음).
- SelectionMode 검증 추가 — 선택 없이 호출해도 안전하게 실패.

## [0.0.8] — 2026-04-15 — Phase D · E

### Added

- `app.styles` ([`StylesAccessor`](hwpapi/classes/styles.py)) — 문단
  스타일 조회/적용/삭제/import/export.
- `app.controls` ([`ControlsAccessor`](hwpapi/classes/controls.py)) —
  문서 내 컨트롤 (표·그림·구역·하이퍼링크 등) linked-list 순회·검색.
- `Control`, `Style` 값 객체 — `ctrl.select()`, `style.apply()` 등.

## [0.0.7] — 2026-04-15 — 내부 중복 감사

### Deprecated

- `App.charshape()` legacy 빌더 — `DeprecationWarning` 추가. v0.1.0 에서 제거 예정.

### Changed

- 6개 public member 의 docstring cross-reference 추가 (selection / visible /
  new_document / rgb_color / set_visible / 단위 변환).

## [0.0.6] — 2026-04-15 — Phase C

### Added

- `app.goto_page(n)`, `app.highlight(color)`,
  `app.save_page_image(n, path)`, `app.save_all_page_images(dir)`
- `app.mm_to_hwpunit`, `point_to_hwpunit`, `hwpunit_to_mm`,
  `hwpunit_to_point`, `rgb_color`
- `app.status` property, `Color.from_rgb`, `Color.from_hex`

### Fixed

- `app.current_page` 가 section 번호를 반환하던 버그 (KeyIndicator[5] 로 수정)

## [0.0.5] — 2026-04-15 — Phase A · B

### Added

- **Field API (Mail Merge)** — `create_field`, `set_field`, `get_field`,
  `fields`, `fields_dict`, `field_exists`, `move_to_field`,
  `delete_field`, `delete_all_fields`, `rename_field`,
  `replace_brackets_with_fields`
- **pandas 연동** — `insert_table(data=df)`, `read_table(to=...)`
- **`app.silenced()`** context manager — 대화상자 자동 응답
- **Mail Merge 튜토리얼** 추가

## [0.0.4] — 2026-04-15 — Phase 1 refactor (_Action + Color/UNSET)

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
