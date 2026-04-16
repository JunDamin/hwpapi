# Changelog

모든 주요 변경사항이 이 문서에 기록됩니다.
포맷은 [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), 버전은
[Semantic Versioning](https://semver.org/spec/v2.0.0.html) 을 따릅니다.

## [Unreleased]

*(준비 중인 변경사항 없음 — 다음 릴리즈 예정)*

## [0.0.25] — 2026-04-15 — v1.0 Phase 3: ENUM 통일 + Color 상수

사용자 지적 ("선 모양도 enum 같은 걸로 되어야 할 거 같은데?") 반영.
매직 넘버 (`BorderType=8`, `hatch_style=6`) → 사용자 친화 이름 (`"double"`, `"diagonal_cross"`).

### Added

**`hwpapi/parametersets/mappings.py`** — 4개 신규 enum/MAP:
- `BORDER_TYPE_MAP` — 16종 선 종류 (`"none"`, `"solid"`, `"dash"`, `"dot"`,
  `"dash_dot"`, `"long_dash"`, `"double"`, `"wave"`, `"thick_3d"` 등)
- `HATCH_STYLE_MAP` — 13종 빗금 패턴 (`"horizontal"`, `"diagonal_cross"`,
  `"dense_horizontal"` 등)
- `CELL_APPLY_TO_MAP` — 셀 적용 범위 (`"current"`/`"selected"`/`"all"`)
- `DIAGONAL_FLAG_MAP` — 대각선 비트플래그 (`"top"`/`"middle"`/`"bottom"`/`"all"`)
- `resolve_enum(map, value)` 헬퍼 — 문자열/정수 양방향 해석 (case-insensitive)

**`Color` 클래스 상수 16개** ([`properties.py`](hwpapi/parametersets/properties.py)):
```python
Color.RED, Color.GREEN, Color.BLUE, Color.BLACK, Color.WHITE,
Color.YELLOW, Color.CYAN, Color.MAGENTA, Color.ORANGE, Color.PURPLE,
Color.PINK, Color.BROWN, Color.GRAY, Color.LIGHT_GRAY,
Color.DARK_GRAY, Color.NAVY
```

**`get_rgb_tuple` named-color 5 → 16개 확장** ([`functions.py`](hwpapi/functions.py))

### Changed

- **`set_cell_border(top="solid", bottom="double", ...)`** — 문자열 enum
  지원. 기존 정수도 그대로 작동.
- **`set_cell_color(hatch_style="diagonal_cross", ...)`** — 문자열 지원.
  추가로 FillAttr 위치 버그 fix (SelCellsBorderFill.FillAttr 사용 + ApplyTo=2)

### Tests

`tests/test_enums.py` — 14 신규 테스트 (BORDER_TYPE_MAP / HATCH_STYLE_MAP /
CELL_APPLY_TO_MAP / DIAGONAL_FLAG_MAP / resolve_enum / Color.RED 등 16 상수).

전체 1388/1388 통과.

## [0.0.24] — 2026-04-15 — v1.0 Phase 1: P0/P1 버그 fix

v1.0 release 전 분명히 깨진 5개 P0 + P1 버그 수정. 사용자 지적
("폰트 같은 건 잘 모르겠다") 가 정확히 깨진 부분.

### Fixed (P0)

**P0-1: `app.convert.replace_font(old, new)` 가 `old` 인자를 무시**
([`hwpapi/classes/convert.py`](hwpapi/classes/convert.py))
- 이전: SelectAll → set_charshape(facename_*=new) — `old` 무시, 문서 전체 덮어씀
- 이후: HFindReplace pset 의 FindCharShape/ReplaceCharShape facename 7개
  설정 + AllReplace 액션 — `old` 폰트 영역만 정확히 교체
- legacy 동작은 `replace_all=True` 로 명시적 opt-in

**P0-2: `app.config.default_font` no-op**
([`hwpapi/classes/lint.py`](hwpapi/classes/lint.py))
- 이전: dict 저장만, 어디서도 적용 안 됨
- 이후: `app.config.apply_defaults()` 메소드 추가 — 명시 호출 시 charshape
  + parashape 에 실제 적용 (font 7개 facename fan-out + size + line_spacing)

**P0-3: `App.charshape()` method vs `charshape` property 이름 충돌**
([`hwpapi/core/app.py`](hwpapi/core/app.py))
- deprecated method 가 property 에 의해 완전히 shadowed — 호출 불가
- 메소드 완전 제거, property 만 유지

**P0-4: `app.api.Run()` None 반환 패턴 11곳**
([`hwpapi/classes/accessors.py`](hwpapi/classes/accessors.py),
[`hwpapi/core/app.py`](hwpapi/core/app.py))
- HWP 12 의 Run() 이 항상 None 반환 — `if not Run(...): break` 즉시 탈출
- `hwpapi.functions.cell_addr(app)` + `navigate_until(app, action)` 신규
  공유 헬퍼 — KeyIndicator()[8] 의 셀 주소로 진행 추적
- TableAccessor.header_row/footer_row/clean_excel_paste, App.insert_table
  cursor 복귀, Selection.current_word 모두 fix

**P0-5: `app.view.zoom()` 의 `PictureScale` 액션 잘못됨**
([`hwpapi/classes/view.py`](hwpapi/classes/view.py))
- 이전: `CreateAction("PictureScale")` — 그림 크기 조절용 액션 (zoom 아님)
- 이후: `XHwpDocumentInfo.ZoomRate` property 직접 설정. fallback 으로
  `HAction.Execute("ScreenZoom", HZoom)`

### Fixed (P1)

**P1: `set_charshape(facename=...)` / `find_text(facename=...)` 다국어 fan-out**
- 이전: FaceNameHangul 만 설정 — 영문/일본어 텍스트 매칭 실패
- 이후: 7개 facename (Hangul/Latin/Japanese/Hanja/Other/Symbol/User) 모두에
  자동 fan-out. 사용자가 `facename_hangul` 등을 명시하면 그 값이 우선

### Added

- **`hwpapi.functions.cell_addr(app)`** — 현재 셀 주소 ("A1" 등) 반환
  ([`hwpapi/functions.py`](hwpapi/functions.py))
- **`hwpapi.functions.navigate_until(app, action, max_iters)`** —
  cell_addr 변화로 루프 종료 감지 헬퍼

### Tests

- `tests/test_convert.py` — replace_font 가 AllReplace 사용 + replace_all
  legacy + empty old 검증 (3개 신규)
- `tests/test_view.py` — zoom 이 ZoomRate property 설정 확인 (4개 수정)
- 전체 1374/1374 통과

## [0.0.22] — 2026-04-15 — 튜토리얼 강화 · 4개 신규 튜토리얼 추가

v0.0.14~21 의 18개 accessor 와 30+ 프리셋이 튜토리얼에 반영이 안 돼 있던
것을 보강. **4개 신규 튜토리얼** 추가, 기존 튜토리얼 2개에 discovery 포인터 삽입.

### Added

- **`10_accessors_overview.ipynb`** — 18개 accessor 전체 투어.
  - 매트릭스 표 + 카테고리별 소개 (Navigation/Collections/Structure/
    Transform/Quality/Presets)
  - 각 accessor 의 대표 메소드 + 예제 1~2개
  - 8개 context manager 목록 + 사용법
- **`11_presets_gallery.ipynb`** — 문서 꾸미기 프리셋 쇼케이스.
  - 11개 preset 의 before/after 예제
  - 조합 레시피: 공공 보고서 / 컬러 보고서
- **`12_batch_and_workflow.ipynb`** — 대량 처리 + context manager.
  - silenced / suppress_errors / batch_mode / undo_group /
    charshape_scope / use_document
  - 3개 실전 workflow: 급여명세서 1000장 · 폰트 재귀 교체 · 디렉터리 품질 감사
- **`13_debugging_tools.ipynb`** — 디버깅 / 품질 / 설정.
  - `app.help()` / `repr(app)` discovery
  - `app.debug.state/print/timing/trace`
  - `app.lint()` + CI 대량 품질 감사 예제
  - `app.template.save/apply` + 대량 적용
  - `app.config` + 팀 공유 json

### Changed

- **`_quarto.yml` sidebar** 재구성 — 3개 섹션:
  - 📘 기초 튜토리얼 (10_accessors_overview 추가)
  - 📗 사용 사례 (Mail Merge 노출)
  - 🎨 v0.0.14+ 신규 기능 (11, 12, 13)
- **`index.qmd`** 랜딩 카드 추가 — 신규 튜토리얼 3개 + 새 기능 미리보기 bullet 10개
- **`03_feature_tour.ipynb`** — 상단에 신규 기능 pointer 섹션 삽입
- **`01_app_basics.ipynb`** — `app.help()` tip 박스 삽입

### Tests

영향 없음 — 튜토리얼은 런타임 테스트 대상 아님. 1362/1362 유지.

## [0.0.21] — 2026-04-15 — 테스트 도메인별 재구성

**Phase C.** 버전별 테스트 파일(`test_v014_features.py` ~ `test_v020_deprecations.py`)
를 **도메인별 13개 파일** 로 재편. 이제 "Selection 관련 테스트" 를 찾으려면
``tests/test_selection.py`` 하나만 열면 된다.

### Changed

- 기존 7개 버전별 파일 → 13개 도메인별 파일:
  - `test_selection.py` (18) — Selection accessor 전반
  - `test_images.py` (5) — Images accessor
  - `test_presets.py` (20) — 모든 Preset 메소드
  - `test_table_accessor.py` (8) — TableAccessor batch + clean_excel_paste
  - `test_debug.py` (12) — Debug (state/trace/timing/print)
  - `test_context_managers.py` (2) — batch_mode, undo_group
  - `test_convert.py` (6, 14 실행) — Convert + _int_to_korean parametrize
  - `test_view.py` (9) — View accessor
  - `test_lint.py` (10) — Linter
  - `test_template.py` (4) — Template
  - `test_config.py` (8) — Config
  - `test_discovery.py` (9) — app.help() + __repr__
  - `test_deprecations.py` (12) — 레거시 API DeprecationWarning

- 각 파일은 AST 파싱 기반 자동 재분배. 공통 imports + 관련 fixture 는 여러
  파일에 복제되어 독립 실행 가능 (pytest 단일 파일 실행 친화).

### Tests

테스트 수는 그대로 유지 (1362/1362 통과) — 재배치만 수행.

## [0.0.20] — 2026-04-15 — 레거시 API DeprecationWarning + API_GUIDE.md 재작성

**Phase B 중복 제거.** 17개 레거시 method 에 DeprecationWarning 부착 (호환
유지). ``hwpapi.*`` 패키지 내부에서 호출되면 경고 무시 — accessor 가
레거시 method 를 내부적으로 delegate 해도 사용자에겐 spam 없음.

### Added

- **`hwpapi.core.app._warn_legacy(old, new)`** — 스마트 deprecation 헬퍼.
  caller frame 의 모듈이 ``hwpapi.`` 로 시작하면 조용히 통과, 사용자
  코드에서 호출되면 DeprecationWarning 발생.

### Changed

- 10개 레거시 field 관련 method 에 DeprecationWarning 부착:
  - ``app.create_field`` → ``app.fields.add``
  - ``app.set_field`` → ``app.fields[n] = v``
  - ``app.get_field`` → ``app.fields[n].value``
  - ``app.field_exists`` → ``n in app.fields``
  - ``app.move_to_field`` → ``app.fields[n].goto()``
  - ``app.delete_field`` → ``app.fields.remove(n)``
  - ``app.delete_all_fields`` → ``app.fields.remove_all()``
  - ``app.rename_field`` → ``app.fields.rename(o, n)``
  - ``app.field_names`` (property) → ``list(app.fields)``
  - ``app.fields_dict`` (property) → ``app.fields.to_dict()``
- ``app.insert_bookmark`` / ``app.insert_hyperlink`` / ``app.replace_brackets_with_fields``:
  soft deprecation note in docstring (Fields/Bookmarks/Hyperlinks
  accessor 가 현재 이들을 호출하므로 warning 은 생략).
- **`docs/API_GUIDE.md` 재작성** — 18개 accessor 매트릭스 표, context
  manager 목록, property 목록, legacy→신규 마이그레이션 표 추가.
- **`App.field_names_internal()`** — warning 없는 내부용 field list (accessor 와 fields_dict 가 사용).

### Tests

12 개 신규 단위 테스트 ([`tests/test_v020_deprecations.py`](tests/test_v020_deprecations.py)):
- `_warn_legacy` 기본 동작
- ``hwpapi.*`` 모듈에서 호출 시 silent 확인
- 각 레거시 method 가 warning 발생시키는지
- Fields accessor 가 레거시 호출해도 사용자 코드에 warning 안 가는지

전체 1362/1362 통과.

## [0.0.19] — 2026-04-15 — Discovery · app.help() + 상태 요약 __repr__

**Phase A 사용자 경험 개선** (18개 accessor 가 97개 method 속에 숨어있던 문제 해결).

### Added

- **`app.help()`** — App 에서 사용 가능한 accessor·context manager·주요
  property 를 **6개 카테고리로 그룹핑**해서 출력. 사용자가 `app.` 찍고
  혼란스러워하지 않도록 최초 학습 곡선을 대폭 축소.
  - Navigation & Selection (move, sel)
  - Collections (documents, fields, bookmarks, hyperlinks, images, styles, controls)
  - Structure (cell, table, page)
  - Transform & View (convert, view)
  - Quality & Templates (lint, template, config)
  - Presets & Debug (preset, debug)
  - Context managers (8개)
- **`App.__repr__`** — 상태 요약을 한 줄로:
  ``App(visible=True, version='13.0.0', docs=2, page=5/20)``.
  예외 완전 안전 (초기화 안 된 인스턴스에도 작동).
- **`App._ACCESSOR_MAP`** — 문서화/help 용도의 공개 가능한 카테고리 매핑.
- **`App._CONTEXT_MANAGERS`** — 동일 패턴으로 context manager 목록.

### Changed

- **`__init__` accessor 할당부 재구성** — 도메인별로 그룹핑 + 한 줄 주석.
  시대순 ("v0.0.12 추가", "v0.0.14 추가") 에서 용도 기반으로 전환.
- **`__str__` / `__repr__` 분리** — 기존에는 ``__repr__ = __str__`` 으로
  파일 경로만 표시됐으나, 이제 ``__str__`` 는 파일 경로, ``__repr__`` 은
  상태 요약. 초기화 안 된 인스턴스에서도 안전하게 작동.

### Tests

9 개 신규 단위 테스트 ([`tests/test_v019_discovery.py`](tests/test_v019_discovery.py)).
전체 1350/1350 통과.

## [0.0.18] — 2026-04-15 — Linter · Template · Config

### Added

**`app.lint()`** ([`classes/lint.py`](hwpapi/classes/lint.py)) — 문서 품질 체크:
- callable accessor: ``report = app.lint()`` 반환
- 검사 항목:
  - 긴 문장 (기본 80자 초과) → ``report.long_sentences``
  - 긴 문단 (기본 500자 초과) → ``report.long_paragraphs``
  - 빈 문단 → ``report.empty_paragraphs``
  - 연속 공백 → ``report.double_spaces``
  - 끝 공백 → ``report.trailing_whitespace``
- ``report.has_issues``, ``.issue_count``, ``.summary`` 헬퍼

**`app.template`** ([`classes/lint.py`](hwpapi/classes/lint.py)) — 문서 템플릿:
- ``template.save(path)`` — 현재 charshape/parashape/page 설정을 JSON 으로
- ``template.apply(path)`` — JSON 템플릿을 현재 문서에 적용

**`app.config`** ([`classes/lint.py`](hwpapi/classes/lint.py)) — App 선호도:
- ``config.default_font``, ``.default_size``, ``.default_line_spacing``,
  ``.default_table_style``, ``.palette``
- ``config.update(**kw)``, ``.reset()``, ``.to_dict()``
- ``config.save(path="~/.hwpapirc")`` / ``.load(path)``

### Tests

22 개 신규 단위 테스트 ([`tests/test_v018_features.py`](tests/test_v018_features.py)).
전체 1341/1341 통과.

## [0.0.17] — 2026-04-15 — Convert · View · more presets

### Added

**`app.convert`** ([`classes/convert.py`](hwpapi/classes/convert.py)):
- `convert.number_to_korean(text=None)` — 숫자 → 한글숫자
  (``"1,234,567" → "일백이십삼만사천오백육십칠"``). ``text=None`` 이면
  현재 선택 영역 변환. 최대 경(10^16) 까지.
- `convert.wrap_by_word()` / `.wrap_by_char()` — 줄 나눔 방식 전환
- `convert.replace_font(old, new)` — 문서 전체 폰트 일괄 교체

**`app.view`** ([`classes/view.py`](hwpapi/classes/view.py)):
- `view.zoom(percent)` — 10~500% 클램프
- `view.zoom_fit_page()` / `.zoom_fit_width()` / `.zoom_actual()`
- `view.zoom_current` — 현재 확대율 read-only property
- `view.scroll_to_cursor()`
- `view.full_screen()` / `.exit_full_screen()`
- `view.page_mode()` / `.draft_mode()`
- `view.toggle_rulers()`

**추가 Presets** ([`presets/__init__.py`](hwpapi/presets/__init__.py)):
- `preset.page_border(enable, style)` — 바탕쪽 테두리 (편집 영역 디버그)
- `preset.highlight_yellow(toggle=True)` — 선택 영역 노란 배경 토글
- `preset.summary_box(text, variant, bg_color)` — 강조 박스

### Tests

30 개 신규 단위 테스트 ([`tests/test_v017_features.py`](tests/test_v017_features.py)).
전체 1319/1319 통과.

## [0.0.16] — 2026-04-15 — 개발자 생산성 · batch/undo/debug · 표 일괄 서식

### Added

**Context managers** ([`core/app.py`](hwpapi/core/app.py)):
- `with app.batch_mode(hide=True):` — 대량 처리 시 화면 숨김 + dialog 억제
  + ScrollFollow off. 종료 시 자동 복원. 일반 대비 5~10배 빠름.
- `with app.undo_group("설명"):` — 블록 내 모든 편집을 단일 undo 경계로 묶음.

**`app.debug`** ([`classes/debug.py`](hwpapi/classes/debug.py)) — 디버깅 accessor:
- `debug.state()` — 커서, 페이지, 선택, charshape, in_table, 열린 문서 수,
  visible, version, filepath 를 dict 로 덤프
- `debug.print()` — state 를 예쁘게 출력
- `debug.timing(fn, *args)` — 함수 호출 시간 측정 (ms)
- `with debug.trace():` — 블록 내 모든 ``Run()`` 호출 로그

**표 일괄 서식** ([`classes/accessors.py`](hwpapi/classes/accessors.py) TableAccessor):
- `table.header_row(bold, bg, text_color)` — 첫 행 서식 일괄
- `table.footer_row(bold, bg, text_color)` — 마지막 행 서식 일괄
- `table.current_row(bold, bg, text_color)` — 현재 행 서식
- `table.align(horz, vert, scope)` — 정렬 일괄 적용
  (``scope``: ``"current_cell" | "current_row" | "current_col" | "all"``)

### Tests

23 개 신규 단위 테스트 ([`tests/test_v016_features.py`](tests/test_v016_features.py)).
전체 1289/1289 통과.

## [0.0.15] — 2026-04-15 — Preset Phase 1 · structure + TOC + spacing

### Added

**Structure presets** (승승아빠 매크로 Alt+1, Alt+2 이식):
- `app.preset.title_box(text, subtitle, bg_color, font_size)`
- `app.preset.subtitle_bar(text, bg_color)`

**Table header/footer** (Alt+4, Alt+5, Alt+Shift+4, Alt+Shift+5):
- `app.preset.table_header(color="sky", text_color="#FFFFFF", rows=1)`
- `app.preset.table_footer(color="gray", rows=1)`
- 색상 preset: ``"gray" | "sky" | "dark_blue" | "green" | "red"`` 또는 hex

**Navigation** (Alt+3, 쪽번호 매크로):
- `app.preset.toc(with_bookmarks=True, dot_leader=True, levels=3)` — 점끌기탭 TOC
- `app.preset.page_numbers(position, format, header_filename)`

**Selection — 자간/장평** (Alt+Shift+8, Alt+Shift+9):
- `app.sel.compress(step=1)` — 자간/장평 동시 축소
- `app.sel.expand(step=1)` — 자간/장평 동시 확대

### Fixed

- ``while app.api.Run(...)`` 패턴을 ``for _ in range(500)`` 으로 교체 —
  병든 표(500행 이상)나 mock 테스트에서 무한 루프 방지 (clean_excel_paste,
  striped_rows, table_header, table_footer 모두).

### Tests

16 개 신규 단위 테스트 ([`tests/test_v015_features.py`](tests/test_v015_features.py)).
전체 1266/1266 통과.

## [0.0.14] — 2026-04-15 — Presets · Images · Selection accessors

DOCUMENT_PRESETS_PLAN.md 의 Phase 1 착수. 승승아빠 매크로에서 범용성
높은 기능을 추려 Python API 로 이식.

### Added

- **`app.preset`** ([`Presets`](hwpapi/presets/__init__.py)) — 문서 꾸미기 프리셋 accessor.
  - `preset.striped_rows(colors=[...], header_color=None)` — 줄무늬 표 (zebra).
- **`app.images`** ([`Images`](hwpapi/classes/images.py)) — 이미지 control 컬렉션.
  - iter / len / `images[i]` — 이미지 순회
  - `images.resize_all(width="100mm", keep_ratio=True)` — 일괄 크기 조정
  - `images.grayscale_all()` — 흑백 변환
- **`app.sel`** ([`Selection`](hwpapi/classes/selection.py)) — 선택 동작 accessor.
  (``app.selection`` 은 str property 로 그대로 유지, ``app.sel`` 은 ``app.move`` 와 대칭인 accessor.)
  - `sel.current_word()` / `.current_line()` / `.current_paragraph()` / `.current_sentence()`
  - `sel.to_paragraph_end()` / `.to_paragraph_begin()` / `.to_line_end()` / `.to_line_begin()`
  - `sel.to_document_end()` / `.to_document_begin()`
  - `sel.expand_char(n)` / `.expand_word(n)`
  - `sel.clear()` / `sel.text` / `sel.is_empty`
- **`app.table.clean_excel_paste()`** — 엑셀→HWP 붙여넣은 표의 빈 행/열/공백 정돈
  (승승아빠 매크로 ``엑셀_복사표_숫자_빈칸지우기`` 의 Python 포팅).

### Tests

19 개의 새 단위 테스트 ([`tests/test_v014_features.py`](tests/test_v014_features.py)).
전체 1250/1250 통과 (regression 0).

## [0.0.13] — 2026-04-15 — V1 Phase 1 · Fluent API

### Changed

V1.0 청사진의 "모든 set 메소드에 `return self` 추가" 항목 이행.
**기존 None 반환처에 영향 없음** — 반환값을 쓰지 않는 코드는 그대로
작동, 반환값을 체이닝에 쓰는 새 코드도 자연스럽게 작성 가능:

- `app.insert_text(...)` → ``self``
- `app.styled_text(...)` → ``self``
- `app.insert_heading(...)` → ``self``
- `app.insert_table(...)` → ``self``
- `app.insert_hyperlink(...)` → ``self``
- `app.insert_bookmark(...)` → ``self``
- `app.insert_paragraph_break()` → ``self``
- `app.insert_page_break()` → ``self``
- `app.insert_line_break()` → ``self``
- `app.insert_tab()` → ``self``

값을 반환해야 하는 메소드 (``save`` → 경로, ``open`` → 경로,
``find_text`` → bool, ``replace_all`` → count, ``page_count`` 등)
는 그대로 유지. 이는 V1.0 청사진의 "prefix 기반 분류" 원칙에 부합.

### Added

- Fluent chain 사용 예:

  ```python
  (app
   .insert_heading("보고서", level=1)
   .insert_text("1. 개요\n")
   .insert_paragraph_break()
   .insert_table(rows=3, cols=4)
   .insert_text("서명: ")
   .styled_text("대표이사", bold=True)
   .insert_paragraph_break())
  ```

- 9개의 Fluent API 단위 테스트 ([`tests/test_fluent_api.py`](tests/test_fluent_api.py)).

## [0.0.12] — 2026-04-15 — V1 Phase 1 · Fields/Bookmarks/Hyperlinks accessors + charshape/parashape properties

### Added

- **`app.fields` 컬렉션 accessor** ([`Fields`](hwpapi/classes/fields.py)) —
  v1.0 청사진 Phase 1. **하위 호환**: iteration / ``in`` / ``len()`` 은
  여전히 필드 이름 (str) 으로 동작. 추가로 dict-style + collection 메소드:
  - ``app.fields["name"]`` → ``Field`` 객체 (값 객체 — `.value`, `.goto()`, `.remove()`)
  - ``app.fields["name"] = "값"`` → 자동 생성 + 값 주입
  - ``app.fields.add(name, memo, direction)``
  - ``app.fields.remove(name)`` / ``remove_all()``
  - ``app.fields.find(name)`` → Optional[Field]
  - ``app.fields.rename(old, new)``
  - ``app.fields.update({...})`` / ``update(**kw)`` — 일괄 주입
  - ``app.fields.to_dict()``
  - ``app.fields.from_brackets(pattern)`` — `{{tag}}` → 필드 변환
- **`app.bookmarks`** ([`Bookmarks`](hwpapi/classes/fields.py)) —
  ``add(name)``, ``remove(name)``, ``goto(name)``, ``"name" in app.bookmarks``.
- **`app.hyperlinks`** ([`Hyperlinks`](hwpapi/classes/fields.py)) —
  ``add(text, url)`` → ``Hyperlink`` 값 객체.
- **`app.charshape` / `app.parashape` properties** — read 는 현재 커서의
  스냅샷 반환, write 는 전체 교체 (또는 dict 로 partial). 기존
  :meth:`get_charshape` / :meth:`set_charshape` 는 그대로 유지.
- 28개의 신규 단위 테스트 ([`tests/test_fields_accessor.py`](tests/test_fields_accessor.py)).

### Changed

- **`app.fields`** 가 단순 list 가 아닌 ``Fields`` 컬렉션 반환. 단,
  ``for n in app.fields:`` / ``"x" in app.fields`` / ``len(app.fields)``
  같은 기존 사용 패턴은 그대로 작동 (deprecation warning 없음).
- **`app.field_names`** — 기존 list 반환 property 의 새 이름. 명시적 list
  가 필요한 경우 권장 (``app.fields`` 도 ``list(app.fields)`` 로 동일).

### Deprecated

- ``app.fields[0]`` (정수 인덱싱) → ``DeprecationWarning``. 대신
  ``list(app.fields)[0]`` 또는 ``app.fields["name"]`` 사용.

## [0.0.11] — 2026-04-15 — silenced() context manager · v1.0 청사진

### Added

- **`SILENCE_NO_SAVE = 0x00100000`** — hwp-mcp 호환 alias ("저장 안함" 자동 선택).
- **`silenced()` 의 풍부한 string preset** — `"ok"`, `"save"`, `"save_yes"`,
  `"save_no"`, `"no_save"`, `"okcancel_ok"`, `"okcancel_no"` (단일 카테고리
  세밀 제어). 기존 `"yes"` / `"no"` / `"reset"` 와 공존.
- **`docs/V1_CONSISTENCY_PLAN.md`** — 하위 호환을 포기한 v1.0 일관성
  로드맵. 6대 원칙(Pythonic state vs action, domain accessor, Fluent
  return, 동사 어휘 통일, context manager 명명, 명시적 에러 정책),
  전체 API 청사진, 10주 마이그레이션 일정 포함.

### Changed

- **`silenced()` docstring 명확화** — context manager 임을 강조. 영구
  적용은 `set_message_box_mode()`, scoped 적용은 `silenced()` 로 분리.
- **smoke 테스트 (smoke_scenarios.py / smoke_features.py)** — 영구
  `set_message_box_mode()` 호출 대신 `with app.silenced("yes"):` context
  manager 로 감싸 자동 복원. 백그라운드 dismisser 는 fallback 으로 유지.

## [0.0.10] — 2026-04-15 — silenced() 6 dialog categories

### Added

- **`silenced()` 6-카테고리 비트필드 지원** — 0xF/0xF0/0xF00/0xF000/0xF0000/0xF00000
  6개 dialog 카테고리 모두 한 번에 자동응답. `SILENCE_ALL_YES = 0x111111`,
  `SILENCE_ALL_NO = 0x222222`, `SILENCE_RESET = 0` 등 상수 추가.
- **`suppress_errors()` context manager** — 에러 dialog 자동 ABORT + Python
  예외 swallowing. 대량 자동화에서 일부 실패에도 루프 계속 진행.
- **`register_security_module()`** — 보안 dialog 차단을 위한 외부 모듈 등록.

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
