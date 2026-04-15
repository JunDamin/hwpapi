# Session Summary: v0.0.11 → v0.0.22

> 작성일: 2026-04-15
> 범위: 한 세션 내에서 진행된 **12개 릴리즈** (v0.0.11 ~ v0.0.22)

본 문서는 `docs/V1_CONSISTENCY_PLAN.md` 와 `docs/DOCUMENT_PRESETS_PLAN.md`
에서 제시된 로드맵의 **실행 기록** 이다. 모든 기능은 backward compatible
로 추가되었으며, 레거시 API 는 deprecation warning 과 함께 유지되고 있다.

---

## 릴리즈 타임라인

| 버전 | 주제 | 핵심 변화 |
|:---:|:---|:---|
| 0.0.11 | silenced() context manager 강화 | hwp-mcp 호환 string preset, V1_CONSISTENCY_PLAN.md |
| 0.0.12 | V1 Phase 1 — Collection accessor | `app.fields` / `app.bookmarks` / `app.hyperlinks` (list+dict-like) |
| 0.0.13 | Fluent API | `insert_*` 류 → `return self` |
| 0.0.14 | Presets Phase 1 | `app.preset.striped_rows()`, `app.images`, `app.sel`, `table.clean_excel_paste()` |
| 0.0.15 | Presets 구조 / TOC / 자간 | title_box, subtitle_bar, table_header/footer, toc, page_numbers, compress/expand |
| 0.0.16 | 개발자 생산성 | `batch_mode()`, `undo_group()`, `app.debug`, 표 일괄 서식 |
| 0.0.17 | 변환 · 뷰 · 추가 preset | `app.convert`, `app.view`, page_border, highlight_yellow, summary_box |
| 0.0.18 | 품질 · 템플릿 · 설정 | `app.lint()`, `app.template`, `app.config` |
| 0.0.19 | **Phase A — Discovery** | `app.help()`, `App.__repr__`, `__init__` 도메인 그룹핑 |
| 0.0.20 | **Phase B — 중복 제거** | Legacy DeprecationWarning (hwpapi 내부 호출 silent), API_GUIDE 재작성 |
| 0.0.21 | **Phase C — 테스트 재편** | 버전별 7개 → 도메인별 13개 파일 |
| 0.0.22 | 튜토리얼 강화 | 4개 신규 튜토리얼 (10~13) + sidebar 재구성 |

---

## 성과 수치

### 코드 / 테스트

| 지표 | v0.0.10 | v0.0.22 | 변화 |
|:---|---:|---:|---:|
| 테스트 | 1,222 | **1,362** | +140 |
| 테스트 파일 | 7 (version-based) | **13 (domain-based)** | +6 |
| App accessor | 7 | **18** | +11 |
| Preset method | 0 | **11** | +11 |
| Context manager | 4 | **8** | +4 |
| Public App method | ~60 | 75 | +15 |
| 신규 클래스 | — | Fields, Bookmarks, Hyperlinks, Images, Selection, Debug, Convert, View, Linter, Template, Config, Presets (12개) | +12 |

### 문서

| 항목 | 변화 |
|:---|:---|
| 튜토리얼 | 9 → **13** (+10_accessors_overview, 11_presets_gallery, 12_batch_and_workflow, 13_debugging_tools) |
| 기획 문서 | V1_CONSISTENCY_PLAN.md, DOCUMENT_PRESETS_PLAN.md (신규) |
| API_GUIDE.md | 18-accessor 매트릭스 + 레거시 마이그레이션 표 추가 |
| CHANGELOG | 12개 릴리즈 항목 (v0.0.11~22) |

---

## 도입된 신규 accessor (11개)

### Collections (v0.0.12 ~ 14)

```python
app.fields         # Fields       — 누름틀 (dict/list-like)
app.bookmarks      # Bookmarks    — 책갈피
app.hyperlinks     # Hyperlinks   — 하이퍼링크
app.images         # Images       — 이미지 control
```

### Navigation & Selection (v0.0.14)

```python
app.sel            # Selection    — 선택 제어 (current_*, to_*, compress/expand)
```

### Transform & View (v0.0.17)

```python
app.convert        # Convert      — 숫자/폰트/줄 나눔 변환
app.view           # View         — zoom, 전체화면, 페이지 모드
```

### Quality & Templates (v0.0.18)

```python
app.lint           # Linter       — callable → LintReport
app.template       # Template     — 문서 스타일 save/apply
app.config         # Config       — App 선호도 (hwpapirc)
```

### Presets & Debug (v0.0.14 ~ 16)

```python
app.preset         # Presets      — 11개 꾸미기 프리셋
app.debug          # Debug        — state/trace/timing
```

---

## 도입된 신규 Preset (11개)

승승아빠 한글매크로(`ref/승승아빠-한글매크로-250124(공개용).zip`) 의 29개
JScript 매크로 중 범용성 높은 11개를 Python 으로 이식.

| Preset | 대응 매크로 | 요약 |
|:---|:---|:---|
| `striped_rows()` | (자체) | 줄무늬 표 (zebra) |
| `table_header(color=)` | Alt+4 | 상단 제목행 색상 |
| `table_footer(color=)` | (자체) | 하단 바닥행 색상 |
| `title_box()` | Alt+1 | 문서 타이틀 박스 |
| `subtitle_bar()` | Alt+2 | 소제목 바 |
| `summary_box()` | Alt+9 | 요약 박스 |
| `toc(with_bookmarks=, dot_leader=)` | Alt+3 | 목차 + 북마크 + 점끌기 |
| `page_numbers(position=, header_filename=)` | (자체) | 쪽번호 + 머리글 |
| `page_border(enable=)` | Alt+Shift+1 | 바탕쪽 테두리 (디버그) |
| `highlight_yellow(toggle=)` | Alt+Shift+2 | 형광펜 토글 |
| `striped_rows()` 색상 preset | — | gray/sky/dark_blue/green/red |

---

## Context Manager (4 → 8)

### 기존 (v0.0.10 이전)

```python
with app.silenced("yes"): ...             # dialog 자동 응답
with app.suppress_errors(): ...           # 에러 swallow
with app.charshape_scope(**fmt): ...      # 임시 문자 모양
with app.parashape_scope(**fmt): ...      # 임시 문단 모양
```

### 신규 (v0.0.16 +)

```python
with app.batch_mode(hide=True): ...       # 5~10배 가속 (창 숨김 + dialog off)
with app.undo_group("설명"): ...          # 단일 undo 경계
with app.use_document(doc): ...           # 활성 문서 일시 전환
with app.debug.trace(): ...               # COM Run() 호출 로그
```

---

## 3-Phase 마무리 정리 (v0.0.19 ~ 21)

### Phase A (v0.0.19) — Discovery

사용자가 `app.` 찍고 **97개 attrs** 속에서 헤매던 문제 해결.

```python
app.help()        # 6 카테고리로 accessor/context manager 출력
repr(app)         # App(visible=True, version='13.0.0', docs=2, page=5/20)
```

- `App._ACCESSOR_MAP` — 공식 매핑 (6 카테고리, 17개 accessor)
- `App._CONTEXT_MANAGERS` — 8개 context manager 목록
- `__init__` 재구성 — 시대순 → 도메인 그룹별 주석

### Phase B (v0.0.20) — 중복 제거

17개 레거시 method → 신 accessor 이중화 제거.

**_warn_legacy 헬퍼** — caller frame 의 `__name__` 이 `hwpapi.*` 로 시작하면
silent, 사용자 코드에서 호출되면 DeprecationWarning. 덕분에 accessor 가
내부적으로 legacy 를 delegate 해도 사용자는 warning 안 받음.

경고 부착된 method 10개:
`create_field, set_field, get_field, field_exists, move_to_field,
delete_field, delete_all_fields, rename_field, field_names,
fields_dict`

### Phase C (v0.0.21) — 테스트 재편

`test_v014_features.py` ~ `test_v020_deprecations.py` 7개 파일 →
도메인별 13개:

```
tests/test_selection.py
tests/test_images.py
tests/test_presets.py
tests/test_table_accessor.py
tests/test_debug.py
tests/test_context_managers.py
tests/test_convert.py
tests/test_view.py
tests/test_lint.py
tests/test_template.py
tests/test_config.py
tests/test_discovery.py
tests/test_deprecations.py
```

AST 기반 자동 재분배 스크립트로 123개 테스트를 클래스별 파일로 이동.

---

## 튜토리얼 체계 (v0.0.22)

### 사이드바 3섹션 구성

```
📘 기초 튜토리얼
    01_app_basics          — app.help() tip 추가
    02_find_and_replace
    03_feature_tour         — 신규 기능 pointer 섹션 추가
    08_multi_document
    10_accessors_overview  🆕 18개 accessor 투어

📗 사용 사례
    04_usecase_meeting_minutes
    05_usecase_bulk_edit
    06_usecase_data_to_table
    07_usecase_report_generation
    09_usecase_mail_merge   — v0.0.12 에서 Fields 패턴 추가됨

🎨 v0.0.14+ 신규 기능
    11_presets_gallery     🆕 11개 preset before/after + 2 레시피
    12_batch_and_workflow  🆕 6개 context manager + 3 실전 workflow
    13_debugging_tools     🆕 debug/lint/template/config 실전
```

---

## 남은 작업 (v1.0 breaking)

v1.0 에서 진행될 주요 breaking changes (현재 계획):

1. **legacy method 17개 완전 제거** (field/bookmark/hyperlink 관련)
2. **App 을 `core/ops/*` 로 분할** — 100+ method 를 역할별 믹스인으로
3. **일관된 동사 어휘** — `add`/`remove`/`find` 통일 (V1_CONSISTENCY_PLAN 참고)
4. **Context manager 명명 통일** — `silencing()`/`using()` verb 형태
5. **HwpError 계열 커스텀 예외** 도입
6. **자동 마이그레이션 도구** — `python -m hwpapi.migrate <file.py>`

계획 문서: [`docs/V1_CONSISTENCY_PLAN.md`](V1_CONSISTENCY_PLAN.md)

---

## 참고

- 기획 문서: [`V1_CONSISTENCY_PLAN.md`](V1_CONSISTENCY_PLAN.md), [`DOCUMENT_PRESETS_PLAN.md`](DOCUMENT_PRESETS_PLAN.md)
- API 레퍼런스: [`API_GUIDE.md`](API_GUIDE.md)
- 변경 이력: [`../CHANGELOG.md`](../CHANGELOG.md)
- 승승아빠 매크로 원본: `../ref/승승아빠-한글매크로-250124(공개용).zip`
