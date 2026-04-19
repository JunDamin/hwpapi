# v1.0 일관성 확보 계획서 (Breaking Changes 허용)

> 작성일: 2026-04-15 (v0.0.10 기준)
> 목적: **하위 호환성을 포기**하더라도 API 일관성을 확보하는 v1.0 청사진

---

## 배경

v0.0.x 동안 빠르게 기능을 더하면서 누적된 **8가지 일관성 문제** :

| # | 문제 | 예시 |
|:---:|:---|:---|
| 1 | 명명 스타일 혼재 | `set_charshape` (snake) vs `actions.CharShape` (Pascal) |
| 2 | get/set vs property 혼재 | `get_text()`/`set_charshape()` vs `app.text`/`app.visible` |
| 3 | 반환 타입 일관성 없음 | None / bool / int / str / Document / pset / dict |
| 4 | 도메인 접근자 비대칭 | `app.documents` accessor vs `app.fields` flat methods |
| 5 | 동사 다양 | `create_field` / `insert_table` / `add` / `new_document` (모두 "추가") |
| 6 | get/insert/set 비대칭 | `insert_heading` 만 있고 `get_heading` 없음 |
| 7 | Context manager 명명 | `silenced` (부정사) vs `charshape_scope` vs `use_document` |
| 8 | App 비대 | 85개 멤버 — 정리 필요 (`core/ops/*` 분리) |

---

## v1.0 핵심 원칙

### 원칙 1 — Pythonic State vs Action

> **상태(state)는 property, 행동(action)은 method**

| 종류 | 예 |
|:---|:---|
| Property (state read/write) | `app.text`, `app.visible`, `app.charshape`, `app.parashape` |
| Method (action) | `app.insert_text(...)`, `app.find(...)`, `app.save(...)` |

### 원칙 2 — 도메인 접근자 일관 적용

> 같은 도메인의 모든 작업은 **하나의 접근자 객체** 로 모음.

```python
# v0.x (현재) — 도메인이 App 에 평탄화
app.set_charshape(bold=True)
app.create_field("name")
app.set_field("name", "값")
app.insert_hyperlink(text, url)
app.insert_bookmark("ch1")

# v1.0 — 접근자 패턴
app.charshape.set(bold=True)              # 또는 app.charshape = ...
app.fields.add("name")
app.fields["name"] = "값"
app.hyperlinks.add(text=text, url=url)
app.bookmarks.add("ch1")
```

### 원칙 3 — Fluent return (chaining)

> 메소드는 **`self` 반환**, 값은 property/get 메소드가 반환.

```python
# v0.x
app.insert_text("앞: ")
app.styled_text("중요", bold=True)
app.insert_text(" 뒤")
app.save("out.hwp")

# v1.0
(app
 .insert_text("앞: ")
 .styled_text("중요", bold=True)
 .insert_text(" 뒤")
 .save("out.hwp"))
```

값을 반환하는 메소드는 **별도 prefix** 로 구분:

| Prefix | 용도 | 반환 |
|:---|:---|:---|
| `find_*` | 검색 | bool / Match / None |
| `read_*` | 추출 | DataFrame / str / list |
| `count_*` | 개수 | int |
| 그 외 (verb_*) | 동작 | self |

### 원칙 4 — 일관된 동사 어휘

| 의미 | v1.0 동사 | 폐기 |
|:---|:---|:---|
| 컬렉션에 추가 | `add` | ~~create, insert, new~~ |
| 컬렉션에서 제거 | `remove` | ~~delete~~ |
| 첫 매치 검색 | `find` | ~~search, lookup~~ |
| 모두 검색 | `find_all` | — |
| 적용 | `apply` | — |
| 읽기 | `read_*` (값) / property | ~~get~~ |
| 쓰기 | property setter / `set_*` (인자 많을 때) | — |

### 원칙 5 — Context manager 명명 통일

> **모두 동사구**: `using_X(...)` / `silencing_X(...)` 형태

```python
# v0.x — 혼재
with app.silenced(): ...
with app.charshape_scope(bold=True): ...
with app.parashape_scope(...): ...
with app.use_document(doc): ...

# v1.0 — 통일
with app.silencing(): ...
with app.with_charshape(bold=True): ...
with app.with_parashape(...): ...
with app.using(doc): ...
```

또는 동사구가 어색하면 **`scope` 접미사 통일**:

```python
with app.silence_scope(): ...
with app.charshape_scope(...): ...
with app.parashape_scope(...): ...
with app.document_scope(doc): ...
```

선호: 첫 번째 (동사형) — Pythonic context manager 관습.

### 원칙 6 — 명시적 에러 정책

| 상황 | v1.0 정책 |
|:---|:---|
| 사용자 입력 오류 (없는 필드명 등) | `KeyError` / `ValueError` raise |
| HWP 작업 실패 (권한, 파일 lock) | `HwpError` (커스텀) raise + 로그 |
| Optional / 검색 결과 없음 | `None` 반환 |
| Boolean 검사 (existence) | `True / False` 반환 |
| 부수효과만 있는 setter | `self` 반환 |

---

## v1.0 API 청사진 (요약)

### App

```python
# Properties (state)
app.text                # str (get/set)
app.visible             # bool (get/set)
app.version             # str (read-only)
app.page_count          # int (read-only)
app.current_page        # int (read-only)
app.selection           # str (read-only)
app.charshape           # CharShape (현재 커서 — get) / setter 도 가능
app.parashape           # ParaShape

# Domain accessors (collections / scoped)
app.documents           # Documents (open docs)
app.fields              # Fields (mail merge)
app.styles              # Styles (paragraph styles)
app.controls            # Controls (table/image/...)
app.bookmarks           # Bookmarks (NEW)
app.hyperlinks          # Hyperlinks (NEW)
app.images              # Images (NEW)

# Cursor sub-accessors (이미 있음)
app.move.doc.top()
app.move.line.end()
...

# 인스턴스 헬퍼 (단일 액션)
app.cell                # CellAccessor (현재)
app.table               # TableAccessor (현재)
app.page                # PageAccessor (현재)

# Actions (Fluent — return self)
app.insert_text(text) -> self
app.styled_text(text, **fmt) -> self
app.insert_heading(text, level) -> self
app.insert_table(...) -> self
app.insert_paragraph_break() -> self
app.find(text) -> bool         # 단순 검색
app.replace(old, new) -> int   # 치환 횟수 반환
app.save(path) -> self
app.open(path) -> self
app.close() -> self

# Context managers
with app.silencing(...) as a: ...
with app.with_charshape(**fmt): ...
with app.with_parashape(**fmt): ...
with app.using(doc): ...

# 저수준 (raw COM)
app.api                 # 그대로 유지

# 폐기
app.charshape(...)              # 빌더 — 제거 (deprecated since v0.0.7)
app.set_visible(...)             # property visible = 로 통일
app.get_text(spos, epos)         # → app.text 또는 app.read_range(...)
app.get_charshape() / set_*      # → app.charshape property
app.create_field / set_field /   # → app.fields accessor
   get_field / delete_field
app.insert_hyperlink              # → app.hyperlinks.add
app.insert_bookmark               # → app.bookmarks.add
app.create_action / parameterset # → app.actions.X 또는 raw
```

### 새 접근자 클래스

```python
class Fields:
    """app.fields — 누름틀(필드) 컬렉션."""
    def __iter__(self) -> Iterator[Field]
    def __getitem__(self, name) -> Field          # app.fields["name"]
    def __setitem__(self, name, value)            # app.fields["name"] = "값"
    def __contains__(self, name) -> bool          # "name" in app.fields
    def __len__(self) -> int
    def add(self, name, *, memo="", direction="") -> Field
    def remove(self, name) -> bool
    def remove_all(self) -> int
    def find(self, name) -> Optional[Field]
    def rename(self, old, new) -> bool
    def from_brackets(self, pattern=...) -> List[Field]   # 기존 replace_brackets

class Field:
    name: str
    value: str   # property (get/set)
    memo: str
    def remove(self) -> None
    def goto(self) -> None

class Bookmarks:
    """app.bookmarks — 책갈피 컬렉션."""
    def add(self, name) -> Bookmark
    def remove(self, name) -> bool
    def find(self, name) -> Optional[Bookmark]
    def goto(self, name) -> bool

class Hyperlinks:
    """app.hyperlinks — 하이퍼링크 컬렉션."""
    def add(self, text, url) -> Hyperlink
    def find(self, text) -> Optional[Hyperlink]
```

### Document

`Document` 의 xlwings-style proxy 는 그대로 유지. App API 가 일관되면 doc.xxx 도 자동으로 일관.

```python
doc.text = "..."           # property set (v0.x 와 동일, 잘 작동)
doc.fields["name"] = "값"
doc.styles["제목 1"].apply()
doc.save() -> doc          # Fluent
doc.close() -> Documents   # 자기 포함 컬렉션 반환
```

### Color / 단위

이미 v0.0.4+ 에서 일관됨 — `Color` / `UNSET` / `from_rgb` / `from_hex`. 유지.

단위 변환:
```python
# v1.0: namespace 통일
import hwpapi.units as U
U.mm(10) -> 2834                   # → HWPUNIT
U.pt(12) -> 1200
U.inch(8.27) -> 59544
U.from_hwpunit(2834, "mm") -> 10.0

# App method 폐기 (mm_to_hwpunit 등) — module-level 함수가 더 적절
```

---

## Migration 전략

### 단계별 deprecation

```
v0.0.10 (현재) ────► v0.1.0 ────► v0.2.0 ────► v1.0
                    경고 추가      경고+migrate   완전 제거
                    호환 유지      도구 제공       breaking
```

### v0.1.0 — Soft migration (호환 유지)

신규 API 추가 + 구 API DeprecationWarning:

- `app.fields = Fields(self)` 추가
- 기존 `app.set_field(...)` 등은 `Fields.__setitem__` 으로 위임 + warning
- `app.get_charshape()` → `app.charshape` property 와 동일 동작 + warning
- 모든 set 메소드 → return self 추가 (return None 코드 영향 없음)
- 새 context manager 명 추가 (`silencing`, `with_charshape`) — 기존도 alias

### v0.2.0 — Migration 도구

- `python -m hwpapi.migrate <file.py>` — AST 변환기로 자동 마이그레이션
- 변환 매트릭스 docs/MIGRATION_v1.md 에 정리

### v1.0 — Breaking 적용

- 모든 deprecated API 제거
- 새 모듈 구조 정립 (`core/ops/*` delegate 분리)
- API 레퍼런스 새로 작성

---

## 문서화

각 단계마다 다음 문서 갱신:

- `CHANGELOG.md` — 신규 + deprecated 모두 명시
- `MIGRATION_v0_x_to_v1.md` — 사용자용 가이드
- `docs/V1_API_REFERENCE.md` — 자동 생성 (현재 02_api_reference.qmd 의 후속)

---

## 결정 필요 사항

1. **Context manager 명명**: 동사형 (`silencing`, `with_charshape`) vs noun_scope (`silence_scope`, `charshape_scope`)?
   → **추천: 동사형** — Pythonic, `with` 키워드와 자연스러움

2. **`charshape` property 동작**:
   - Read: 현재 cursor 의 CharShape 값 (snapshot)
   - Write: `app.charshape = CharShape(...)` 로 전체 교체?
   - Or: `app.charshape.set(bold=True)` 로 partial update?
   → **추천: 둘 다 지원**. read 는 CharShape 객체 반환, partial update 는 `app.charshape.update(bold=True)`.

3. **단위 변환**: instance method 유지 vs module function 만 유지?
   → **추천: module function 만 유지** (`hwpapi.units.mm(10)`). instance method 는 의미없는 wrapping.

4. **`Document` 와 `App` 메소드 분리**:
   - 현재: Document 는 App 메소드를 자동 proxy. 사용자가 `app.xxx()` 와 `doc.xxx()` 어느 쪽이든 사용 가능.
   - v1.0: 동일 (이건 잘 됨)

5. **에러 클래스 도입**:
   - `class HwpError(Exception)` — 모든 HWP 작업 실패의 공통 베이스
   - 서브: `HwpFieldNotFound`, `HwpFileLocked`, `HwpDialogTimeout` 등
   → **추천: 도입**. 명시적 try/except 가능.

---

## 일정 (예상)

| 단계 | 기간 | 내용 |
|:---|:---|:---|
| **v0.1.0** | 4주 | 신규 accessor (`Fields`, `Bookmarks`, `Hyperlinks`) 도입, 기존 API DeprecationWarning |
| **v0.1.x** | 2주 | `core/ops/*` delegate 분리 (REFACTORING_PLAN P1-1) |
| **v0.2.0** | 2주 | `python -m hwpapi.migrate` 도구, MIGRATION 가이드 |
| **v1.0-rc** | 2주 | breaking 적용, 베타 테스트 |
| **v1.0** | release | |

총 약 10주.

---

## 다음 단계 (즉시 진행 가능)

1. **`Fields`, `Bookmarks`, `Hyperlinks` 접근자** 추가 (호환 유지) — v0.0.11
2. **모든 set 메소드에 `return self` 추가** (기존 None 사용처에 영향 없음) — v0.0.11
3. **`charshape` / `parashape` property 추가** (기존 method 와 공존) — v0.0.12
4. 사용 사례 튜토리얼에 새 패턴 노출 — v0.0.12

이 4단계는 **호환성 유지** 하면서 새 패턴을 도입하므로 v1.0 전에 사용자가 체험할 수 있습니다.

---

## 참고 문서

- [REFACTORING_PLAN.md](REFACTORING_PLAN.md) — 구조적 리팩토링 (`app.py` 분할 등)
- [PYHWPX_COMPARISON.md](PYHWPX_COMPARISON.md) — 외부 라이브러리 기능 도입
- [DUPLICATION_AUDIT.md](DUPLICATION_AUDIT.md) — 내부 중복 감사
