# hwpapi Refactoring Plan

> **작성일:** 2026-04-15  
> **대상 버전:** v0.0.3 → v0.1.0 → v1.0
> **상태:** 신규 계획 수립 완료 — Phase 1 실행 대기

이 문서는 hwpapi 패키지의 전반적 리팩토링 계획서입니다. 최근 확장 (Document/Documents xlwings-style proxy, styled_text/charshape_scope/parashape_scope, 표준 메소드들, 자동 생성 API 레퍼런스) 이후 구조적으로 정리할 부분을 체계화했습니다.

---

## 목차

1. [과거 완료 (이미 처리된 항목)](#과거-완료)
2. [현황 진단](#현황-진단)
3. [우선순위별 과제](#우선순위별-과제)
   - [P0 — 긴급](#p0--긴급)
   - [P1 — 구조적 문제](#p1--구조적-문제)
   - [P2 — 코드 품질 · 일관성](#p2--코드-품질--일관성)
   - [P3 — DX · 문서화](#p3--dx--문서화)
4. [단계별 실행 계획](#단계별-실행-계획)
5. [마이그레이션 전략](#마이그레이션-전략)
6. [성공 지표](#성공-지표)

---

## 과거 완료

| Phase | 내용 |
|:---|:---|
| **1-1** | actions.py 704개 `@property` → `__getattr__` (-2,813줄, 73% 감소) |
| **1-2** | SyntaxWarning 14개 수정 (raw string, escape sequences) |
| **1-3** | parametersets.py (4,198줄) → parametersets/ 패키지 (15개 파일) |
| **1-4** | core.py 분할: `core/engine.py` + `core/app.py` |
| **1-5** | classes.py 분할: `classes/accessors.py` + `classes/shapes.py` |
| **2-1** | Document / Documents / use_document — 다중 문서 관리 추가 |
| **2-2** | Document.__getattr__ 기반 xlwings-style proxy |
| **2-3** | styled_text / charshape_scope / parashape_scope — snapshot + restore |
| **2-4** | 표준 word-processor 메소드들 (insert_heading, insert_table, insert_hyperlink, undo/redo/copy/paste, `text`/`visible`/`version`/`page_count` properties) |
| **2-5** | 자동 생성 API 레퍼런스 페이지 (277K chars, 6,676줄) |
| **3-1** | Quarto 기반 홈페이지 재설계 + 8개 튜토리얼 |
| **3-2** | PyPI Trusted Publisher 설정 — GitHub Release 자동 배포 |

---

## 현황 진단

### 정량 지표 (2026-04-15)

| 영역 | 현재 상태 | 판정 |
|:---|:---|:---:|
| `hwpapi/core/app.py` | **~1,800줄** (Document + Documents + App 통합) | 🟠 비대 |
| `hwpapi/core/engine.py` | ~385줄 | ✅ |
| `hwpapi/actions.py` | ~1,000줄, `_Action` 캐싱 버그 있음 | 🔴 버그 |
| `hwpapi/parametersets/` | 15 파일로 분리 완료 (~4,500줄) | ✅ |
| `hwpapi/classes/` | accessors.py (860) + shapes.py (922) | 🟠 CharShape 중복 |
| 단위 테스트 | 1,038개 통과 (hparam + architecture) | 🟠 커버리지 확대 필요 |
| E2E 테스트 | smoke_scenarios 13 + smoke_features 18 | ✅ |
| 문서 | Quarto 웹사이트 + 8 튜토리얼 + 전체 API 레퍼런스 | ✅ |
| PyPI | v0.0.3 업로드됨 (Trusted Publisher) | ✅ |

### 발견된 핵심 문제점

1. **🔴 `_Action` 문서 바인딩 버그** — 캐시된 `self.act` 가 초기 활성 문서에 고정. `insert_text` 는 수동 패치됐으나 `app.actions.CharShape.run()` 경로는 여전히 취약.

2. **🟠 `app.py` 비대화** — 1,800줄로 IDE 네비게이션 저하, 리뷰 어려움.

3. **🟠 `ColorProperty` 의미론 모호** — `shade_color="#FFFFFF"` 가 "제거" 인지 "흰색 shade 적용" 인지 불명확. snapshot/restore 로 대증 치료했지만 근본 해결 아님.

4. **🟡 API 일관성 부족** — 반환값이 `None` / `bool` / `str` / `int` / `Document` / `pset` 등으로 제각각.

5. **🟡 Docstring 스타일 혼재** — NumPy/Google 스타일 섞임, 한글/영문 혼재.

6. **🟡 `CharShape` 중복** — `hwpapi.parametersets.sets.primitives.CharShape` (ParameterSet) 과 `hwpapi.classes.shapes.CharShape` (dataclass) 가 동시 존재하여 import 혼란.

---

## 우선순위별 과제

### P0 — 긴급

#### P0-1. `_Action` 을 lazy-eval 로 변경

**문제:**
```python
# hwpapi/actions.py — 현재
class _Action:
    def __init__(self, app, action_key):
        self.act = app.api.CreateAction(action_key)   # ← 초기 활성 doc 에 바인딩
        self.pset = self._wrap_pset(self.act.CreateSet())

# 결과
doc2.activate()                  # 활성 문서 전환
app.actions.CharShape.run()      # ⚠️ doc1 에 적용됨
```

**해결:** 문서별로 action 을 lazy-cache.
```python
class _Action:
    def __init__(self, app, action_key):
        self.app = app
        self.action_key = action_key
        self._act_cache = {}     # {doc_id: action}
        self._pset_cache = {}    # {doc_id: wrapped_pset}

    def _current_doc_id(self):
        try:
            return self.app.api.XHwpDocuments.Active_XHwpDocument.DocumentID
        except Exception:
            return 0

    @property
    def act(self):
        did = self._current_doc_id()
        if did not in self._act_cache:
            self._act_cache[did] = self.app.api.CreateAction(self.action_key)
        return self._act_cache[did]

    @property
    def pset(self):
        did = self._current_doc_id()
        if did not in self._pset_cache:
            raw = self.act.CreateSet()
            self.act.GetDefault(raw)
            self._pset_cache[did] = self._wrap_pset(raw)
        return self._pset_cache[did]
```

| 항목 | 값 |
|:---|:---|
| 규모 | M (~100줄 수정 + 5개 단위 테스트) |
| 위험도 | 중 (704 액션의 내부 행동 변경) |
| 호환성 | Public API 유지 |
| 검증 | 새 test `tests/integration/test_action_binding.py` — 문서 2개에서 각자 CharShape 실행 후 내용 분리 확인 |

---

#### P0-2. `Color` 래퍼 + `UNSET` 센티넬 도입

**문제:** `shade_color=None` 이 "필드 제거" 인지 "기본값 리셋" 인지 모호. `"#FFFFFF"` 설정은 "흰색 shade" 를 실제로 적용해버림.

**해결:**
```python
# hwpapi/parametersets/properties.py
class _Unset:
    """'이 필드를 손대지 마라' 센티넬."""
    def __repr__(self): return "UNSET"
    def __bool__(self): return False

UNSET = _Unset()

class Color:
    """HWP 색상 래퍼. None/UNSET 의미 명확."""
    __slots__ = ('_hwp_value',)
    def __init__(self, value):
        if value is None or value is UNSET:
            self._hwp_value = None
        else:
            self._hwp_value = convert_to_hwp_color(value)
    @property
    def hex(self):
        return convert_hwp_color_to_hex(self._hwp_value) if self._hwp_value is not None else None
    def __str__(self): return self.hex or ""
    def __eq__(self, other):
        if isinstance(other, Color): return self._hwp_value == other._hwp_value
        if other is None: return self._hwp_value is None
        return self == Color(other)
    def __repr__(self):
        return f"Color({self.hex!r})" if self._hwp_value is not None else "Color(UNSET)"

class ColorProperty(PropertyDescriptor):
    def __get__(self, instance, owner):
        if instance is None: return self
        return Color(self._get_value(instance))

    def __set__(self, instance, value):
        if value is UNSET: return
        if value is None: return self._del_value(instance)
        color = value if isinstance(value, Color) else Color(value)
        if color._hwp_value is None: return self._del_value(instance)
        return self._set_value(instance, color._hwp_value)
```

사용:
```python
app.set_charshape(shade_color="#FFFF00")  # 기존 동일, 형광펜
app.set_charshape(shade_color=UNSET)       # 건드리지 않음 (의도적 무시)
app.set_charshape(shade_color=None)        # 필드 제거 (snapshot 복원용)
cs.shade_color                             # → Color 객체 반환
str(cs.shade_color)                        # → "#ffff00"
```

| 항목 | 값 |
|:---|:---|
| 규모 | M (~200줄 + 테스트) |
| 위험도 | 중 (반환 타입이 str → Color 변경, `__str__`/`__eq__` 로 완화) |
| 검증 | `tests/unit/test_color_semantics.py` — `Color == "#FFFF00"`, `Color in (None, 0)` 등 edge case |

---

### P1 — 구조적 문제

#### P1-1. `app.py` 분할 — delegate 패턴

**현재:** `App` 클래스 하나에 File I/O · Text · Formatting · Table · Page · Document · Visibility 7개 관심사 집중 (~1,800줄).

**해결:** 이미 작동 중인 accessor 패턴을 internal delegate 로 확장.
```
hwpapi/core/
├── __init__.py            # re-exports (backward compat)
├── engine.py              # (유지)
├── app.py                 # App 라우팅만 (~500줄)
├── document.py            # Document + Documents + use_document (~400줄)
└── ops/                   # ← 신규
    ├── __init__.py
    ├── text_ops.py        # insert_text, styled_text, find/replace, get_text
    ├── format_ops.py      # set/get_charshape + parashape + scopes + snapshots
    ├── file_ops.py        # open, save, save_block, insert_file, close, quit
    ├── structure_ops.py   # insert_heading, insert_table, insert_hyperlink, bookmark
    └── clipboard_ops.py   # cut/copy/paste/undo/redo/select_all/clear
```

App 구조:
```python
class App:
    def __init__(self, ...):
        # 기존 engine/actions/accessors
        self._text = TextOps(self)
        self._format = FormatOps(self)
        self._file = FileOps(self)
        self._structure = StructureOps(self)
        self._clipboard = ClipboardOps(self)

    # 얇은 라우팅 (public API 완전 동일)
    def insert_text(self, text, **kw):  return self._text.insert_text(text, **kw)
    def styled_text(self, text, **f):   return self._format.styled_text(text, **f)
    ...
```

| 항목 | 값 |
|:---|:---|
| 규모 | L (1,800줄 재배치 → 500 + 400 + 5×200) |
| 위험도 | 중 (public API 변경 없음) |
| 영향 | 사용자 코드 무영향 |
| 우선 순서 | `document.py` 먼저 분리 → `ops/` 순차 |

---

#### P1-2. `CharShape` 중복 제거

| 위치 | 종류 |
|:---|:---|
| `hwpapi.parametersets.sets.primitives.CharShape` | ParameterSet 래퍼 (쓰기 가능, COM) |
| `hwpapi.classes.shapes.CharShape` | 독립 dataclass (읽기 전용 snapshot) |

같은 이름 → import 혼란 + 기능 중복.

**해결:**
- `classes.shapes.CharShape` → `CharShapeView` 로 rename
- 1 버전 동안 `CharShape` alias + `DeprecationWarning`

| 항목 | 값 |
|:---|:---|
| 규모 | S (~50줄) |
| 위험도 | 중 |

---

#### P1-3. Accessor sub-grouping (MoveAccessor 38개, TableAccessor 30개가 flat)

```python
# 기존
app.move.top_of_file()
app.move.end_of_line()
app.move.start_of_word()

# 제안
app.move.doc.top()
app.move.line.end()
app.move.word.start()
```

Deprecated alias 를 1 버전 유지.

| 항목 | 값 |
|:---|:---|
| 규모 | M |
| 위험도 | 중 |

---

### P2 — 코드 품질 · 일관성

#### P2-1. 반환값 표준화 — Fluent API 도입

**옵션:**
- **A**: `ActionResult` dataclass — 보수적
- **B**: `self` 반환 (xlwings-style chaining) — 권장

```python
# 옵션 B
class App:
    def insert_text(self, text, **kw) -> "App":
        ...
        return self

    def styled_text(self, text, **fmt) -> "App":
        ...
        return self

# 사용
app.insert_text("앞 — ").styled_text("중요", bold=True).insert_text(" — 뒤").save("out.hwp")
```

값 반환이 필수인 메소드 (`find_text` → bool, `get_text` → str, `replace_all` → int) 는 예외.

| 항목 | 값 |
|:---|:---|
| 규모 | M |
| 위험도 | **높** (breaking) |
| 전략 | v0.1.0 에서 새 메소드만 `-> self`, 기존은 유지. v1.0 에서 전환 |

---

#### P2-2. Type Hints 완성

**현재:** 공개 메소드 ~60% 에 hint, `**kwargs: Any` 남발.

**해결:** `TypedDict` 로 kwargs 명세.
```python
from typing import TypedDict, Unpack

class CharShapeKwargs(TypedDict, total=False):
    bold: bool
    italic: bool
    underline_type: int
    strike_out_type: int
    super_script: int
    sub_script: int
    text_color: str
    shade_color: str
    height: int
    spacing_hangul: int
    # ...

class App:
    def set_charshape(
        self,
        charshape: Optional[CharShape] = None,
        **kw: Unpack[CharShapeKwargs],
    ) -> None: ...
```

| 항목 | 값 |
|:---|:---|
| 규모 | M |
| 위험도 | 낮 |
| 도구 | `mypy --strict` 통과 |

---

#### P2-3. 테스트 재구조화

**권장:**
```
tests/
├── conftest.py               # shared fixtures, markers
├── unit/                     # HWP 불필요
│   ├── test_properties.py
│   ├── test_parameterset.py
│   ├── test_action_binding.py    # P0-1 회귀
│   ├── test_color_semantics.py   # P0-2 회귀
│   ├── test_document_proxy.py
│   └── test_constants.py
├── integration/              # @pytest.mark.hwp
│   ├── test_text_ops.py
│   ├── test_format_ops.py
│   ├── test_multi_document.py
│   └── test_scope_managers.py
└── e2e/                      # 전체 시나리오 (현재 smoke_*)
    ├── test_scenarios.py
    └── test_features.py
```

`pytest -m "not hwp"` → CI 에서 unit 만 실행.

---

### P3 — DX · 문서화

#### P3-1. Docstring 표준화

Google-style + 한글.
```python
def method(self, arg: Type) -> ReturnType:
    """한 줄 요약 (마침표 포함).

    필요시 상세 설명.

    Args:
        arg: 설명.

    Returns:
        반환값 설명.

    Examples:
        >>> app.method(123)
        result

    Raises:
        ValueError: 조건.
    """
```

#### P3-2. CLAUDE.md 갱신

Known Issues 섹션에 P0-1, P0-2 기록 + 최근 확장 (styled_text, charshape_scope, Document proxy) 반영.

#### P3-3. CHANGELOG.md 도입

Keep a Changelog 포맷 + 각 Phase 릴리즈 기록.

#### P3-4. API 레퍼런스 개선

- 자동 생성 페이지에 "Deprecated API" 섹션
- 각 메소드에 "Since v0.X.X" 표시
- 검색 인덱스 강화

---

## 단계별 실행 계획

### Phase 1 — 긴급 안정화 (2주, v0.0.4)

**목표:** P0 버그 제거 + 회귀 테스트.

- [ ] P0-1: `_Action` lazy binding (M, 1주)
- [ ] P0-2: `Color` / `UNSET` semantics (M, 5일)
- [ ] CLAUDE.md Known Issues 갱신
- [ ] v0.0.4 tag → trusted publisher 가 자동 PyPI 업로드

### Phase 2 — 구조 재편 (4~6주, v0.1.0)

- [ ] P1-1: `app.py` 분할 (L, 2주)
- [ ] P1-2: CharShape 중복 제거 (S, 3일)
- [ ] P2-1: Fluent API 도입 — 새 메소드만 (M, 1주)
- [ ] P2-3: 테스트 재구조화 (M, 1주)
- [ ] v0.1.0 릴리즈

### Phase 3 — 정리 · 문서 (1~2주, v0.1.1)

- [ ] P1-3: Accessor sub-grouping (M, 1주)
- [ ] P2-2: Type hints 완성 (M, 1주)
- [ ] P3-1: Docstring 표준화 (점진적)
- [ ] P3-2: CLAUDE.md 전체 갱신
- [ ] P3-3: CHANGELOG.md 작성
- [ ] P3-4: API 레퍼런스 개선
- [ ] v0.1.1 릴리즈

### v1.0 (향후)

- [ ] 반환값 완전 통일 (Fluent)
- [ ] Deprecated API 제거
- [ ] `MIGRATION_v1.md` 작성

---

## 마이그레이션 전략

### Deprecation 정책

1. **v0.0.x → v0.1.0**: 변경 시 `DeprecationWarning` + alias 유지
2. **v0.1.0 → v1.0**: Deprecated 제거, 반환값 통일

### 버전별 호환성 매트릭스

| 버전 | Python | 주요 변경 | 호환성 |
|:---|:---:|:---|:---:|
| v0.0.3 (현재) | 3.7+ | Document proxy, 표준 메소드, 31 E2E 테스트 | — |
| v0.0.4 | 3.7+ | _Action 버그 수정, Color 클래스 | ✅ 100% |
| v0.1.0 | 3.8+ | app.py 분할, accessor 재그룹 | ✅ Public API 동일 |
| v0.1.1 | 3.8+ | Type hints, docstring, test 재편 | ✅ 100% |
| v1.0 | 3.10+ | 반환값 통일 (Fluent), deprecated 제거 | ⚠️ Breaking (migration guide) |

### 사용자 공지 체계

- PyPI Release Notes + GitHub Release 에 각 Phase 주요 변경점
- v0.1.0 → `MIGRATION_v0_1.md`
- v1.0 → `MIGRATION_v1.md`
- 비공개 API (`_Action`, `_Actions._cache` 등) 는 변경 자유

---

## 성공 지표

### 코드 품질
- [ ] `hwpapi/core/app.py` < 500줄 (현재 ~1,800)
- [ ] Type hint 커버리지 95%+ (`mypy --strict` 통과)
- [ ] Docstring 표준 100%
- [ ] 파일당 평균 < 300줄

### 안정성
- [ ] 단위 테스트 1,200+ (현재 1,038)
- [ ] 통합 테스트 50+ 시나리오
- [ ] `_Action` 바인딩 회귀 테스트 5+
- [ ] 다중 문서 상호 간섭 테스트

### 사용자 경험
- [ ] API 레퍼런스 6,600+ 줄 유지 + Deprecated 섹션
- [ ] 튜토리얼 8개 + migration guide
- [ ] `pip install hwpapi` → hello world 5분 이내

### 호환성
- [ ] v0.0.4 / v0.1.0 → 기존 코드 수정 0
- [ ] v1.0 → migration guide 100% 커버

---

## 참고

- 현재 구조: [`CLAUDE.md`](../CLAUDE.md)
- ParameterSet 아키텍처: [`docs/PARAMETERSET_ARCHITECTURE.md`](PARAMETERSET_ARCHITECTURE.md)
- 이전 리팩토링: [`REFACTORING_SUMMARY.md`](../REFACTORING_SUMMARY.md), [`PSET_MIGRATION_SUMMARY.md`](../PSET_MIGRATION_SUMMARY.md)
- 라이브 API: https://jundamin.github.io/hwpapi/02_api_reference.html
- GitHub: https://github.com/JunDamin/hwpapi
