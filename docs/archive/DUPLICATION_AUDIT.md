# 내부 API 중복·겹침 감사

> 작성일: 2026-04-15 (v0.0.6 기준)
> 목적: Phase D/E 확장 전에 **기존 중복/별칭/기능 분산** 을 정리

## 수치

| 클래스 | 멤버 수 | 메소드 | 속성 |
|:---|---:|---:|---:|
| `App` | **82** | 72 | 10 |
| `Document` | 15 | 7 | 8 |
| `Documents` | 6 | 5 | 1 |
| `MoveAccessor` | 38 | 38 | 0 |
| `TableAccessor` | 30 | 0 | 30 |
| `PageAccessor` | 12 | 0 | 12 |
| `CellAccessor` | 2 | 0 | 2 |

App 82개는 **단일 클래스로 과부하**. Phase D/E 추가 전에 정리가 필요함.

---

## 🔴 P0 — 제거/통합 권장 (중대 중복)

### 1. `app.charshape(facename=..., ..., bold=..., ...)` — **legacy 빌더, 제거 권장**

**정체**:
```python
def charshape(self, facename=None, fonttype=None, ..., bold=None, italic=None, ...):
    """return null charshape"""   # ← 실제 docstring 그대로
    charshape_pset = self.create_parameterset("CharShape")
    # value_set 만들고 setattr 루프 돌려서 pset 돌려줌
    return charshape_pset
```

- 29개 파라미터를 모두 keyword 로 받아서 **unbound** `CharShape` pset 을 반환
- 쓰려면 결과를 다시 `app.set_charshape(result)` 로 전달해야 함
- 같은 일을 `app.set_charshape(bold=True, ...)` 한 줄로 가능
- docstring 은 "return null charshape" 뿐 — 설명 부실

**사용 흐름 비교**:
```python
# charshape() 사용 — 2줄 + 간접
cs = app.charshape(bold=True, italic=True)
app.set_charshape(cs)

# 없으면 — 1줄 직접
app.set_charshape(bold=True, italic=True)
```

**권장 조치**: **제거** (or deprecation warning + v0.2.0 제거).

---

### 2. 서식 스코핑 3중 중복 — 정리 필요

현재 서식을 "그 범위에만" 적용하는 방법이 **3가지**:

| API | 사용 시나리오 | 동작 |
|:---|:---|:---|
| `app.styled_text(text, **fmt)` | 한 텍스트 조각 | 삽입 → 선택 → 포맷 → 복원 |
| `with app.charshape_scope(**fmt):` | 여러 줄 블록 | 진입 snapshot → 종료 시 범위 선택 + 포맷 + 복원 |
| `with app.parashape_scope(**fmt):` | ParaShape 블록 | 동일 (para) |

**평가**: 세 API 는 **서로 다른 사용 패턴** 을 다루므로 중복이 아님. 단, 사용자가 혼동하지 않도록 **하나의 튜토리얼 섹션에 비교표** 를 명시해야 함 (현재 01_app_basics §6-2 에 있음 ✓).

**권장**: 유지. 정리 (rename) 불필요.

---

### 3. `selection` property vs `get_selected_text()` — 별칭

```python
app.selection           # property → str
app.get_selected_text() # method → str
```

둘 다 같은 문자열 반환. `selection` 은 의도적 alias (짧고 Pythonic).

**권장**: 양쪽 유지. `selection` 을 문서상 **canonical** 로 표기.

---

### 4. `visible` property vs `set_visible()` method — 양립

```python
app.visible = True                       # property setter
app.set_visible(is_visible=True)          # legacy method
```

`set_visible` 은 `window_i` 파라미터까지 받음 — 단일 window 외 사용자는 이 쪽을 써야 함.

**권장**: 양쪽 유지 (`visible` 은 간편 단축, `set_visible` 은 고급).

---

### 5. `new_document()` vs `documents.add()` — 명시적 별칭

```python
app.new_document()       # 편의
app.documents.add()      # 정식 경로
```

docstring 에 "alias" 명시됨. 의도적 편의 단축.

**권장**: 유지.

---

### 6. `rgb_color(r,g,b)` vs `Color.from_rgb(r,g,b)` — 반환 타입이 다름

```python
app.rgb_color(255, 0, 0)      # → int 255 (HWP BBGGRR)
Color.from_rgb(255, 0, 0)     # → Color 객체 (raw=255)
```

반환이 다르므로 엄밀히는 중복이 아니지만 **두 가지를 기억해야 함** 이 부담.

**권장**: 
- `app.rgb_color` 을 `Color.from_rgb(r,g,b).raw` 의 shortcut 으로 재정의하면 의미 일원화 가능
- 또는 `app.rgb_color(r,g,b)` 를 deprecated 로 표시하고 `Color.from_rgb()` 를 canonical 로

현 시점 결정: **유지**, 단 docstring 에 "returns int — use Color.from_rgb() for type-safe version" 명시.

---

### 7. 단위 변환 — 3곳에 동일 기능 분산

```python
# 1) App instance methods (Phase C 에서 추가)
app.mm_to_hwpunit(10)
app.point_to_hwpunit(12)

# 2) hwpapi.functions 의 module-level 함수
from hwpapi.functions import to_hwpunit, from_hwpunit, mili2unit, point2unit
to_hwpunit("210mm")

# 3) UnitProperty 내부에서도 사용
```

**평가**: App method (2) 는 이미 내부적으로 `self.api.MiliToHwpUnit` 호출. functions.py 의 함수 (1) 는 HWP 없이도 동작하는 순수 파이썬 유틸리티. 둘의 역할 구분 명확 (instance vs module).

**권장**:
- App method 는 유지 (COM 호출 필요한 상황에 편리)
- functions.py 유틸리티도 유지 (pure Python)
- 다만 문서에서 **"둘 중 뭐 쓰지?"** 가이드 필요

---

## 🟡 P1 — Action-method 이중화 (의도된 래퍼)

아래는 **같은 HWP 기능을 메소드 + action 두 경로로 접근** 가능한 케이스. 모두 **의도된 디자인** 이지만 사용자가 어느 쪽을 써야 하는지 명시해야 함.

| 편의 메소드 | Action 경로 | 차이 |
|:---|:---|:---|
| `app.insert_text("x")` | `app.actions.InsertText.run()` | 메소드가 얇은 래퍼 |
| `app.find_text("x")` | `app.actions.ForwardFind.run()` | 메소드 = `RepeatFind` 래퍼 |
| `app.replace_all("a","b")` | `app.actions.AllReplace.run()` | 메소드가 pset 설정 포함 |
| `app.insert_table(...)` | `app.actions.TableCreate.run()` | 메소드 = 헬퍼 (pandas 지원) |
| `app.insert_hyperlink(text,url)` | `app.actions.InsertHyperlink.run()` | 메소드 = pset 자동 |
| `app.insert_bookmark(name)` | `app.actions.Bookmark.run()` | 메소드 = pset 자동 |
| `app.select_all()` | `app.actions.SelectAll.run()` | `api.Run("SelectAll")` |
| `app.copy() / paste() / cut()` | `app.actions.Copy/Paste/Cut.run()` | 동일 |
| `app.undo() / redo()` | `app.actions.Undo/Redo.run()` | 동일 |

**권장**: 모두 유지 — 편의 메소드는 90% 사용자, action 경로는 고급 제어용.

---

## 🟢 P2 — 구조적 중복 (장기 과제)

### 8. `CharShape` / `ParaShape` 두 군데 정의 (v0.1.x)

- `hwpapi.parametersets.sets.primitives.CharShape` — COM 래퍼 (쓰기 가능)
- `hwpapi.classes.shapes.CharShape` — 독립 dataclass

이미 `REFACTORING_PLAN.md` P1-2 로 기록됨. v0.1.x 에서 `CharShapeView` 로 rename.

### 9. `app.py` 비대 — 82 개 멤버

Phase B/C 에서 Field API(9) + 클립보드(8) + pandas(2) + 단위(5) + 대화상자(3) + 페이지 이미지(2) + 하이라이트(1) + goto(1) 등 **~30개** 추가해 더 비대해짐.

이미 `REFACTORING_PLAN.md` P1-1 로 기록 (`core/ops/*_ops.py` delegate 분리).

---

## 📋 정리 실행 계획 — v0.0.7

### 🔴 P0 — 당장 (v0.0.7)

1. **`app.charshape()` 빌더 제거** (정확히는 deprecation)
   - DeprecationWarning 추가 + 문서에서 빼기
   - v0.1.0 에서 완전 제거

2. **Docstring 품질 개선** — `app.charshape` 의 "return null charshape" 같은 설명 부실 수정 (실제로는 유지는 안 하고 deprecate 이므로 이 항목은 자동 해결)

3. **선택/가시성/문서 추가 별칭** docstring 에 cross-reference 추가
   - `app.selection` docstring → "getter alias for get_selected_text()"
   - `app.visible` docstring → "property form of set_visible(); prefer this"
   - `app.new_document()` → "alias for documents.add()"
   - `app.rgb_color()` → "int shortcut; Color.from_rgb() for type-safe Color"

### 🟡 P1 — 이후 (v0.0.8 이상)

4. **Action-method 문서화 가이드** — "언제 메소드, 언제 action" 튜토리얼 섹션 추가

5. **Move/Table accessor 정리** — pyhwpx 참고, 38개 flat 을 `move.line / move.word / move.para / move.doc / move.cell` 등 sub-grouping (v0.1.x REFACTORING_PLAN P1-3)

### 🟢 P2 — 장기

6. **CharShape 중복 제거** (REFACTORING_PLAN P1-2)
7. **`app.py` 분할** (REFACTORING_PLAN P1-1, `core/ops/`)

---

## 🎯 v0.0.7 — 이번 정리 범위 (backward-compat)

- [ ] `app.charshape()` 빌더 → `DeprecationWarning` 추가
- [ ] docstring cross-reference 6개 (selection, visible, new_document, rgb_color, set_visible 등)
- [ ] `hwpapi.functions.to_hwpunit` vs `app.mm_to_hwpunit` 가이드 추가
- [ ] API 레퍼런스 재생성 (새 경고 반영)
- [ ] 버전 v0.0.7

v0.0.7 는 **단 5개 커밋**, **zero breaking change** 로 끝낼 수 있는 작은 정리 릴리즈.

이후 Phase D/E (`app.styles`, `app.controls` 접근자) 를 v0.0.8~v0.0.9 로 이어가면 됩니다.

---

## 참고 문서

- [`REFACTORING_PLAN.md`](REFACTORING_PLAN.md) — 중장기 구조 리팩토링
- [`PYHWPX_COMPARISON.md`](PYHWPX_COMPARISON.md) — 외부 기능 도입 로드맵
- [`CHANGELOG.md`](../CHANGELOG.md) — 변경 이력
