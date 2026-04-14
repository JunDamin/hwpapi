# Refactoring Plan

현재 코드베이스 분석 결과를 기반으로 한 리팩토링 계획입니다.

## 진행 상황

| Phase | 상태 | 내용 |
|-------|------|------|
| **1-1** | ✅ 완료 | actions.py 704개 @property → __getattr__ (-2813줄, 73% 감소) |
| **1-2** | ✅ 완료 | SyntaxWarning 14개 수정 (raw string, escape sequences) |
| **2-2** | ✅ 완료 | __getattr__/__setattr__ O(n)→O(1) lookup 테이블 |
| **2-3** | ⏭️ 보류 | Backend 단순화 (하위 호환성 우선으로 유지) |
| **3** | ✅ 완료 | 테스트 커버리지 확대: functions(50), constants(67), classes(18) |
| **2-1** | 📋 연기 | parametersets.py 모듈 분리 (하단 참조) |

전체 테스트: 2171 → **2306 passed** (+135)
코드베이스: 13,046 → **10,233줄** (-2813, -21.6%)

---

## Phase 2-1 연기 사유

`parametersets.py` (4135줄)를 `hwpapi/parametersets/` 패키지로 분리하는 작업은
다음 리스크로 인해 별도 작업으로 연기합니다:

1. **Forward references**: 143개 ParameterSet 서브클래스가 복잡하게 상호 참조
2. **메타클래스 + 자동 등록**: 임포트 순서가 PARAMETERSET_REGISTRY 동작에 민감
3. **백엔드 싸이클**: `hwpapi/actions.py`가 `import hwpapi.parametersets as parametersets`로 속성 접근
4. **실질 이득 제한**: 의미적 변화 없이 주로 IDE 탐색성만 개선

대안: 주요 리팩토링 효과는 Phase 1, 2-2, 3에서 이미 달성함.

---

---

## 현재 상태

| 파일 | 줄 수 | 상태 |
|------|-------|------|
| actions.py | 3,839 | 과대 - 704개 @property 보일러플레이트 |
| parametersets.py | 4,148 | 과대 - 143개 클래스 단일 파일 |
| core.py | 1,734 | 보통 |
| classes.py | 1,713 | 보통 |
| functions.py | 678 | 적정 |
| constants.py | 624 | 적정 |
| logging.py | 303 | 적정 |
| **합계** | **13,046** | |

---

## Phase 1: Quick Wins (예상 2시간)

### 1-1. actions.py: 704개 @property → `__getattr__` 동적 디스패치

**문제:** `_action_info` 딕셔너리에 모든 데이터가 이미 있는데, 704개의 동일한 `@property` 메서드가 중복 존재.

**Before (3,839줄):**
```python
class _Actions:
    @property
    def CharShape(self):
        return _Action(self._app, "CharShape")
    
    @property
    def AllReplace(self):
        return _Action(self._app, "AllReplace")
    # ... 702개 더
```

**After (~1,100줄):**
```python
class _Actions:
    def __getattr__(self, name):
        if name.startswith('_'):
            raise AttributeError(name)
        if name in _action_info:
            return _Action(self._app, name)
        raise AttributeError(f"No action '{name}'")
    
    def __dir__(self):
        return list(_action_info.keys()) + super().__dir__()
```

**효과:** 2,700줄 감소 (71%), 유지보수 간소화, `_action_info`가 단일 진실 소스(SSOT)

---

### 1-2. SyntaxWarning 수정

| 파일 | 줄 | 문제 | 수정 |
|------|---|------|------|
| functions.py:28 | `"\s"` | 이스케이프 시퀀스 | `r"\s"` |
| parametersets.py:3048 | `"\_"` | docstring 이스케이프 | raw string |

---

## Phase 2: 구조 개선 (예상 4시간)

### 2-1. parametersets.py 모듈 분리

**문제:** 4,148줄 단일 파일에 기반 클래스, 백엔드, 디스크립터, 143개 서브클래스 혼재.

**제안 구조:**
```
hwpapi/parametersets/
├── __init__.py          # 전체 re-export (하위호환)
├── base.py              # ParameterSet, ParameterSetMeta, GenericParameterSet
├── backends.py          # PsetBackend, HParamBackend, ComBackend, AttrBackend
├── properties.py        # PropertyDescriptor 및 10개 하위 타입
├── mappings.py          # DIRECTION_MAP, ALIGNMENT_MAP 등 30+개
├── registry.py          # PARAMETERSET_REGISTRY, wrap_parameterset, helpers
├── text.py              # CharShape, ParaShape, BulletShape, NumberingShape 등
├── table.py             # Table, Cell, CellBorderFill, TableCreation 등
├── drawing.py           # ShapeObject, Draw* 시리즈, Caption 등
├── document.py          # DocumentInfo, PageDef, SecDef, SummaryInfo 등
├── formatting.py        # BorderFill, Style, HeaderFooter, FootnoteShape 등
├── actions_psets.py     # FileOpen, FileSaveAs, Print, FindReplace 등
└── runtime.py           # 런타임 발견 클래스 (SelectionOpt, FileOpenSave 등)
```

**효과:**
- 파일당 200~500줄로 분리
- IDE 탐색/자동완성 성능 향상
- `__init__.py`에서 전체 re-export → 기존 import 호환

---

### 2-2. `__getattr__` 성능 최적화

**문제:** snake_case → PascalCase 변환 시 `dir(type(self))` O(n) 순회.

**수정:**
```python
class ParameterSetMeta(type):
    def __new__(mcs, name, bases, namespace):
        cls = super().__new__(mcs, name, bases, namespace)
        # 메타클래스에서 한 번만 빌드
        cls._attr_lookup = {}
        for key in cls._property_registry:
            cls._attr_lookup[key.lower()] = key
            snake = _pascal_to_snake(key)
            cls._attr_lookup[snake] = key
        return cls
```

**효과:** 속성 접근 O(n) → O(1)

---

### 2-3. Backend 단순화

**현재 4개:**
```
PsetBackend      ← action.CreateSet() (주력)
HParamBackend    ← HParameterSet (레거시)
ComBackend       ← 일반 COM (거의 미사용)
AttrBackend      ← Python 객체 (테스트용)
```

**제안:** ComBackend + AttrBackend → `FallbackBackend`로 통합. 실제 프로덕션에서는 PsetBackend과 HParamBackend만 사용.

---

## Phase 3: 테스트 커버리지 확대 (예상 4시간)

### 현재 테스트 커버리지

| 모듈 | 직접 테스트 | 상태 |
|------|-----------|------|
| parametersets.py | test_all_parametersets.py, test_architecture.py | **충분** |
| actions.py | test_all_actions.py | **충분** (HWP 필요) |
| core.py | test_hparam.py (간접) | **부족** |
| classes.py | 없음 | **누락** |
| functions.py | 없음 | **누락** |
| constants.py | 없음 | **누락** |

### 추가 테스트 계획

1. **test_functions.py**: 단위 변환 (`to_hwpunit`, `from_hwpunit`, `mili2unit`), 색상 변환, 경로 처리
2. **test_classes.py**: Accessor 클래스 mock 테스트
3. **test_constants.py**: Enum 값 검증, 폰트 리스트 무결성

---

## Phase 4: 선택적 개선 (필요 시)

### 4-1. ParameterSet 서브클래스 데이터 주도 생성

**제안:** 속성이 적은 클래스(50+개)를 메타데이터로 생성:
```python
# 데이터 정의
SIMPLE_PSETS = {
    'ChCompose': {'Chars': '겹칠 문자'},
    'CodeTable': {'Text': '삽입 문자열'},
    'ConvertCase': {'ConvertStyle': '변환 스타일'},
    ...
}

# 팩토리 생성
for name, props in SIMPLE_PSETS.items():
    globals()[name] = _make_parameterset(name, props)
```

**효과:** ~800줄 감소. 단, 코드 가독성 트레이드오프 존재.

### 4-2. 타입 힌트 강화

현재 `from __future__ import annotations`만 사용. 주요 public API에 타입 힌트 추가:
- `App.__init__`, `App.open`, `App.save` 등
- `ParameterSet.__init__`, `bind`, `apply`
- `_Action.run`, `_Action.set_parameter`

---

## 우선순위 요약

| 순위 | 작업 | 줄 감소 | 난이도 | 효과 |
|------|------|--------|--------|------|
| **1** | actions.py @property → __getattr__ | -2,700 | 낮음 | 즉시 체감 |
| **2** | SyntaxWarning 수정 | 0 | 매우 낮음 | 경고 제거 |
| **3** | parametersets.py 모듈 분리 | 0 (재배치) | 중간 | 유지보수성 |
| **4** | __getattr__ 성능 최적화 | -50 | 낮음 | 런타임 성능 |
| **5** | 테스트 커버리지 확대 | +300 | 중간 | 안정성 |
| **6** | Backend 단순화 | -100 | 중간 | 복잡도 감소 |
| **7** | 데이터 주도 ParameterSet 생성 | -800 | 높음 | 유지보수성 |

---

## 예상 결과

```
현재:  13,046줄, 7개 파일
Phase 1 후: 10,346줄, 7개 파일  (21% 감소)
Phase 2 후: 10,346줄, 15개 파일 (모듈 분리)
Phase 3 후: 10,646줄, 18개 파일 (테스트 추가)
전체 완료:  ~9,000줄, 18개 파일 (31% 감소, 구조 개선)
```
