# ParameterSet Architecture

hwpapi의 ParameterSet 시스템은 HWP COM API의 파라미터를 Pythonic하게 감싸는 계층 구조입니다.

---

## 전체 구조

```
PARAMETERSET_REGISTRY (글로벌 딕셔너리)
│
├── ParameterSetMeta (메타클래스)
│   └── 클래스 정의 시 _property_registry 자동 수집, REGISTRY 자동 등록
│
├── ParameterSet (기반 클래스)
│   ├── _raw: 원본 COM 객체
│   ├── _backend: ParameterBackend 인스턴스
│   ├── _staged: dict (변경사항 스테이징)
│   ├── _snapshot: dict (원격 값 스냅샷)
│   └── _property_registry: dict (속성명 → PropertyDescriptor)
│
├── Backend 시스템 (COM 객체 추상화)
│   ├── PsetBackend     ← action.CreateSet() 결과 (직접 읽기/쓰기)
│   ├── HParamBackend   ← HParameterSet (점 경로 탐색)
│   ├── ComBackend      ← 일반 COM 객체
│   └── AttrBackend     ← 일반 Python 객체
│
└── Property Descriptor 시스템 (타입별 자동 변환)
    ├── PropertyDescriptor (기반)
    ├── IntProperty      (정수, 범위 검증)
    ├── BoolProperty     (bool ↔ 0/1)
    ├── StringProperty   (문자열)
    ├── ColorProperty    (BBGGRR ↔ #RRGGBB)
    ├── UnitProperty     (HWPUNIT ↔ mm/cm/in/pt)
    ├── MappedProperty   (문자열 ↔ 정수 매핑)
    ├── TypedProperty    (중첩 ParameterSet)
    ├── NestedProperty   (자동 생성 중첩)
    ├── ArrayProperty    (HArray 래퍼)
    └── ListProperty     (Python 리스트)
```

---

## 1. 메타클래스: ParameterSetMeta

`parametersets.py:1238`

ParameterSet 서브클래스가 **정의될 때** 자동으로:
1. 클래스 바디에서 `PropertyDescriptor` 인스턴스를 수집 → `_property_registry`에 저장
2. MRO(Method Resolution Order)를 따라 부모 클래스 속성도 병합
3. `PARAMETERSET_REGISTRY`에 등록 (PascalCase + lowercase 키)

```python
class CharShape(ParameterSet):
    Bold = BoolProperty("Bold", "Bold formatting")
    # → _property_registry = {"Bold": <BoolProperty>}
    # → PARAMETERSET_REGISTRY["CharShape"] = CharShape
    # → PARAMETERSET_REGISTRY["charshape"] = CharShape
```

---

## 2. Backend 시스템

`parametersets.py:153-424`

### Backend 프로토콜
```
get(key) → Any          값 읽기
set(key, value) → None  값 쓰기
delete(key) → bool      값 삭제
```

### Backend 선택 우선순위 (make_backend)

| 우선순위 | 조건 | Backend | 사용 사례 |
|---------|------|---------|----------|
| 1 | `_looks_like_pset(obj)` | PsetBackend | `action.CreateSet()` 결과 |
| 2 | `_is_com(obj)` | ComBackend | 일반 COM 객체 |
| 3 | 기타 | AttrBackend | Python 객체 |

### PsetBackend (주력, 즉시 모드)
- `pset.Item(key)` / `pset.SetItem(key, value)` 사용
- **즉시 반영**: set 호출 시 COM 객체에 바로 기록
- `create_itemset(key, setid)`: 중첩 ParameterSet 생성

### HParamBackend (레거시, 스테이징 모드)
- 점(`.`) 경로로 속성 탐색: `"HFindReplace.FindString"`
- `_resolve_parent_and_leaf()`: 부모 객체와 최종 속성명 분리
- `_coerce_for_put()`: 기존 타입에 맞게 값 변환

### 즉시 모드 vs 스테이징 모드

```python
# PsetBackend (즉시): set 호출 시 COM에 바로 기록
pset.Bold = True  # → backend.set("Bold", 1) → COM.SetItem("Bold", 1)

# HParamBackend/AttrBackend (스테이징): _staged에 저장 후 apply()로 일괄 적용
pset.Bold = True  # → _staged["Bold"] = 1
pset.apply()      # → backend.set("Bold", 1) for all staged
```

---

## 3. Property Descriptor 시스템

`parametersets.py:531-1235`

### 데이터 흐름

```
사용자 코드                  PropertyDescriptor           Backend           COM
───────────                ────────────────           ───────           ───
pset.bold = True     →     BoolProperty.__set__  →    _staged["Bold"]=1
                           (True → 1 변환)
                           
val = pset.bold      ←     BoolProperty.__get__  ←    _staged or backend.get("Bold")
                           (1 → True 변환)
```

### Property 타입별 변환 규칙

| Property | Python 입력 | COM 저장값 | Python 출력 |
|----------|-----------|-----------|-----------|
| IntProperty | `42` | `42` | `42` |
| BoolProperty | `True` | `1` | `True` |
| StringProperty | `"hello"` | `"hello"` | `"hello"` |
| ColorProperty | `"#FF0000"` | `0x0000FF` (BBGGRR) | `"#ff0000"` |
| UnitProperty | `"210mm"` 또는 `210` | `59430` (HWPUNIT) | `210.0` |
| MappedProperty | `"center"` | `1` | `"center"` |

### NestedProperty (자동 생성)
```python
class FindReplace(ParameterSet):
    find_char_shape = NestedProperty("FindCharShape", "CharShape", CharShape)

# 사용:
pset.find_char_shape.bold = True  # CreateItemSet 자동 호출!
```

### ArrayProperty (리스트 인터페이스)
```python
class TabDef(ParameterSet):
    tab_stops = ArrayProperty("TabStops", int, "Tab positions")

# 사용:
pset.tab_stops = [1000, 2000, 3000]
pset.tab_stops.append(4000)
```

---

## 4. ParameterSet 생명주기

```
1. 생성        ps = CharShape()           # unbound (backend=None)
2. 바인딩      ps = CharShape(pset_com)   # bound (backend=PsetBackend)
3. 값 설정     ps.Bold = True             # PsetBackend: 즉시 / 기타: _staged
4. 값 읽기     val = ps.Bold              # _staged → _snapshot → backend.get()
5. 적용        ps.apply()                 # _staged → backend (스테이징 모드만)
6. 리로드      ps.reload()                # backend → _snapshot (캐시 갱신)
```

### snake_case ↔ PascalCase 자동 변환
```python
ps.bold = True       # → __setattr__("bold") → PascalCase 변환 → "Bold" property
val = ps.Bold        # → 직접 PascalCase 접근
```

---

## 5. 표현 시스템 (__repr__)

`parametersets.py:1898-2048`

```python
print(pset)
# CharShape(
#   Bold=True  # Bold formatting
#   FontSize=12.0pt  # Font size (HWPUNIT, 100=1pt)
#   TextColor="#ff0000"  # Text color (BBGGRR)
#   Align=1 (center)  # Text alignment
# )
```

| 값 유형 | 원시값 | 표시 형식 |
|--------|--------|---------|
| 색상 | `0x0000FF` | `"#ff0000"` |
| 글꼴 크기 | `1200` | `12.0pt` |
| 치수 | `59430` | `210.0mm` |
| 매핑값 | `1` | `1 (center)` |

---

## 6. 등록 시스템

### PARAMETERSET_REGISTRY
```python
# 자동 등록 (ParameterSetMeta):
PARAMETERSET_REGISTRY["CharShape"]  = CharShape   # PascalCase
PARAMETERSET_REGISTRY["charshape"]  = CharShape   # lowercase

# _pset_id가 정의된 경우:
class MySet(ParameterSet):
    _pset_id = "CustomID"
# → PARAMETERSET_REGISTRY["CustomID"] = MySet
```

### wrap_parameterset() 자동 래핑
```
Tier 1: pset.GetIDStr() → REGISTRY 조회 → 전용 클래스
Tier 2: type(pset).__name__ → REGISTRY 조회 → 전용 클래스
Tier 3: GenericParameterSet 폴백
```

---

## 7. 통계

| 항목 | 수량 |
|------|------|
| ParameterSet 서브클래스 | 143개 (GenericParameterSet 포함) |
| Property Descriptor 타입 | 10종 |
| 전체 속성 정의 | ~900개 |
| Backend 구현체 | 4개 |
| 매핑 딕셔너리 | 30+개 |

---

## 8. 의존 관계도

```
parametersets.py
├── functions.py     ← from_hwpunit, to_hwpunit, convert_hwp_color_to_hex, convert_to_hwp_color
├── logging.py       ← get_logger
└── (내부)
    ├── PARAMETERSET_REGISTRY (모듈 레벨)
    ├── ParameterSetMeta → ParameterSet → 143개 서브클래스
    ├── PropertyDescriptor → 10개 하위 타입
    └── 4개 Backend 구현체
```

---

## 9. 주의사항

1. **`self.attributes_names = [...]` 금지**: 읽기 전용 property (`_property_registry.keys()`)
2. **Backend None 체크 필수**: unbound ParameterSet은 `_backend=None`
3. **PascalCase가 COM 키**: snake_case는 Python 편의용, COM은 항상 PascalCase
4. **NestedProperty는 바인딩 필요**: unbound 상태에서 접근 시 `RuntimeError`
5. **GenericParameterSet은 registry 제외**: 폴백 전용
