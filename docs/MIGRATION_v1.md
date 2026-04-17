# v1.0 Migration Guide

> 이전 버전 (v0.0.x) 코드를 v1.0 으로 마이그레이션하는 가이드.

v1.0 은 기존 API 를 **대부분 유지** 하지만 일부 메소드에 `DeprecationWarning`
이 붙었고 (v1.1 에서 제거 예정), 새로운 더 간결한 API 가 권장됩니다.

**단계별 체크리스트**:

1. [ ] 파이썬 3.9+ 사용 (3.7, 3.8 미지원)
2. [ ] 테스트 코드를 실행해서 `DeprecationWarning` 확인
3. [ ] 이 문서의 대응표대로 교체
4. [ ] `python tests/` 전체 통과 확인

---

## 1. 필드 API (v0.0.12+)

v0.0.12 에서 `app.fields` 컬렉션이 도입됨. v1.1 에서 구 API 10개 완전 제거 예정.

| 구 API | 신규 API |
|:---|:---|
| `app.create_field("name")` | `app.fields.add("name")` |
| `app.create_field("n", memo="...", direction="...")` | `app.fields.add("n", memo="...", direction="...")` |
| `app.set_field("name", "value")` | `app.fields["name"] = "value"` |
| `app.get_field("name")` | `app.fields["name"].value` |
| `app.field_exists("name")` | `"name" in app.fields` |
| `app.move_to_field("name")` | `app.fields["name"].goto()` |
| `app.delete_field("name")` | `app.fields.remove("name")` |
| `app.delete_all_fields()` | `app.fields.remove_all()` |
| `app.rename_field("old", "new")` | `app.fields.rename("old", "new")` |
| `app.field_names` | `list(app.fields)` |
| `app.fields_dict` | `app.fields.to_dict()` |

### 대량 값 주입 (일반 패턴)

```python
# ❌ 구 API
for name, value in data.items():
    app.create_field(name)
    app.set_field(name, value)

# ✅ 신 API
app.fields.update(data)    # dict 한 번에

# ✅ 또는
for name, value in data.items():
    app.fields[name] = value   # 필요 시 자동 생성
```

### 브래킷 → 필드 변환

```python
# ❌ 구 API
app.replace_brackets_with_fields()

# ✅ 신 API (의미도 더 명확)
app.fields.from_brackets()
```

---

## 2. 책갈피 / 하이퍼링크 (soft deprecation)

v1.1 에서 제거 예정은 아니지만 accessor 방식 권장:

```python
# 구 API
app.insert_bookmark("chapter1")
app.insert_hyperlink("GitHub", "https://...")

# 신 API (권장)
app.bookmarks.add("chapter1")
app.hyperlinks.add("GitHub", "https://...")

# 추가 기능 — accessor 만의 강점
"chapter1" in app.bookmarks
app.bookmarks.goto("chapter1")
app.bookmarks.remove("chapter1")
```

---

## 3. CharShape / ParaShape (property 권장)

`get_charshape()` / `set_charshape(**kwargs)` 는 여전히 작동하지만
property 형태가 더 Pythonic:

```python
# 구 API
cs = app.get_charshape()
app.set_charshape(bold=True, text_color="red")

# 신 API — property
cs = app.charshape            # 현재 값 읽기
app.charshape = {"bold": True, "text_color": "red"}   # dict 로 설정

# 또는 CharShape 객체 전달
from hwpapi.parametersets import CharShape
new_cs = CharShape(...)
app.charshape = new_cs

# 기존 kwargs 스타일도 유지 (편의성)
app.set_charshape(bold=True, text_color="red")   # 이건 v1.0 에서도 정상
```

---

## 4. 단위 변환 (v0.0.23+)

App 인스턴스 없이도 호출 가능한 `hwpapi.units` 모듈:

```python
# 구 API (instance method)
hwp = app.mm_to_hwpunit(210)
mm = app.hwpunit_to_mm(59430)

# 신 API (module function)
from hwpapi import units as U
hwp = U.mm(210)                  # 59430
mm = U.to_mm(59430)              # 210.0
hwp = U.parse("210mm")           # 문자열도 OK
```

**장점**: App 생성 전이나 테스트 코드에서도 사용 가능.

---

## 5. ENUM / Color (v0.0.25+)

### BorderType

```python
# 구 API (매직 넘버)
app.set_cell_border(top=1, bottom=8)   # 1=solid, 8=double... 의미 불명

# 신 API
app.set_cell_border(top="solid", bottom="double")

# 숫자도 여전히 호환
app.set_cell_border(top=1)   # 정상 작동
```

### HatchStyle

```python
# 구
app.set_cell_color(bg_color="yellow", hatch_style=6)   # 6=?

# 신
app.set_cell_color(bg_color="yellow", hatch_style="diagonal_cross")
```

### Color 상수

```python
# 구
app.set_charshape(text_color="#FF0000")
app.set_charshape(text_color="red")

# 신 (셋 다 정상)
from hwpapi.parametersets.properties import Color
app.charshape = {"text_color": Color.RED}
app.charshape = {"text_color": "red"}      # 명명 색상 16개 확장됨
app.charshape = {"text_color": "#FF0000"}   # hex 도 가능
```

**추가된 색상 이름**: `yellow`, `cyan`, `magenta`, `orange`, `purple`,
`pink`, `brown`, `gray`, `light_gray`, `dark_gray`, `navy` (이전엔 5색만).

---

## 6. `App.charshape()` method 제거 (breaking)

v0.0.7 부터 DeprecationWarning 이었고 property 와 이름 충돌로 호출 불가였지만
v1.0 에서 완전 제거:

```python
# 구 — 호출 시 DeprecationWarning 후 작동 (property 와 충돌로 실제론 호출 안 됨)
cs = app.charshape(bold=True, italic=True)
app.set_charshape(cs)

# 신 (v1.0 이전부터 권장)
app.set_charshape(bold=True, italic=True)        # 직접 적용
# 또는
app.charshape = {"bold": True, "italic": True}   # property setter
```

---

## 7. `replace_font` 버그 수정 (breaking semantic)

v0.0.23 이하의 `replace_font(old, new)` 는 `old` 를 무시하고 문서 전체를
덮어썼음. v0.0.24+ 부터는 정상 동작:

```python
# v0.0.23 이하 (버그 — old 무시)
app.convert.replace_font("맑은 고딕", "함초롬바탕")
# → 모든 폰트를 함초롬바탕으로 덮어씀 ❌

# v0.0.24+ (정상)
app.convert.replace_font("맑은 고딕", "함초롬바탕")
# → "맑은 고딕" 만 함초롬바탕으로 교체 ✓

# legacy 동작이 필요하면:
app.convert.replace_font(None, "함초롬바탕", replace_all=True)
```

**영향**: 이전 버전의 의도가 "전체 덮어쓰기" 였다면 `replace_all=True` 추가.

---

## 8. `app.view.zoom()` 액션 수정 (breaking)

v0.0.23 이하는 `PictureScale` 액션을 잘못 사용 (그림 크기 조절 액션). v0.0.24+:

```python
# 변경 없음 — 사용자 코드는 동일
app.view.zoom(150)

# 내부적으로 ZoomRate property 직접 설정으로 변경됨.
# 이전 버전에서는 작동 안 하던 기능이 이제 정상 동작.
```

---

## 9. find_text / set_charshape facename 다국어

v0.0.24+ 부터 `facename=` 단일 인자가 7개 facename 으로 자동 fan-out:

```python
# 구 — 한글만 매칭 (영문 텍스트 매칭 실패)
app.set_charshape(facename_hangul="맑은 고딕")
app.find_text("HWP", facename="Arial")   # 실제론 FaceNameHangul 만 설정

# 신 — 자동으로 7개 facename 모두에 적용
app.set_charshape(facename="맑은 고딕")
app.find_text("HWP", facename="Arial")    # 7개 facename 모두에 Arial 설정

# 개별 facename 은 여전히 명시 가능
app.set_charshape(facename_hangul="맑은 고딕",
                  facename_latin="Arial")
```

---

## 10. Config 기본값 적용 (breaking semantic)

```python
# v0.0.23 이하 — 저장만 되고 적용 안 됨 (no-op)
app.config.default_font = "함초롬바탕"

# v0.0.24+ — 명시적 적용 호출 필요
app.config.default_font = "함초롬바탕"
app.config.default_size = 1100
app.config.apply_defaults()    # 이제 실제로 charshape 에 반영
```

---

## 자동 마이그레이션 도구 (v1.1 예정)

```bash
python -m hwpapi.migrate --dry-run my_script.py
python -m hwpapi.migrate --write my_script.py
```

현재는 수동 교체. 이 문서의 표를 따라 찾아바꾸기 하면 대부분 변환 가능.

---

## 테스트로 검증

```bash
# 코드 실행 시 DeprecationWarning 출력
python -W always::DeprecationWarning my_script.py

# 예상 출력:
# DeprecationWarning: create_field() is deprecated (v0.0.20+);
# use app.fields.add(name, memo=..., direction=...) instead.
# Will be removed in v1.0.
```

---

## 도움이 필요하면

- GitHub Issues: https://github.com/JunDamin/hwpapi/issues
- 튜토리얼: https://JunDamin.github.io/hwpapi/
- API 레퍼런스: https://JunDamin.github.io/hwpapi/02_api/
