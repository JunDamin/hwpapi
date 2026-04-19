# pyhwpx vs hwpapi — 기능 비교 및 도입 권장 리스트

> 작성일: 2026-04-15  
> 비교 대상: `pyhwpx 1.7.2` vs `hwpapi 0.0.4`

## 요약 통계

| 지표 | pyhwpx | hwpapi |
|:---|---:|---:|
| Public 메소드 수 (단일 기준) | 216 | 152 |
| pyhwpx 에만 있는 메소드 | **193** | — |
| hwpapi 에만 있는 메소드 | — | 129 |
| 공통 (동등) | — | 약 30 |

hwpapi 와 pyhwpx 는 **설계 철학이 다릅니다** — hwpapi 는 xlwings 스타일의 계층적 · 객체지향 설계 (App/Document/접근자/Action/ParameterSet), pyhwpx 는 flat한 300줄짜리 Hwp 클래스에 모든 것을 담는 방식. 그래서 pyhwpx 에 있는 **단일 이름 메소드** 들이 hwpapi 에서는 **접근자에 분산** 되어 있거나 **action/pset 경로** 로 제공됩니다.

아래는 pyhwpx 에 있고 hwpapi 에 **실제로 부재하는** 기능만 추려 **가치 vs 난이도** 로 분류한 것입니다.

---

## 🔴 High Priority — 가치 높음, 꼭 추가 권장 (14개)

| 메소드 | 기능 | 난이도 | hwpapi 추가 제안 |
|:---|:---|:---:|:---|
| **`create_field(name, ...)`** | 필드(변수) 생성 — mail merge 핵심 | S | `app.create_field(name, direction="", memo="")` |
| **`put_field_text(name, text)`** | 필드에 값 채우기 | S | `app.set_field(name, value)` |
| **`get_field_text(name)`** | 필드 값 읽기 | S | `app.get_field(name)` |
| **`get_field_list()`** | 전체 필드 목록 | S | `app.fields` property (list) |
| **`fields_to_dict()`** | `{이름: 값}` 딕셔너리로 반환 | S | `app.fields_dict` property |
| **`set_field_by_bracket()`** | `{{name}}` 같은 브래킷 → 필드 자동 변환 | M | 동일 이름 |
| **`field_exist(name)`** | 존재 확인 | S | `"name" in app.fields` |
| **`delete_field_by_name(name)`** | 필드 삭제 | S | 동일 이름 |
| **`table_from_data(df)`** | pandas DataFrame → HWP 표 | S | `app.insert_table(data=df)` 확장 |
| **`table_to_df()`** | 현재 표 → DataFrame | M | `app.read_table()` 또는 `table.to_df()` |
| **`table_to_csv()`** | 현재 표 → CSV 문자열 | S | `app.table.to_csv()` |
| **`goto_page(n)`** | N 번째 페이지로 이동 | S | `app.goto_page(n)` |
| **`rgb_color(r, g, b)`** | RGB 튜플 → HWP 색상 정수 | S | `Color.from_rgb(r, g, b)` (class method) |
| **`markpen_on_selection(color)`** | 선택에 형광펜 적용 (shade 와 다름) | S | `app.highlight(color)` |

### mail merge 중요성

HWP 의 필드 (field) 시스템은 **mail merge · 템플릿 · 양식** 의 핵심입니다. pyhwpx 가 이 영역에 16개의 메소드를 투입한 이유.

```python
# 기대되는 사용 흐름 (pyhwpx 스타일)
tmpl = app.open("contract_template.hwp")  # {{name}}, {{date}} 필드 있음
tmpl.set_field_by_bracket()                 # 브래킷을 실제 필드로 변환
for row in customers:
    app.set_field("name", row["name"])
    app.set_field("date", row["date"])
    app.save_as(f"out/contract_{row['id']}.hwp")
```

hwpapi 에 이 기능이 없으면 mail merge 가 어려움.

---

## 🟡 Medium Priority — 편의성, 추가 가치 있음 (12개)

| 메소드 | 기능 | 난이도 | 제안 |
|:---|:---|:---:|:---|
| **`hwp_unit_to_inch / point / mm`** | 단위 변환 함수들 | S | `hwpapi.functions` 에 이미 있음 — App 메소드로도 노출 |
| **`inch_to_hwp_unit`, `point_to_hwp_unit`** | 역변환 | S | 위와 동일 |
| **`get_image_info(path)`** | 이미지 메타 (크기/형식) | S | 유틸리티 |
| **`resize_image(ctrl, w, h)`** | 삽입된 이미지 크기 조정 | M | 편의 |
| **`save_all_pictures(dir)`** | 문서 내 모든 이미지 추출 | M | 독특한 기능 |
| **`save_pdf_as_image()`** | 페이지를 PDF 거쳐 PNG로 | M | 우리도 `generate_doc_artifacts` 에서 이미 쓰고 있음 — 메소드화 |
| **`create_page_image(n, path)`** | 특정 페이지를 이미지로 내보내기 | M | 위와 유사 |
| **`key_indicator()`** | 상태바 정보 (페이지/줄/칸/글자 수) | S | `app.status` (dict property) |
| **`is_empty_page() / is_empty_para()`** | 페이지/문단 비어있음 여부 | S | `page.is_empty`, `app.current_para_empty` |
| **`is_modified()`** | 수정 여부 | S | hwpapi 는 `doc.modified` 있음 — 동등 |
| **`insert_lorem(n)`** | Lorem ipsum 삽입 (demo/테스트) | S | 편의 |
| **`compose_chars(jamo_list)`** | 자모 → 글자 조합 | M | 한글 특화 |

---

## 🟢 Low Priority — 특수 용도, 선택적 (많음)

### Style 관리 (10개)

```python
app.get_style(name)
app.set_style(name)
app.modify_style(name, **kwargs)
app.delete_style_by_name(name)
app.remove_unused_styles()
app.import_style(path)
app.export_style(path)
app.get_style_dict()
app.get_used_style_dict()
app.goto_style(name)
```
→ **제안**: `app.styles` 접근자로 묶기 (`app.styles.get(name)`, `app.styles.import_from(...)`)

### Metatag (7개)

```python
app.get_metatag_list()
app.move_to_metatag(name)
app.put_metatag_name_text(name, text)
app.rename_metatag(old, new)
...
```
→ 사용 빈도 낮음. 필요한 사용자가 `app.api` 로 접근 가능.

### Control 조작 (8개)

```python
app.ctrl_list()               # 문서 내 모든 컨트롤
app.find_ctrl(type)           # 컨트롤 찾기
app.delete_ctrl(ctrl)         # 삭제
app.select_ctrl(ctrl)         # 선택
app.move_to_ctrl(ctrl)        # 이동
app.get_ctrl_by_ctrl_id(id)   # ID 로 조회
```
→ **제안**: `app.controls` 접근자 (리스트) + `for ctrl in app.controls` 순회.

### 보안 / 개인정보 (4개)

```python
app.protect_private_info()
app.register_private_info_pattern(pattern)
app.find_private_info(pattern)
app.set_private_info_password(pw)
```
→ 특수 용도, 필요할 때 구현.

### 기타

- `msgbox(text)` — HWP 메시지박스 띄우기 (인터랙티브)
- `run_script_macro(name)` — 매크로 실행
- `export_mathml()` / `import_mathml()` — 수식 MathML 변환
- `auto_spacing()` — 자동 띄어쓰기
- `file_translate()` — 번역 연동
- `lock_command() / release_action()` — 명령 잠금
- `register_module()` — 보안 모듈 등록 (hwpapi 는 `check_dll` 로 처리)

---

## 🔄 이미 hwpapi 에 있지만 이름 다른 것 (10개)

| pyhwpx | hwpapi 동등품 |
|:---|:---|
| `add_doc` | `app.documents.add()` / `app.new_document()` |
| `add_tab` | `app.documents.add(is_tab=True)` |
| `doc_list` | `list(app.documents)` |
| `switch_to(idx)` | `app.documents[idx].activate()` or `with app.use_document(...)` |
| `open_pdf` | `app.open()` (자동 감지) |
| `create_table(r, c)` | `app.insert_table(rows=r, cols=c)` |
| `insert_hyperlink` | `app.insert_hyperlink(text, url)` |
| `find_forward / find_backward` | `app.find_text(direction=0/1)` |
| `find_replace / find_replace_all` | `app.replace_all(old, new)` |
| `current_printpage` | `app.current_page` (약간 다르지만 유사) |
| `save_block_as` | `app.save_block(path)` |
| `maximize_window / minimize_window` | `app.visible = True/False` (부분) |

---

## 📋 권장 실행 계획

### Phase A — Mail Merge 기초 (v0.0.5 목표)

가장 큰 가치, 작은 작업량:

- [ ] `app.fields` — 필드 목록 property
- [ ] `app.create_field(name, memo="")` 
- [ ] `app.set_field(name, value)` (pyhwpx 의 `put_field_text`)
- [ ] `app.get_field(name)` (`get_field_text`)
- [ ] `app.delete_field(name)`
- [ ] `app.replace_brackets_with_fields()` (`set_field_by_bracket`)

예상 규모: ~300줄, 테스트 포함. 위 기능으로 mail merge 사용 사례 (`09_usecase_mail_merge.ipynb`) 추가.

### Phase B — Pandas 연동 (v0.0.5)

- [ ] `app.insert_table(data=df, headers=...)` — DataFrame 직접 지원 확장
- [ ] `app.read_table()` — 현재 표 → DataFrame (pyhwpx 의 `table_to_df`)
- [ ] `app.table.to_csv()` / `table.to_string()`

예상 규모: ~200줄 (pandas optional dep).

### Phase C — 편의 헬퍼 (v0.0.6)

- [ ] `app.goto_page(n)` 
- [ ] `app.highlight(color="#FFFF00")` (markpen_on_selection)
- [ ] `Color.from_rgb(r, g, b)` class method
- [ ] `app.status` property (key_indicator 반환값)
- [ ] `app.save_page_image(n, path)` (create_page_image)
- [ ] `app.extract_all_images(dir)` (save_all_pictures)

예상 규모: ~150줄.

### Phase D — Style 접근자 (v0.1.x)

- [ ] `app.styles` 접근자 클래스
- [ ] `styles.get(name) / set(...) / modify(...) / delete(...)`
- [ ] `styles.import_from(path) / export_to(path)`

예상 규모: ~250줄.

### Phase E — Controls 접근자 (v0.1.x)

- [ ] `app.controls` — 문서 내 컨트롤 iterator
- [ ] `ctrl.select() / delete() / move_to()` 인스턴스 메소드

예상 규모: ~200줄.

---

## 💡 관찰 사항

1. **pyhwpx 의 flat 구조 vs hwpapi 의 계층 구조**: pyhwpx 는 sklearn 스타일 (모든 것이 `Hwp.*`) 이고 hwpapi 는 xlwings 스타일 (접근자 계층). hwpapi 의 방식이 탐색성 (autocomplete) 과 문서화 측면에서 우위.

2. **pandas 연동의 강점**: pyhwpx 는 `table_to_df`, `table_from_data` 등으로 데이터 분석 워크플로우와 자연스럽게 연결됨. 이는 hwpapi 가 반드시 따라잡아야 할 영역.

3. **필드 (field) 는 실무 필수**: 계약서, 양식 문서, 메일 머지 등 한국 실무 HWP 사용의 큰 비중. hwpapi 에서 이 부분이 가장 큰 공백.

4. **hwpapi 에 있지만 pyhwpx 에 없는 것**: 
   - `styled_text`, `charshape_scope`, `parashape_scope` — snapshot/restore 기반 서식 스코핑
   - `use_document` context manager + xlwings 스타일 Document proxy
   - `Documents` 컬렉션 (리스트 인터페이스)
   - ParameterSet 시스템 (286개 클래스, 자동 추출 API 레퍼런스)
   - `Color` / `UNSET` 타입 시스템

이 부분들은 **hwpapi 의 고유한 강점** 으로 유지.

---

## 결론

**바로 시작할 3가지:**

1. 🔴 **필드 API** (Phase A) — mail merge 를 위한 6개 메소드 + 사용 사례 튜토리얼
2. 🔴 **pandas 연동** (Phase B) — `table_to_df`, `insert_table(data=df)` 확장
3. 🟡 **편의 헬퍼** (Phase C) — `goto_page`, `highlight`, `status` 등

이 셋은 모두 backward-compatible 추가이며, 각각 ~200-300줄 규모로 한 번의 세션에서 구현 가능합니다.

v0.0.5 릴리즈 타겟.

---

## 참고

- pyhwpx 공식: https://pyhwpx.github.io/
- GitHub: https://github.com/martiniifun/pyhwpx
- PyPI: https://pypi.org/project/pyhwpx/

Sources:
- [pyhwpx PyPI](https://pypi.org/project/pyhwpx/)
- [pyhwpx 공식 문서](https://pyhwpx.github.io/)
- pyhwpx 1.7.2 source code (core.py 9,730줄, 216 methods)
