# hwpapi

[![PyPI](https://img.shields.io/pypi/v/hwpapi.svg)](https://pypi.org/project/hwpapi/)
[![Python](https://img.shields.io/pypi/pyversions/hwpapi.svg)](https://pypi.org/project/hwpapi/)
[![License](https://img.shields.io/pypi/l/hwpapi.svg)](https://github.com/JunDamin/hwpapi/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-JunDamin.github.io-2780E3)](https://JunDamin.github.io/hwpapi)
[![Status](https://img.shields.io/badge/status-stable%20(v1.0)-success)](https://github.com/JunDamin/hwpapi)

**한글(HWP) 자동화를 위한 Pythonic 라이브러리.** win32com의 복잡한 COM API를
깔끔한 파이썬 인터페이스로 감싸서, 단순 반복 문서 작업을 간결하게 자동화할 수
있도록 도와줍니다.

> 🎉 **v1.0 Stable Release** — 18개 accessor, 11개 preset, 8개 context manager,
> 1,388개 테스트 통과, 25개 객체별 API 문서 완비. [Migration guide](docs/MIGRATION_v1.md)

> 설계 철학은 [xlwings](https://www.xlwings.org/)에서 영감을 받았습니다.
> 자주 쓰는 기능은 메서드로 정리하고, 부족한 부분은 `app.api`로 원시 COM 접근을
> 그대로 유지하여 HWP API의 모든 기능을 사용할 수 있도록 했습니다.

---

## Requirements

- **Windows** (한글 오피스 설치 필수)
- **Python 3.9+**
- pywin32

리눅스 / macOS는 지원하지 않습니다.

---

## Install

```sh
pip install hwpapi
```

개발 버전:

```sh
git clone https://github.com/JunDamin/hwpapi.git
cd hwpapi
pip install -e .
```

---

## Quick Start

```python
from hwpapi.core import App

# HWP 연결 (실행 중이 아니면 자동 실행)
app = App()

# 어떤 기능이 있는지 궁금하면 한 줄:
app.help()                                      # 18개 accessor + context manager 출력
repr(app)                                       # App(visible=True, version='13.0.0', docs=1, page=1/1)

# 파일 열기
app.open("report.hwp")

# 기본 편집
app.insert_text("안녕하세요, 한글!")
app.charshape = {"bold": True, "text_color": "#FF0000"}   # v0.0.12+ property
app.save()
```

---

## v0.0.14~22 하이라이트 — 18개 Accessor

``app.`` 다음에 무엇이 있는지 IDE tab completion 으로 바로 발견:

```python
# Navigation & Selection
app.move.line.end()                             # 커서 이동 (sub-accessor 7종)
app.sel.current_paragraph().compress(step=2)    # 선택 + 자간/장평 축소

# Collections — dict/list-like
app.fields["name"] = "홍길동"                    # 누름틀 (mail merge)
app.fields.update({"date": "2026-04-15"})
app.bookmarks.add("ch1")
app.hyperlinks.add("GitHub", "https://...")
app.images.resize_all(width="100mm")            # 모든 이미지 일괄 크기

# Structure
app.table.header_row(color="sky", bold=True)
app.table.align(horz="right", scope="current_col")
app.table.clean_excel_paste()                    # 엑셀 붙여넣기 빈 행/열 정돈

# Transform & View
app.convert.number_to_korean()                   # "1,234" → "천이백삼십사"
app.convert.replace_font("맑은 고딕", "함초롬바탕")
app.view.zoom(150)

# Quality & Templates
report = app.lint()                              # 긴 문장·빈 문단 감지
app.template.save("company_style.json")
app.config.default_font = "함초롬바탕"

# Presets (승승아빠 매크로 이식 — 11종)
app.preset.striped_rows()                        # 줄무늬 표
app.preset.title_box(text="보고서 제목", subtitle="부서명")
app.preset.table_header(color="sky", text_color="#FFFFFF")
app.preset.toc(with_bookmarks=True, dot_leader=True)
app.preset.page_numbers(position="bottom-center")

# Debug
app.debug.state()                                # 현재 커서/페이지/charshape 덤프
with app.debug.trace(): ...                      # COM 호출 로그
```

### Context Managers — 대량 처리 필수

```python
with app.batch_mode():                          # 창 숨김 + dialog 억제 → 5~10배 가속
    for row in df.iterrows():
        app.fields.update(row.to_dict())
        app.save(f"out/{row['name']}.hwp")

with app.silenced("yes"): ...                    # 모든 dialog 자동 YES
with app.suppress_errors(): ...                  # 에러 + Python 예외 swallow
with app.undo_group("대량 편집"): ...            # Ctrl+Z 한 번으로 전체 rollback
with app.charshape_scope(bold=True): ...         # 블록 내 임시 서식 (종료 시 복원)
```

### 학습 경로

1. **REPL 에서 바로** → `app.help()` 로 전체 API 탐색
2. **튜토리얼** — https://JunDamin.github.io/hwpapi
   - [10. Accessors Overview](https://JunDamin.github.io/hwpapi/01_tutorials/10_accessors_overview.html) — 18개 accessor 매트릭스
   - [11. Presets Gallery](https://JunDamin.github.io/hwpapi/01_tutorials/11_presets_gallery.html) — 문서 꾸미기 프리셋
   - [12. Batch & Workflow](https://JunDamin.github.io/hwpapi/01_tutorials/12_batch_and_workflow.html) — 실전 대량 처리
   - [13. Debugging Tools](https://JunDamin.github.io/hwpapi/01_tutorials/13_debugging_tools.html) — 품질 / 디버깅 / 설정
3. **레퍼런스** — [docs/API_GUIDE.md](docs/API_GUIDE.md) — 18 accessor 매트릭스 + 레거시 마이그레이션 표

---

## 기존 win32com 코드와 비교

### 텍스트 삽입

**win32com 방식:**
```python
import win32com.client as win32
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True

act = hwp.CreateAction("InsertText")
pset = act.CreateSet()
pset.SetItem("Text", "Hello\r\nWorld!")
act.Execute(pset)
```

**hwpapi 방식:**
```python
from hwpapi.core import App
app = App()
app.insert_text("Hello\r\nWorld!")
```

### 글자 모양

**win32com 방식:**
```python
Act = hwp.CreateAction("CharShape")
Set = Act.CreateSet()
Act.GetDefault(Set)
Set.SetItem("Italic", 1)
Act.Execute(Set)
```

**hwpapi 방식:**
```python
app.set_charshape(italic=True)
```

또는 더 자세한 제어:
```python
action = app.actions.CharShape
ps = action.pset
ps.Italic = True
ps.Bold = True
ps.FontSize = 1400          # 14pt (HWPUNIT=100/pt)
ps.TextColor = "#FF0000"    # hex → BBGGRR 자동 변환
action.run()
```

---

## 주요 기능

### 1. 파일 I/O

```python
app.open("input.hwp")
app.save("output.hwp")
app.save_block("selection.hwp")    # 선택 영역만 저장
app.insert_file("append.hwp")       # 다른 파일 삽입
app.close()
```

### 2. 텍스트 조작

```python
app.insert_text("안녕하세요")
text = app.get_text()
selected = app.get_selected_text()

app.find_text("찾을 문자열")
app.replace_all("이전", "이후")
```

### 3. Pythonic 포매팅

```python
# 간단한 방식
app.set_charshape(bold=True, italic=True, text_color="#FF0000")

# 자세한 방식 (ParameterSet 직접 조작)
ps = app.actions.CharShape.pset
ps.bold = True
ps.text_color = "#FF0000"
ps.height = 1400
app.actions.CharShape.run()
```

### 4. 네비게이션 (Accessor)

```python
app.move.top_of_file()
app.move.bottom()
app.move.next_para()
app.move.end_of_line()
```

### 5. 표/셀 조작

```python
# 표 삽입
action = app.actions.TableCreate
action.pset.rows = 3
action.pset.cols = 4
action.run()

# 셀 스타일
app.set_cell_border(left=True, right=True)
app.set_cell_color("#EEEEEE")
```

### 6. 페이지 설정

```python
app.setup_page(
    width="210mm", height="297mm",
    top_margin="20mm", bottom_margin="20mm",
    left_margin="25mm", right_margin="25mm",
)
```

### 7. Action Introspection (COM 호출 없이)

```python
# 사용 가능한 액션 목록
all_actions = app.actions.list_actions()       # 704개
with_pset = app.actions.list_actions(with_pset_only=True)  # 256개

# 특정 액션의 ParameterSet 클래스
cls = app.actions.get_pset_class('CharShape')  # → CharShape

# 설명
desc = app.actions.get_description('CharShape')  # → "글자 모양"

# 액션 존재 확인
if 'FindReplace' in app.actions:
    app.actions.FindReplace.run()
```

### 8. 원시 COM 접근 (필요 시)

```python
# app.api로 모든 HWP COM API 접근 가능
hwp = app.api
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.Bold = True
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
```

---

## 패키지 구조

```
hwpapi/
├── core/                     # App, Engine 엔트리 포인트
│   ├── app.py                #   → App 클래스 (메인 API)
│   └── engine.py             #   → Engine, Engines, Apps
├── actions.py                # 704개 액션 동적 디스패처 (_Action, _Actions)
├── parametersets/            # 143개 ParameterSet 클래스
│   ├── mappings.py           #   → 35개 string↔int 매핑
│   ├── backends.py           #   → 4개 백엔드 (Com/Attr/Pset/HParam)
│   ├── properties.py         #   → PropertyDescriptor 계층
│   ├── base.py               #   → ParameterSet 기반 클래스
│   └── sets/
│       ├── primitives.py     #     CharShape, ParaShape, BorderFill, Cell, ...
│       ├── drawing.py        #     ShapeObject, Draw* (25개)
│       ├── text.py           #     문자 조작 (20개)
│       ├── paragraph.py      #     TabDef, NumberingShape, ... (4개)
│       ├── table.py          #     Table 관련 (12개)
│       ├── document.py       #     DocumentInfo, PageDef, SecDef, ... (17개)
│       ├── file_ops.py       #     FileOpen, Print, ... (14개)
│       ├── find_edit.py      #     FindReplace, BookMark, ... (13개)
│       ├── formatting.py     #     BorderFillExt, StyleDelete, ... (3개)
│       └── media_misc.py     #     OleCreation, HyperLink, ... (27개)
├── classes/                  # Pythonic 헬퍼 클래스
│   ├── accessors.py          #   → MoveAccessor, CellAccessor, ...
│   └── shapes.py             #   → Character, CharShape (dataclass), ...
├── constants.py              # Enum, 폰트 목록
├── functions.py              # 유닛/색상 변환, COM 헬퍼
└── logging.py                # 로거 설정
```

**자세한 API 가이드는 [`docs/API_GUIDE.md`](docs/API_GUIDE.md)를 참고하세요.**

---

## Documentation

- **[API Guide](docs/API_GUIDE.md)** — 전체 API 레퍼런스, 레시피, AI agents용 네비게이션 맵
- **[ParameterSet Architecture](docs/PARAMETERSET_ARCHITECTURE.md)** — 내부 아키텍처 설명
- **[Refactoring Plan](docs/REFACTORING_PLAN.md)** — 이전 리팩토링 기록
- **[CLAUDE.md](CLAUDE.md)** — Claude Code용 개발 가이드

---

## Testing

```sh
# HWP 없이 돌리는 단위 테스트
python -m pytest tests/ --ignore=tests/test_all_actions.py -q

# HWP 실행 필요한 통합 테스트
python -m pytest tests/test_all_actions.py -q
```

현재 테스트: **2,302 passed** (HWP 없이 돌릴 경우 4개 skip)

---

## 왜 hwpapi를 만들었나요?

직장인으로 많은 한글 문서를 편집하고 작성하면서 단순 반복업무가 너무 많다는 것이
불만이었습니다. `win32com`을 통한 HWP 자동화는 가능했지만 코드가 장황했고, 파이썬
생태계의 자연스러움을 잃어버렸습니다.

엑셀 자동화를 위한 [`xlwings`](https://www.xlwings.org/)가 얼마나 작업 효율을 높이는지
경험하고, 한글에도 이런 라이브러리가 있으면 좋겠다는 생각에 만들게 되었습니다.

기본 철학은 xlwings를 따라합니다:
- 자주 쓰는 항목은 사용하기 쉬운 메서드로 정리
- 부족한 부분은 `App.api`로 원시 COM 접근 유지 → HWP API 전체 기능 사용 가능

관련 참고자료:
- [회사원 코딩 블로그](https://employeecoding.tistory.com/72)
- HWP Automation 공식 문서 (`hwp_docs/` 디렉토리)

---

## Contributing

- Issue/PR 환영합니다
- 리팩토링 가이드: [CLAUDE.md](CLAUDE.md) 참고
- 테스트 작성 후 PR 부탁드립니다

---

## License

MIT

---

## Acknowledgements

- 설계 영감: [xlwings](https://www.xlwings.org/)
- 자동화 기법 참고: [회사원 코딩](https://employeecoding.tistory.com/72)
- HWP Automation 공식 문서: `hwp_docs/`
