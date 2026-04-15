# hwpapi API Guide

**Target audience**: AI agents (Claude Code, Copilot, etc.) and developers using `hwpapi` to automate 한/글 (HWP) document operations.

This guide provides complete API coverage with concrete examples and a navigation map for finding any feature quickly.

> **Tip**: 언제든 ``app.help()`` 호출 → 사용 가능한 accessor 와 context manager 를 6개 카테고리로 출력. Python REPL 에서 discovery 시 최우선.

---

## 0. Accessor Matrix (v0.0.20 기준)

App 에는 **18개 accessor** 가 도메인별로 그룹핑돼 있습니다.

| 카테고리 | Accessor | 주요 메소드 |
|:---|:---|:---|
| **Navigation & Selection** | `app.move` | `.doc.top()`, `.line.end()`, `.word.next()`, `.para.start()`, `.page.next()`, `.cell.right()` |
| | `app.sel` | `.current_word/line/paragraph/sentence()`, `.to_*_end/begin()`, `.expand_char(n)`, `.compress(step)`, `.expand(step)` |
| **Collections** | `app.documents` | 열린 문서 컬렉션 — iter, `.add(is_tab)`, `.open(path)` |
| | `app.fields` | 누름틀 — `["name"] = "v"`, `.add()`, `.remove()`, `.update(dict)`, `.from_brackets()`, `.to_dict()` |
| | `app.bookmarks` | `.add(name)`, `.remove(name)`, `.goto(name)`, `"x" in app.bookmarks` |
| | `app.hyperlinks` | `.add(text, url)` |
| | `app.images` | 이미지 control 순회 — `.resize_all(width="100mm")`, `.grayscale_all()` |
| | `app.styles` | 문단 스타일 — `.apply(name)`, iter |
| | `app.controls` | HeadCtrl→Next linked list 순회 |
| **Structure** | `app.cell` | 표 셀 content/formatting/merge |
| | `app.table` | 표 구조 + 배치 서식 — `.header_row()`, `.footer_row()`, `.current_row()`, `.align(horz, vert, scope)`, `.clean_excel_paste()` |
| | `app.page` | 페이지 설정 |
| **Transform & View** | `app.convert` | `.number_to_korean()`, `.wrap_by_word()`, `.wrap_by_char()`, `.replace_font(old, new)` |
| | `app.view` | `.zoom(%)`, `.zoom_fit_page/width`, `.zoom_actual()`, `.full_screen()`, `.page_mode()`, `.draft_mode()`, `.scroll_to_cursor()` |
| **Quality & Templates** | `app.lint` | **callable** — `report = app.lint()` → LintReport with `.long_sentences`, `.empty_paragraphs`, etc. |
| | `app.template` | `.save(path)`, `.apply(path)` |
| | `app.config` | `.default_font`, `.default_size`, `.palette`, `.save()`, `.load()` |
| **Presets & Debug** | `app.preset` | `.striped_rows()`, `.title_box()`, `.subtitle_bar()`, `.table_header(color="sky")`, `.table_footer()`, `.toc()`, `.page_numbers()`, `.page_border()`, `.highlight_yellow()`, `.summary_box()` |
| | `app.debug` | `.state()`, `.print()`, `.timing(fn)`, `.trace()` (ctx mgr) |

### Context Managers

```python
with app.silenced("yes"): ...           # 다이얼로그 자동 응답
with app.suppress_errors(): ...          # 에러 dialog + Python 예외 swallow
with app.batch_mode(hide=True): ...      # 창 숨김 + dialog 억제 (5~10배 가속)
with app.undo_group("설명"): ...          # 단일 undo 경계
with app.charshape_scope(bold=True): ... # 블록 내 문자 모양 임시
with app.parashape_scope(align=1): ...   # 블록 내 문단 모양 임시
with app.use_document(doc): ...          # 활성 문서 일시 전환
with app.debug.trace(): ...              # COM Run() 호출 로그
```

### Properties (state access)

```python
app.text                # 전체 텍스트 (get/set)
app.visible             # bool (get/set)
app.version             # str (read-only)
app.page_count          # int
app.current_page        # int
app.selection           # str — 선택 텍스트
app.charshape           # CharShape snapshot (get/set)
app.parashape           # ParaShape snapshot (get/set)
app.status              # 종합 상태 dict
```

### Deprecated methods (v1.0 에서 제거)

v0.0.20 부터 DeprecationWarning 발생. 동작은 유지.

| 레거시 | 신규 권장 |
|:---|:---|
| `app.create_field(name)` | `app.fields.add(name)` |
| `app.set_field(n, v)` | `app.fields[n] = v` |
| `app.get_field(n)` | `app.fields[n].value` |
| `app.field_exists(n)` | `n in app.fields` |
| `app.move_to_field(n)` | `app.fields[n].goto()` |
| `app.delete_field(n)` | `app.fields.remove(n)` |
| `app.delete_all_fields()` | `app.fields.remove_all()` |
| `app.rename_field(o, n)` | `app.fields.rename(o, n)` |
| `app.field_names` | `list(app.fields)` |
| `app.fields_dict` | `app.fields.to_dict()` |

---

## 1. Quick Start

```python
from hwpapi.core import App

app = App()                                    # Connect to HWP (launches if needed)
app.open("document.hwp")                       # Open a file
app.insert_text("Hello, 한글!")                # Insert text
app.save()                                     # Save

# Manipulate formatting via actions
action = app.actions.CharShape                 # Get typed _Action
pset = action.pset                             # Get typed ParameterSet
pset.bold = True                               # snake_case OR PascalCase both work
pset.FontSize = 1200                           # 12pt (HWPUNIT: 100/pt)
pset.TextColor = "#FF0000"                     # hex color auto-converted
action.run()                                   # Apply
```

---

## 2. Package Structure

```
hwpapi/
├── core.py              # App, Engine, Apps classes — main entry
├── actions.py           # _Action, _Actions — HWP action dispatcher
├── parametersets/       # ParameterSet system (143 typed classes)
│   ├── __init__.py      # Public API: base, helpers, imports from submodules
│   ├── mappings.py      # String↔int mappings (DIRECTION_MAP, etc.)
│   ├── backends.py      # 4 backend implementations (COM ↔ Python)
│   └── sets/            # ParameterSet subclasses by domain
│       ├── primitives.py   # 8 base classes (CharShape, ParaShape, BorderFill, Cell, Caption, CtrlData, Password, Style)
│       ├── drawing.py      # 25 classes: ShapeObject, Draw*
│       ├── text.py         # 20 classes: character ops
│       ├── paragraph.py    # 4 classes: TabDef, NumberingShape, etc.
│       ├── table.py        # 12 classes: Table, CellBorderFill, etc.
│       ├── document.py     # 17 classes: DocumentInfo, PageDef, SecDef, etc.
│       ├── formatting.py   # 3 classes: BorderFillExt, StyleDelete, StyleTemplate
│       ├── file_ops.py     # 14 classes: FileOpen, Print, etc.
│       ├── find_edit.py    # 13 classes: FindReplace, BookMark, etc.
│       └── media_misc.py   # 27 classes: OleCreation, HyperLink, etc.
├── classes.py           # Accessor classes (MoveAccessor, CellAccessor, etc.)
├── constants.py         # Enums, font lists
├── functions.py         # Unit/color conversion, COM helpers
└── logging.py           # Logger configuration
```

---

## 3. Finding What You Need

### 3.1 Need to find a specific ParameterSet class?

| Looking for | Check |
|-------------|-------|
| `CharShape`, `ParaShape`, `BorderFill`, `Cell`, `Caption`, `Password`, `Style`, `CtrlData` | `sets/primitives.py` |
| `ShapeObject`, `DrawFillAttr`, `DrawLineAttr`, `DrawShadow`, etc. | `sets/drawing.py` |
| `BulletShape`, `Convert*`, `InsertText`, `InputHanja*`, character ops | `sets/text.py` |
| `TabDef`, `ListProperties`, `NumberingShape`, `ListParaPos` | `sets/paragraph.py` |
| `Table`, `TableCreation`, `CellBorderFill`, `AutoFill` | `sets/table.py` |
| `DocumentInfo`, `PageDef`, `SecDef`, `HeaderFooter`, `FootnoteShape` | `sets/document.py` |
| `BorderFillExt`, `StyleDelete`, `StyleTemplate` | `sets/formatting.py` |
| `FileOpen`, `FileSaveAs`, `Print`, `Preference` | `sets/file_ops.py` |
| `FindReplace`, `BookMark`, `IndexMark`, `MakeContents` | `sets/find_edit.py` |
| `OleCreation`, `HyperLink`, `FieldCtrl`, `Ftp*`, misc | `sets/media_misc.py` |

### 3.2 Always import from the package root

```python
# ✅ CORRECT — all classes re-exported at package root
from hwpapi.parametersets import CharShape, FindReplace, PARAMETERSET_REGISTRY

# ⚠️ Avoid — internal module paths may change
from hwpapi.parametersets.sets.primitives import CharShape  # works but fragile
```

---

## 4. Core Concepts

### 4.1 App — The Main Entry Point

```python
from hwpapi.core import App

app = App(is_visible=True)     # Visible window
app = App(new_app=True)        # Force new HWP instance
app.api                        # Raw COM object (hwp.HwpObject)
```

Key methods (all on `App`):
- **Files**: `open(path)`, `save()`, `save_block(path)`, `close()`, `quit()`
- **Text**: `insert_text(text)`, `get_text()`, `get_selected_text()`, `find_text()`, `replace_all()`
- **Formatting**: `get_charshape()`, `set_charshape()`, `get_parashape()`, `set_parashape()`
- **Navigation**: `app.move.top_of_file()`, `app.move.bottom()`, etc. (see Accessor section)
- **Page**: `setup_page()`, `get_filepath()`, `get_hwnd()`
- **Window**: `set_visible(bool)`, `reload()`

### 4.2 Actions — Typed Action Dispatcher

Actions are dynamically dispatched via `__getattr__`, cached per-instance:

```python
# Access any of 704 actions
action = app.actions.CharShape              # Returns cached _Action
action = app.actions.FindReplace
action = app.actions.InsertText

# Introspection WITHOUT triggering COM calls
cls = app.actions.get_pset_class('CharShape')  # → CharShape class
desc = app.actions.get_description('CharShape')  # → "글자 모양"
all_names = app.actions.list_actions()
with_pset = app.actions.list_actions(with_pset_only=True)  # 256 actions
'CharShape' in app.actions                  # True
app.actions.refresh()                        # Clear cache (reload defaults)
```

Each `_Action` has:
- `action_key`: Action name string
- `pset_key`: SetID of associated ParameterSet (or None)
- `description`: Korean description from `_action_info`
- `pset`: Bound typed `ParameterSet` instance (lazy, cached)
- `act`: Raw COM action object

Action execution:
```python
action = app.actions.CharShape
action.pset.bold = True           # Modify parameters
action.pset.FontSize = 1200
action.run()                       # Execute with current pset
# OR
action.set_parameter('Bold', 1).run()    # Chain-style
```

### 4.3 ParameterSet — Typed Parameter Wrappers

Each ParameterSet subclass wraps a COM pset object with:
- **143 classes** covering all HWP actions that have parameters
- **snake_case + PascalCase access** (automatic conversion)
- **Typed property descriptors** with validation
- **Auto-wrapping of nested ParameterSets**
- **Native COM methods**: `clone()`, `is_equivalent()`, `merge()`, `item_exists()`

```python
ps = app.actions.CharShape.pset
ps.bold = True          # snake_case
ps.Bold = True          # PascalCase — same attribute
ps.face_name_hangul = "맑은 고딕"
ps.FaceNameHangul                # → "맑은 고딕"
ps.TextColor = "#FF0000"         # hex → 0x0000FF (BBGGRR)
ps.Height = 1200                 # HWPUNIT (100 = 1pt)

# Native COM methods
cloned = ps.clone()
ps.is_equivalent(other_ps)
ps.merge(other_ps)
ps.item_exists("Bold")

# Listing
ps.attributes_names              # List all property names
```

### 4.4 PropertyDescriptor Types (10 types)

Defined in `parametersets/__init__.py`:

| Descriptor | Use for | Example |
|------------|---------|---------|
| `PropertyDescriptor(key, doc)` | Generic | `FaceNameHangul` |
| `IntProperty(key, doc)` | Integers | `Size = IntProperty(...)` |
| `BoolProperty(key, doc)` | Booleans → 0/1 | `Bold = BoolProperty(...)` |
| `StringProperty(key, doc)` | Strings | `Text = StringProperty(...)` |
| `ColorProperty(key, doc)` | Colors (hex ↔ BBGGRR) | `TextColor = ColorProperty(...)` |
| `UnitProperty(key, doc, default_unit)` | mm/cm/in/pt ↔ HWPUNIT | `Width = UnitProperty("Width", "", "mm")` |
| `MappedProperty(key, mapping, doc)` | Enum strings ↔ ints | `Align = MappedProperty("Align", ALIGN_MAP, "")` |
| `TypedProperty(key, doc, cls)` | Nested ParameterSet | `SubPset = TypedProperty(..., SubClass)` |
| `NestedProperty(key, setid, cls, doc)` | Auto-creating nested | Used in `FindReplace` |
| `ArrayProperty(key, item_type, doc)` | HArray (Python list) | `Items = ArrayProperty(...)` |
| `ListProperty(key, doc)` | Python list | Generic list |

### 4.5 Mappings (string ↔ int)

```python
from hwpapi.parametersets import DIRECTION_MAP, ALIGN_MAP, ALL_MAPPINGS

DIRECTION_MAP           # {"left": 0, "right": 1, "top": 2, "bottom": 3}
ALIGN_MAP               # {"left": 0, "center": 1, "right": 2}
ALL_MAPPINGS["align"]   # Same as ALIGN_MAP
```

35 mappings available: `DIRECTION_MAP`, `ALIGNMENT_MAP`, `VERT_ALIGN_MAP`, `HORZ_ALIGN_MAP`, `FONTTYPE_MAP`, `TEXT_DIRECTION_MAP`, `SHADOW_TYPE_MAP`, `BACKGROUND_TYPE_MAP`, `SEARCH_DIRECTION_MAP`, `UNDERLINE_TYPE_MAP`, `OUTLINE_TYPE_MAP`, `STRIKEOUT_TYPE_MAP`, `LINE_WRAP_MAP`, `TEXT_WRAP_MAP`, `TEXT_FLOW_MAP`, ... (full list in `mappings.py`).

### 4.6 Registry Lookup

```python
from hwpapi.parametersets import PARAMETERSET_REGISTRY

PARAMETERSET_REGISTRY["CharShape"]       # → CharShape class
PARAMETERSET_REGISTRY["charshape"]       # same (lowercase alias)
```

Contains 286 entries: each class registered by name, lowercase, and optionally by `_pset_id`.

---

## 5. Common Recipes

### 5.1 Apply Character Formatting to Selection

```python
from hwpapi.core import App

app = App()
app.select_text(0, 0, 0, 10)   # Select some text first (or use Find)

action = app.actions.CharShape
ps = action.pset
ps.bold = True
ps.italic = True
ps.TextColor = "#0000FF"
ps.Height = 1400               # 14pt
action.run()
```

### 5.2 Find and Replace

```python
action = app.actions.AllReplace
ps = action.pset                   # FindReplace ParameterSet
ps.FindString = "old text"
ps.ReplaceString = "new text"
ps.IgnoreCase = False
ps.WholeWordOnly = False

# Optional: nested formatting
ps.find_char_shape.bold = True     # Only find BOLD occurrences

action.run()
```

### 5.3 Set Page Margins

```python
action = app.actions.PageSetup
ps = action.pset                   # PageDef ParameterSet
ps.width = "210mm"                 # A4 width (auto-converts to HWPUNIT)
ps.height = "297mm"
ps.top_margin = "20mm"
ps.bottom_margin = "20mm"
ps.left_margin = "25mm"
ps.right_margin = "25mm"
action.run()
```

### 5.4 Insert a Table

```python
action = app.actions.TableCreate
ps = action.pset                   # TableCreation
ps.rows = 3
ps.cols = 4
ps.width_type = 0                  # Fixed width
ps.width_value = "150mm"
action.run()
```

### 5.5 Iterate All Actions with ParameterSets

```python
for action_name in app.actions.list_actions(with_pset_only=True):
    cls = app.actions.get_pset_class(action_name)
    desc = app.actions.get_description(action_name)
    print(f"{action_name} ({cls.__name__}): {desc}")
```

### 5.6 Inspect a ParameterSet's Properties

```python
ps = app.actions.CharShape.pset
print(f"Class: {type(ps).__name__}")
print(f"Properties: {len(ps.attributes_names)}")
for name in ps.attributes_names:
    descriptor = type(ps)._property_registry[name]
    print(f"  {name}: {type(descriptor).__name__}")
```

### 5.7 Raw COM Access (Advanced)

```python
# Get raw HWP COM object
hwp = app.api

# HAction pattern (modern, uses HParameterSet)
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.Bold = True
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

# IDHwpAction pattern (legacy, uses local pset from CreateSet)
act = hwp.CreateAction("CharShape")
pset = act.CreateSet()
act.GetDefault(pset)
pset.SetItem("Bold", 1)
act.Execute(pset)
```

---

## 6. Troubleshooting

### "No action 'XXX'"
Action name is case-sensitive. Check `'XXX' in app.actions` first.

### "_action_info SetID mismatch for 'YYY'"
Warning (not error): HWP's runtime SetID differs from static metadata.
hwpapi auto-uses HWP's value. No action needed.

### "RuntimeError: Cannot access nested property 'XXX': ParameterSet not bound"
You accessed a `NestedProperty` on an unbound ParameterSet. Either:
1. Get the pset via an action: `app.actions.FindReplace.pset.find_char_shape...`
2. Bind first: `ps = FindReplace(raw_pset_com)`

### Accidentally modified `attributes_names`
It's a read-only property. Use `_property_registry.keys()` instead.

---

## 7. Extension Points

### Add a new ParameterSet class

1. Identify the SetID (HWP docs or runtime introspection)
2. Pick domain: `sets/<domain>.py`
3. Define class:
   ```python
   class MyPset(ParameterSet):
       """Description."""
       MyField = IntProperty("MyField", "My field description")
       OtherField = BoolProperty("OtherField", "...")
   ```
4. Class auto-registers via metaclass; usable via `PARAMETERSET_REGISTRY["MyPset"]`

### Reference another ParameterSet

- **Python-level (import required)**: `from .primitives import CharShape`
- **Logical (docstring only)**: Just mention in description string

---

## 8. Testing

```bash
# Unit tests (no HWP needed)
python -m pytest tests/ -q --ignore=tests/test_all_actions.py

# Integration (HWP required)
python -m pytest tests/test_all_actions.py -q
```

Test files:
- `test_all_parametersets.py`: 1037 tests across 143 classes
- `test_architecture.py`: 1025 structural invariants
- `test_all_actions.py`: 3500+ tests (HWP integration)
- `test_functions.py`: 50 tests for utility functions
- `test_constants.py`: 67 tests for enums/fonts
- `test_classes.py`: 18 tests for accessors/dataclasses
- `test_hparam.py`: 13 tests for HParam backend

---

## 9. Official Reference

- **HwpAutomation_2504.pdf** (hwp_docs/): HAction, HParameterSet, IDHwpAction APIs
- **ActionTable_2504.pdf**: All 704+ actions with SetIDs
- **ParameterSetTable_2504.pdf**: All ParameterSet class definitions with properties

---

## 10. File Navigation (for AI agents)

When you need to modify:
- **Value mappings**: `hwpapi/parametersets/mappings.py`
- **Backend logic**: `hwpapi/parametersets/backends.py`
- **Base class / metaclass / property descriptors / helpers**: `hwpapi/parametersets/__init__.py`
- **Specific ParameterSet class**: `hwpapi/parametersets/sets/<domain>.py` (see table in §3)
- **Action dispatcher**: `hwpapi/actions.py`
- **Main App API**: `hwpapi/core.py`
- **Enum/font constants**: `hwpapi/constants.py`
- **Utility functions**: `hwpapi/functions.py`

Read strategy for large files:
1. Start with `__init__.py` re-exports to find symbol
2. Use `Grep` for the class/function name to locate the file
3. Read the specific file (all under 500 lines for sets/)
