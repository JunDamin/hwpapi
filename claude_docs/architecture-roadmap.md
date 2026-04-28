# HWP Object Model: Official vs Current Architecture

## Overview

This document compares the **official HWP Automation Object Model** (from HwpAutomation_2504.pdf) with the **current hwpapi implementation** to identify gaps, misalignments, and opportunities for better code organization.

**Documentation Source:** `hwp_docs/HwpAutomation_2504.pdf` (Korean, dated 2025-04-15)

## Official HWP Object Model Structure

The official HWP automation follows a **hierarchical object model** similar to Microsoft Office:

```
IHwpObject (Root COM Object)
│
├── IXHwpDocuments (Collection)
│   └── IXHwpDocument (Single)
│       ├── Properties: FullName, Name, Path, Saved, etc.
│       └── Methods: Save(), SaveAs(), Close(), Print(), etc.
│
├── IXHwpWindows (Collection)
│   └── IXHwpWindow (Single)
│       ├── Properties: Width, Height, Left, Top, Active, etc.
│       └── Methods: Activate(), Close(), etc.
│
├── IXHwpForms (Collection)
│   └── Form Controls (Various types)
│       ├── IXHwpFormPushButtons (Collection)
│       ├── IXHwpFormCheckButtons (Collection)
│       ├── IXHwpFormRadioButtons (Collection)
│       ├── IXHwpFormComboBoxes (Collection)
│       └── etc.
│
├── HAction (Action Execution System)
│   ├── GetActionIDByName(name) → ActionID
│   ├── Run(ActionID)
│   └── Execute(ActionID, ParameterSet)
│
├── HParameterSet (Parameter Management)
│   ├── CreateItemSet(SetID, ParamIndex) → Creates nested parameter set
│   ├── Item(ParamIndex) → Get parameter value
│   ├── SetItem(ParamIndex, Value) → Set parameter value
│   └── Clear() → Clear all parameters
│
├── HSet (Parameter Collection - Legacy)
│   └── Collection of parameters for complex actions
│
└── HArray (Parameter Arrays - PIT_ARRAY type)
    ├── Count → Number of elements
    ├── Item(index) → Get element at index
    ├── SetItem(index, value) → Set element at index
    ├── Add(value) → Append element
    └── RemoveAt(index) → Remove element at index
```

## Current hwpapi Architecture

```
App (Main Entry Point)
│
├── Engine
│   └── impl (HwpObject COM object)
│
├── _Actions (900+ actions as properties)
│
├── ParameterSet System (130+ classes in parametersets.py)
│   ├── Base: ParameterSet, ParameterSetMeta
│   ├── Backend Abstraction (PsetBackend, HParamBackend, ComBackend, AttrBackend)
│   ├── Property Descriptors
│   └── 130+ ParameterSet Subclasses
│
├── Custom Accessors (MoveAccessor, CellAccessor, TableAccessor, PageAccessor)
│
└── Dataclasses (Character, Paragraph, PageShape)
```

## Comparison Matrix

| Aspect | Official HWP Model | Current hwpapi | Alignment |
|--------|-------------------|----------------|-----------|
| **Entry Point** | `IHwpObject` COM object | `App` wrapper around `Engine` | ✅ Aligned (wrapped) |
| **Document Access** | `IXHwpDocuments` collection | `App.api` direct access | ❌ Collection pattern not exposed |
| **Window Management** | `IXHwpWindows` collection | `App.set_visible()` only | ⚠️ Partial (no multi-window support) |
| **Form Controls** | `IXHwpForms` collection | Not exposed | ❌ Missing |
| **Action Execution** | `HAction.Execute(id, pset)` | `app.actions.ActionName(pset)` | ✅ Aligned (pythonic wrapper) |
| **Parameter Sets** | `HParameterSet` COM object | `ParameterSet` Python classes | ✅ Well abstracted |
| **Parameter Typing** | COM types | Python property descriptors | ✅ Excellent (better than COM) |
| **Nested Parameters** | `CreateItemSet` method | `NestedProperty` auto-creates | ✅ Enhanced (auto-creating) |
| **Arrays (HArray)** | COM array methods | `ArrayProperty` + `HArrayWrapper` | ✅ Enhanced (Pythonic list) |
| **Navigation** | Object hierarchy | Custom accessors | ⚠️ Different paradigm |
| **Organization** | Domain-based modules | Single monolithic file | ❌ Poor organization |

## Identified Gaps

### 1. Missing Collection Objects ❌
- Cannot enumerate open documents
- Cannot manage multiple windows
- No access to form controls
- Limits multi-document workflows

### 2. Monolithic ParameterSets Module ❌
All 130+ ParameterSet classes crammed into single 3,357-line file. Hard to navigate, no logical grouping, merge conflicts, slow IDE performance.

### 3. Navigation Paradigm Mismatch ⚠️
Official model uses object hierarchy (`document.Sections[0].Paragraphs[5]`), hwpapi uses position-based accessors. **Analysis:** Current approach is more pragmatic for HWP's position-based model. No change needed.

### 4. Form Controls Not Exposed ❌
No access to IXHwpForms, form button controls, etc.

## Proposed Restructuring Plan

### Phase 1: Reorganize ParameterSets Module (High Priority)

Split monolithic `parametersets.py` into domain-based submodules:

```
hwpapi/parametersets/
├── __init__.py              # Re-export all classes for compatibility
├── base.py                  # ParameterSet base class, metaclass
├── backends.py              # Backend protocol, implementations
├── properties.py            # Property descriptors
├── mappings.py              # All DIRECTION_MAP, ALIGNMENT_MAP, etc.
├── text/
│   ├── character.py         # CharShape, BulletShape
│   ├── paragraph.py         # ParaShape, TabDef, ListProperties
│   └── numbering.py         # NumberingShape, AutoNum
├── table/
│   ├── table.py             # Table, TableCreation
│   └── cell.py              # Cell, CellBorderFill
├── drawing/
│   ├── shape.py
│   ├── line.py
│   ├── image.py
│   └── effects.py
├── document/
│   ├── info.py
│   ├── page.py
│   └── section.py
├── formatting/
├── actions/
└── forms/
```

**Benefits:** maintainability, IDE performance, fewer merge conflicts, backward compatible via re-exports.

### Phase 2: Expose Collection Objects (Medium Priority)

```python
class App:
    @property
    def documents(self):
        return DocumentsCollection(self.api)

    @property
    def windows(self):
        return WindowsCollection(self.api)

    @property
    def active_document(self):
        return Document(self.api.ActiveDocument)
```

Usage:
```python
app = App()
print(f"Open documents: {len(app.documents)}")
for doc in app.documents:
    print(doc.full_name)
    doc.save()
new_doc = app.documents.add()
```

### Phase 3: Add Form Controls Support (Low Priority)

```python
class App:
    @property
    def forms(self):
        return FormsCollection(self.api)
```

## Restructuring Priorities

| Priority | Task | Lines Saved | Complexity Reduction | User Impact |
|----------|------|-------------|---------------------|-------------|
| **1** | Split parametersets.py by domain | 0 (reorg) | 🔼🔼🔼 High | Low (internal) |
| **2** | Unify backend modes | ~200 | 🔼🔼 Medium | Low (internal) |
| **3** | Expose Documents/Windows collections | +150 | 🔽 Slight increase | 🔼🔼 High (feature) |
| **4** | Consolidate property types | ~200 | 🔼 Medium | Low (internal) |
| **5** | Add Form controls support | +200 | 🔽 Slight increase | 🔼 Medium (feature) |
| **6** | Remove forward declarations | ~25 | 🔼 Small | None |

**Recommendation:** Start with Priority 1 (split parametersets.py): highest maintainability impact, zero breaking changes, easiest to implement.

## Migration Checklist

**Phase 1 Execution:**
- [ ] Create `hwpapi/parametersets/` package structure
- [ ] Move base classes to `base.py`
- [ ] Move backend classes to `backends.py`
- [ ] Move property descriptors to `properties.py`
- [ ] Move mappings to `mappings.py`
- [ ] Move ParameterSet subclasses to domain files
- [ ] Create `__init__.py` with full re-exports
- [ ] Test all imports
- [ ] Run full test suite
- [ ] Update CLAUDE.md file mapping table
- [ ] Commit changes

**Phase 2 Execution (Optional):**
- [ ] Design DocumentsCollection, WindowsCollection classes
- [ ] Add `app.documents`, `app.windows` properties
- [ ] Write tests for multi-document workflows

**Phase 3 Execution (Optional):**
- [ ] Design FormsCollection classes
- [ ] Add `app.forms` property
- [ ] Write tests for form controls

## Non-Goals

- ❌ Don't change the property descriptor system (it's excellent)
- ❌ Don't change the backend abstraction (it works well)
- ❌ Don't change the Actions pattern (pythonic and convenient)
- ❌ Don't add complexity for theoretical future needs
