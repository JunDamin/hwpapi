# Project Structure - Standard Python Package

## Source Files

Python files in `hwpapi/` are the **direct source code**. Edit them directly.

| File / Directory | Purpose |
|------------------|---------|
| `hwpapi/core.py` | App, Engine classes - main API |
| `hwpapi/actions.py` | HWP action wrappers (704 actions via __getattr__) |
| `hwpapi/parametersets/` | ParameterSet package (split from single file) |
| `hwpapi/parametersets/__init__.py` | ParameterSet base class, Meta, PropertyDescriptors, GenericParameterSet |
| `hwpapi/parametersets/mappings.py` | 35 string↔int MAPs (DIRECTION_MAP, ALIGN_MAP, etc.) |
| `hwpapi/parametersets/backends.py` | 4 backends (PsetBackend, HParamBackend, ComBackend, AttrBackend) |
| `hwpapi/parametersets/sets/primitives.py` | 8 base classes (CharShape, ParaShape, BorderFill, Cell, Caption, CtrlData, Password, Style) |
| `hwpapi/parametersets/sets/drawing.py` | 25 classes: ShapeObject, Draw* |
| `hwpapi/parametersets/sets/text.py` | 20 classes: character ops |
| `hwpapi/parametersets/sets/paragraph.py` | 4 classes: TabDef, NumberingShape, etc. |
| `hwpapi/parametersets/sets/table.py` | 12 classes: Table, CellBorderFill, etc. |
| `hwpapi/parametersets/sets/document.py` | 17 classes: DocumentInfo, PageDef, SecDef, etc. |
| `hwpapi/parametersets/sets/formatting.py` | 3 classes: BorderFillExt, StyleDelete, StyleTemplate |
| `hwpapi/parametersets/sets/file_ops.py` | 14 classes: FileOpen, Print, etc. |
| `hwpapi/parametersets/sets/find_edit.py` | 13 classes: FindReplace, BookMark, etc. |
| `hwpapi/parametersets/sets/media_misc.py` | 27 classes: OleCreation, HyperLink, Ftp*, etc. |
| `hwpapi/functions.py` | Utility functions |
| `hwpapi/classes.py` | Accessor classes (Move, Cell, Table, Page) |
| `hwpapi/constants.py` | Constants and enums |
| `hwpapi/logging.py` | Logging configuration |

See `docs/API_GUIDE.md` for the full API reference with recipes.

## Workflow for Code Changes

```bash
# 1. Edit the .py file directly
# 2. Test your changes
python -m pytest tests/

# 3. Commit
git add hwpapi/changed_file.py
git commit -m "Your change description"
```

## Note on nbs/ Directory

The `nbs/` directory contains archived Jupyter notebooks from the previous nbdev-based workflow. These are kept for reference only and are **not** the source of truth.

## Key Reference Tables

### Important Files

| File | Purpose |
|------|---------|
| `pyproject.toml` | Project configuration, version, dependencies |
| `CLAUDE.md` | Top-level Claude index |
| `claude_docs/` | Detailed topic guides (this folder) |
| `PSET_MIGRATION_SUMMARY.md` | Context on pset refactoring |
| `REFACTORING_SUMMARY.md` | Recent refactoring documentation |
| `DUPLICATE_FIX_SUMMARY.md` | Duplicate class bug fix and display formatting (2025-12-09) |
| `AUTO_PROPERTY_DESIGN.md` | NestedProperty & ArrayProperty design |
| `UNIT_PROPERTY_ENHANCEMENT.md` | Smart unit conversion specification |

### Key Classes

| Class | Location | Purpose |
|-------|----------|---------|
| `App` | `core.py` | High-level API for users |
| `Engine` | `core.py` | Mid-level wrapper around HwpObject |
| `ParameterSet` | `parametersets.py` | Base class for all parameter sets |
| `PropertyDescriptor` | `parametersets.py` | Base for property types |
| `ParameterSetMeta` | `parametersets.py` | Metaclass for auto-registration |

### Key Functions

| Function | Location | Purpose |
|----------|----------|---------|
| `_is_com(obj)` | `parametersets.py` | Check if object is COM |
| `_looks_like_pset(obj)` | `parametersets.py` | Check if object is pset |
| `make_backend(obj)` | `parametersets.py` | Create appropriate backend |
| `resolve_action_args()` | `parametersets.py` | Resolve action arguments |

## Codebase Metrics

**Current State:**
- Total lines: ~15,000
- parametersets.py: ~4,100 lines
- ParameterSet subclasses: 29
- Property descriptors: 438
- Action definitions: 899+
