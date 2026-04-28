# Claude's Guide to Working with hwpapi

A thin index. Read this first, then dive into `claude_docs/` for the topic you need.

## Golden Rules

1. **Standard Python project** — files in `hwpapi/` are the source of truth. Edit `.py` files directly.
2. **Test after every change** — run `python -m pytest tests/ -v`.
3. **Use property descriptors** — never set `self.attributes_names = [...]` in subclasses.
4. **Trust the backend factory** — `make_backend()` auto-detects pset / HSet / COM / attr.
5. **Always check for None backend** — unbound ParameterSets are valid.
6. **Default log level is `WARNING`** — set `HWPAPI_LOG_LEVEL=DEBUG` for development.

## Topic Index

Detailed information lives in `claude_docs/`. Pick the file that matches your task:

| File | When to read |
|------|--------------|
| [`claude_docs/project-structure.md`](claude_docs/project-structure.md) | Source layout, file/class/function reference tables, codebase metrics |
| [`claude_docs/architecture.md`](claude_docs/architecture.md) | Backend abstraction, property descriptor system, staging vs immediate mode, HWP/win32com domain notes |
| [`claude_docs/code-patterns.md`](claude_docs/code-patterns.md) | 5 patterns: adding a ParameterSet, NestedProperty, custom property, COM checks, optional backend |
| [`claude_docs/auto-properties.md`](claude_docs/auto-properties.md) | NestedProperty, UnitProperty, ArrayProperty — full reference with migration guide & decision tree |
| [`claude_docs/display-enhancements.md`](claude_docs/display-enhancements.md) | `__repr__` enhancements: human-readable values, enum display, description comments |
| [`claude_docs/issues-and-solutions.md`](claude_docs/issues-and-solutions.md) | Common errors and fixes (NameError `_is_com`, attributes_names setter, None backend, duplicates) |
| [`claude_docs/testing.md`](claude_docs/testing.md) | Test structure, running tests, writing new tests, auto-yes dialog mode |
| [`claude_docs/debugging.md`](claude_docs/debugging.md) | Logging env vars, import errors, test failures, duplicate-method detection script |
| [`claude_docs/best-practices.md`](claude_docs/best-practices.md) | DO / DON'T lists, dev workflow, simplification strategy, contributor checklist |
| [`claude_docs/architecture-roadmap.md`](claude_docs/architecture-roadmap.md) | Official HWP object model vs current architecture, gap analysis, 3-phase restructuring plan |
| [`claude_docs/history.md`](claude_docs/history.md) | Lessons learned, version history, prior pitfalls |

## Quick Reference

```bash
# Run tests
python -m pytest tests/ -v

# Verify imports
python -c "import hwpapi; print('OK')"

# Install in dev mode
pip install -e .
```

---

*This document is a living index. When adding new topics, create a file under `claude_docs/` and link it in the table above. Keep this file under ~60 lines.*

**Last updated:** 2026-04-19 (split into claude_docs/)
