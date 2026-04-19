# hwpapi v1.0 Baseline Snapshot

Captured: 2026-04-19 (Phase 0 of v2 redesign)
Source commit: `main` @ HEAD (tagged as `v1.0.0`)
Branch: `v1.x` mirrors this state; `v2-redesign` is the working branch.

---

## Version

| Source | Value |
|---|---|
| `pyproject.toml` `[project].version` | `1.0.0` |
| `hwpapi/__init__.py` `__version__` | `"0.0.2.5"` ⚠️ stale |

> The `__version__` constant was not kept in sync with `pyproject.toml`. Tracked as a Phase 0.5 chore; corrected alongside Phase 7 release bump to `2.0.0`.

## Public API surface

| Metric | Value |
|---|---|
| Non-underscore symbols exposed at `import hwpapi` | 244 |
| Source of truth | [`public-api-v1.0.txt`](public-api-v1.0.txt) |
| Wildcard import chain in `hwpapi/__init__.py` | `from .core import *`, `from .actions import *`, `from .functions import *`, `from .parametersets import *` |

> Acceptance target for v2.0: **≤ 10** non-underscore top-level symbols (`App`, `Document`, `__version__`, and a handful of error classes only).

## Tests

| Metric | Value |
|---|---|
| `pytest --collect-only` items | **5,953** |
| Collection wall-clock | 17.40 s |
| Pass/fail baseline | See test_suite_baseline.txt once run completes |

### HWP-dependent layer

The test suite mixes pure-Python and HWP-integration tests. The latter group:

- `tests/smoke_features.py`, `smoke_scenarios.py`, `smoke_visual.py`
- `tests/test_all_actions.py`, `test_all_parametersets.py`
- Any test that spins up a real `HwpObject` COM instance

Per `CLAUDE.md`, HWP-dependent tests are expected to skip gracefully when HWP is not installed or visible. This baseline captures the status from the current developer machine (Windows 11 + HWP-available assumption).

## Baseline guarantees for Phase 1

During Phase 1 (`hwpapi.low/` relocation), the following must remain **identical** to this baseline:

1. `len(dir(hwpapi))` of non-underscore symbols — drift tolerance ±2
2. Every symbol listed in `public-api-v1.0.txt` must still be importable from `hwpapi` at the top level (via the new `from hwpapi.low.* import *` re-exports)
3. `pytest --collect-only` exits 0 with the same collection count (±5 for any test collection that touches module paths)
4. No new `ImportError` in any test file

Phase 1 does **not** change public behaviour — only relocates modules under `hwpapi/low/`. Any drift beyond these bounds is a Phase 1 regression.

## Files touched by the move (relocation map)

| From | To |
|---|---|
| `hwpapi/actions.py` | `hwpapi/low/actions.py` |
| `hwpapi/parametersets/` (whole subpackage) | `hwpapi/low/parametersets/` |
| `hwpapi/core/engine.py` | `hwpapi/low/engine.py` |

### Callers needing import updates

- `hwpapi/__init__.py` (4 wildcard lines)
- `hwpapi/core/__init__.py` (`from .engine import ...`)
- `hwpapi/core/app.py` (`from .engine import ...`)
- 9 files under `hwpapi/parametersets/sets/` (internal relative imports — preserved automatically when moved as a whole)
- 4 files under `hwpapi/parametersets/` (base/properties/mappings/backends — relative intra-package imports preserved)
- 7 test files referencing full-path `hwpapi.parametersets` / `hwpapi.actions`

Detailed grep count: **16** hwpapi internal files, **7** tests/*.py files.

## Changelog of Phase 0 itself

- 2026-04-19: `v1.0.0` annotated tag created.
- 2026-04-19: `v1.x` branch forked from the tagged commit.
- 2026-04-19: `v2-redesign` branch created from the same commit — the working branch.
- 2026-04-19: `pyproject.toml` `[project.optional-dependencies].docs` extended with `quartodoc`, `griffe`, `jupyter` for Phase 6.
- 2026-04-19: `docs/design/public-api-v1.0.txt` and `docs/design/baseline-v1.0.md` captured.
