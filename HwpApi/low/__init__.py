"""hwpapi.low — low-level layer (escape hatch).

Direct access to the raw building blocks:

- `hwpapi.low.actions` — 900+ HWP action wrappers
- `hwpapi.low.parametersets` — ParameterSet classes (CharShape, ParaShape, ...)
- `hwpapi.low.engine` — Engine / Engines / Apps

High-level users should prefer `hwpapi.App` (Phase 2+); this namespace
is the escape hatch for dropping down to raw HWP automation calls.
"""

from . import actions, engine, parametersets

__all__ = ["actions", "engine", "parametersets"]
