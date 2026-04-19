"""hwpapi.low — low-level layer (escape hatch).

Direct access to the raw building blocks:

- `hwpapi.low.actions` — 900+ HWP action wrappers
- `hwpapi.low.parametersets` — ParameterSet classes (CharShape, ParaShape, ...)
- `hwpapi.low.engine` — Engine / Engines / Apps

High-level users should prefer `hwpapi.App`; this module is kept
for power users who need to drop down to raw HWP automation calls.
"""

from . import actions, engine, parametersets

__all__ = ["actions", "engine", "parametersets"]
