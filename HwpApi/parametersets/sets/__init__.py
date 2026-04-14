"""
ParameterSet subclasses organized by HWP functional domain.

Import order respects the dependency graph:
1. primitives.py  — foundation classes (CharShape, ParaShape, etc.)
2. All other domains (each may import from primitives)

All classes are auto-registered to ``PARAMETERSET_REGISTRY`` via
``ParameterSetMeta`` at definition time, regardless of import order.
"""
from __future__ import annotations

# Import primitives FIRST (other domains depend on them)
from .primitives import *  # noqa: F401,F403

# Domain modules (alphabetical after primitives)
from .document import *    # noqa: F401,F403
from .drawing import *     # noqa: F401,F403
from .file_ops import *    # noqa: F401,F403
from .find_edit import *   # noqa: F401,F403 (imports CharShape/ParaShape from primitives)
from .formatting import *  # noqa: F401,F403
from .media_misc import *  # noqa: F401,F403
from .paragraph import *   # noqa: F401,F403
from .table import *       # noqa: F401,F403
from .text import *        # noqa: F401,F403
