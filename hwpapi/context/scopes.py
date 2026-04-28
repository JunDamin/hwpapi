"""
:mod:`hwpapi.context.scopes` — formatting context managers + one-shot helper.

Three public names, all module-level (so they work uniformly whether
the caller is driving an :class:`~hwpapi.core.app.App`, a nested
:class:`~hwpapi.document.Document` facade, or something else):

.. list-table::
   :header-rows: 1

   * - Name
     - Shape
     - Use when …
   * - :func:`styled_text`
     - one-shot (plain function, returns ``None``)
     - You want to insert a short run of text with a different style and
       have the cursor formatting snap back immediately afterwards.
   * - :func:`charshape_scope`
     - ``with`` block
     - You're doing several char-level operations (insert_text,
       run some action, set color, …) that all share one style and
       need the original style restored when the block ends.
   * - :func:`parashape_scope`
     - ``with`` block
     - Paragraph-level changes — alignment, line spacing, indentation.
       Same save/restore semantics as :func:`charshape_scope`.

Decision tree — "what should I pick?"::

    Single line of text with a tweak?            → styled_text(app, "...", ...)
    Multiple ops sharing char format?            → with charshape_scope(app, ...):
    Paragraph alignment/line-spacing block?      → with parashape_scope(app, ...):

All three drive formatting through the public escape hatch
``app.actions`` and the raw HWP COM ``HParameterSet`` / ``HAction`` —
no reliance on removed v1 App members.

Implementation notes
--------------------
Both scopes snapshot the cursor's current char/para shape by running
``HAction.GetDefault("CharShape", hset)`` and reading back the keys the
caller is about to override. On exit the snapshot is re-applied
verbatim, so the user's baseline formatting survives the block even
if the caller's code raised inside.

``styled_text`` is a plain function (not a context manager) — it opens
a :func:`charshape_scope`, issues a single ``InsertText`` action, then
lets the scope's ``__exit__`` restore.
"""
from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Dict, Iterable, Mapping, Optional

from hwpapi.logging import get_logger

logger = get_logger("context.scopes")


# ---------------------------------------------------------------------
# Property-name translation
# ---------------------------------------------------------------------
#
# Callers use pleasant Python names (``bold``, ``size``, ``color``,
# ``align``). Internally we set attributes on HWP's HCharShape /
# HParaShape COM objects, which expect CamelCase keys.
#
# Anything not in these maps is passed through as-is — advanced users
# can write ``charshape_scope(app, Height=1200)`` and it will Just Work.

_CHAR_ALIAS: Mapping[str, str] = {
    "bold":           "Bold",
    "italic":         "Italic",
    "underline":      "UnderlineType",
    "size":           "Height",         # HWP treats font size as Height (HWPUNIT).
    "height":         "Height",
    "color":          "TextColor",
    "text_color":     "TextColor",
    "shade_color":    "ShadeColor",
    "face":           "FaceNameHangul",
    "face_name":      "FaceNameHangul",
    "font":           "FaceNameHangul",
}

_PARA_ALIAS: Mapping[str, str] = {
    "align":              "AlignType",
    "alignment":          "AlignType",
    "line_spacing":       "LineSpacing",
    "left_margin":        "LeftMargin",
    "right_margin":       "RightMargin",
    "indentation":        "Indentation",
    "prev_spacing":       "PrevSpacing",
    "next_spacing":       "NextSpacing",
}

_ALIGN_WORDS: Mapping[str, int] = {
    "left": 0, "center": 1, "right": 2, "justify": 3, "distribute": 4,
}


def _translate(fmt: Mapping[str, Any], aliases: Mapping[str, str]) -> Dict[str, Any]:
    """Return ``fmt`` with friendly keys mapped to HWP COM property names."""
    out: Dict[str, Any] = {}
    for key, value in fmt.items():
        com_key = aliases.get(key, key)
        out[com_key] = value
    return out


def _normalise_align(value: Any) -> Any:
    """``"center"`` → ``1`` etc.; pass numbers through unchanged."""
    if isinstance(value, str):
        return _ALIGN_WORDS.get(value.lower(), value)
    return value


# ---------------------------------------------------------------------
# Snapshot / apply helpers
# ---------------------------------------------------------------------

def _hparameterset(app, slot: str):
    """Return ``app.api.HParameterSet.H<slot>`` — one place to fail early."""
    try:
        return getattr(app.api.HParameterSet, f"H{slot}")
    except Exception as exc:  # pragma: no cover — defensive, HWP should have it
        logger.debug(f"_hparameterset({slot!r}): {exc!r}")
        raise


def _snapshot(app, action_name: str, slot: str, keys: Iterable[str]) -> Dict[str, Any]:
    """Capture the cursor's current values for the named COM keys."""
    hpset = _hparameterset(app, slot)
    try:
        app.api.HAction.GetDefault(action_name, hpset.HSet)
    except Exception as exc:
        logger.debug(f"_snapshot GetDefault {action_name!r}: {exc!r}")

    snap: Dict[str, Any] = {}
    for k in keys:
        try:
            snap[k] = getattr(hpset, k)
        except Exception as exc:
            logger.debug(f"_snapshot read {k!r}: {exc!r}")
    return snap


def _apply(app, action_name: str, slot: str, values: Mapping[str, Any]) -> None:
    """Set values on the HParameterSet slot and Execute the action."""
    hpset = _hparameterset(app, slot)
    try:
        app.api.HAction.GetDefault(action_name, hpset.HSet)
    except Exception as exc:
        logger.debug(f"_apply GetDefault {action_name!r}: {exc!r}")

    for k, v in values.items():
        try:
            setattr(hpset, k, v)
        except Exception as exc:
            logger.debug(f"_apply set {k!r}={v!r}: {exc!r}")

    try:
        app.api.HAction.Execute(action_name, hpset.HSet)
    except Exception as exc:
        logger.debug(f"_apply Execute {action_name!r}: {exc!r}")


# ---------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------

@contextmanager
def charshape_scope(app, **fmt: Any):
    """Apply a char-shape override inside the ``with`` block, then restore.

    Parameters
    ----------
    app : hwpapi.App
        Live App instance (or any object exposing ``.actions`` and
        ``.api``).
    **fmt
        Char-shape overrides. Friendly keys — ``bold``, ``italic``,
        ``size``, ``color``, ``shade_color``, ``font`` — are translated
        to HWP COM names; anything else is passed through.

    Example
    -------
    >>> with charshape_scope(app, bold=True, size=1400, color="#FF0000"):
    ...     app.doc.insert_text("important!")
    """
    # Touching the action resolves it through the cache and (on real
    # HWP) ensures the CharShape pset class is registered. Tests rely on
    # this being called — it's a cheap and unambiguous "the scope
    # actually started".
    _ = app.actions.CharShape

    translated = _translate(fmt, _CHAR_ALIAS)
    snap = _snapshot(app, "CharShape", "CharShape", translated.keys())
    _apply(app, "CharShape", "CharShape", translated)
    try:
        yield
    finally:
        if snap:
            _apply(app, "CharShape", "CharShape", snap)


@contextmanager
def parashape_scope(app, **fmt: Any):
    """Apply a para-shape override inside the ``with`` block, then restore.

    Parameters
    ----------
    app : hwpapi.App
    **fmt
        Paragraph-shape overrides. Friendly keys — ``align``,
        ``line_spacing``, ``left_margin``, ``right_margin``,
        ``indentation`` — are translated to HWP COM names.

    Example
    -------
    >>> with parashape_scope(app, align="center"):
    ...     app.doc.insert_text("centred paragraph")
    """
    _ = app.actions.ParaShape

    translated = _translate(fmt, _PARA_ALIAS)
    if "AlignType" in translated:
        translated["AlignType"] = _normalise_align(translated["AlignType"])

    snap = _snapshot(app, "ParaShape", "ParaShape", translated.keys())
    _apply(app, "ParaShape", "ParaShape", translated)
    try:
        yield
    finally:
        if snap:
            _apply(app, "ParaShape", "ParaShape", snap)


def styled_text(app, text: str, **fmt: Any) -> None:
    """Insert ``text`` with a temporary char-shape, then restore.

    Shorthand for ::

        with charshape_scope(app, **fmt):
            app.api.HAction.Run("InsertText") / InsertText action

    The function returns ``None`` — it's a side-effecting one-shot, not
    a context manager.

    Parameters
    ----------
    app : hwpapi.App
    text : str
        Literal text to insert at the cursor.
    **fmt
        Same friendly-key char-shape overrides as :func:`charshape_scope`.

    Example
    -------
    >>> styled_text(app, "Hello", bold=True, color="#FF0000")
    """
    with charshape_scope(app, **fmt):
        _insert_text(app, text)


def _insert_text(app, text: str) -> None:
    """InsertText action — uses the public ``app.actions`` escape hatch."""
    try:
        action = app.actions.InsertText
        pset = action.pset
        # ``pset.Text`` is the canonical key; try the legacy ``text`` too.
        for key in ("Text", "text"):
            try:
                setattr(pset, key, text)
                break
            except Exception:  # pragma: no cover — try the next key
                continue
        action.run(pset)
    except Exception as exc:
        logger.debug(f"_insert_text failed: {exc!r}")
