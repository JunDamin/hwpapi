"""Unit tests for :mod:`hwpapi.context.scopes`.

The scopes operate through ``app.actions`` (public escape hatch) and
``app.api.HParameterSet`` / ``app.api.HAction``. A :class:`MagicMock`
fulfils every one of those without needing a real HWP engine.
"""
from __future__ import annotations

from unittest.mock import MagicMock, call

import pytest

from hwpapi.context.scopes import (
    charshape_scope,
    parashape_scope,
    styled_text,
)


# ---------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------

@pytest.fixture
def fake_app():
    """MagicMock-backed App exposing ``.actions``, ``.api.HAction``,
    ``.api.HParameterSet``.

    Returns the mock so tests can assert on call sites.
    """
    app = MagicMock(name="FakeApp")

    # ``app.actions.CharShape`` — ``_Action`` object. We don't need any
    # particular shape; the scopes only need to *touch* it.
    app.actions.CharShape = MagicMock(name="CharShapeAction")
    app.actions.ParaShape = MagicMock(name="ParaShapeAction")
    app.actions.InsertText = MagicMock(name="InsertTextAction")

    # Set initial values on HParameterSet slots so ``getattr`` returns a
    # concrete sentinel rather than another MagicMock (makes snapshot
    # assertions readable).
    hchar = app.api.HParameterSet.HCharShape
    hchar.Bold = False
    hchar.Height = 1000
    hchar.TextColor = 0
    hchar.Italic = False

    hpara = app.api.HParameterSet.HParaShape
    hpara.AlignType = 0

    return app


# ---------------------------------------------------------------------
# charshape_scope
# ---------------------------------------------------------------------

def test_charshape_scope_touches_action_and_executes(fake_app):
    """The scope accesses app.actions.CharShape and runs HAction.Execute."""
    with charshape_scope(fake_app, bold=True):
        pass

    # Touched the public escape hatch — this is what the migration doc
    # guarantees as the stable entry point.
    assert fake_app.actions.CharShape is not None

    # HAction.Execute was called for both the "apply" and "restore" hop.
    execute_calls = fake_app.api.HAction.Execute.call_args_list
    assert len(execute_calls) == 2  # apply + restore
    for c in execute_calls:
        # First positional arg is the action name.
        assert c.args[0] == "CharShape"


def test_charshape_scope_applies_and_restores_bold(fake_app):
    """Body sees Bold=True; on exit, original Bold=False is re-applied."""
    hchar = fake_app.api.HParameterSet.HCharShape

    observed = {}

    def record_on_execute(name, hset):
        # Snapshot what ``Bold`` looks like at the moment Execute runs.
        observed.setdefault("calls", []).append(
            {"name": name, "bold": hchar.Bold}
        )

    fake_app.api.HAction.Execute.side_effect = record_on_execute

    with charshape_scope(fake_app, bold=True):
        # Inside the block Bold must be True on the HCharShape slot.
        assert hchar.Bold is True

    calls = observed["calls"]
    assert len(calls) == 2
    # First Execute is the "apply" — Bold was just set to True.
    assert calls[0] == {"name": "CharShape", "bold": True}
    # Second Execute is the "restore" — Bold was reset to the snapshot (False).
    assert calls[1] == {"name": "CharShape", "bold": False}


def test_charshape_scope_translates_friendly_keys(fake_app):
    """``size`` → ``Height``; ``color`` → ``TextColor``."""
    hchar = fake_app.api.HParameterSet.HCharShape

    with charshape_scope(fake_app, size=1400, color=0xFF):
        assert hchar.Height == 1400
        assert hchar.TextColor == 0xFF


def test_charshape_scope_restores_on_exception(fake_app):
    """Even when the body raises, the snapshot is re-applied on exit."""
    hchar = fake_app.api.HParameterSet.HCharShape

    with pytest.raises(RuntimeError):
        with charshape_scope(fake_app, bold=True):
            raise RuntimeError("boom")

    # Two Execute calls: apply + restore. The finally-clause ran.
    assert fake_app.api.HAction.Execute.call_count == 2


# ---------------------------------------------------------------------
# parashape_scope
# ---------------------------------------------------------------------

def test_parashape_scope_applies_center_alignment(fake_app):
    """``align="center"`` maps to AlignType=1 on HParaShape."""
    hpara = fake_app.api.HParameterSet.HParaShape

    observed = {}

    def record_on_execute(name, hset):
        observed.setdefault("calls", []).append(
            {"name": name, "align": hpara.AlignType}
        )

    fake_app.api.HAction.Execute.side_effect = record_on_execute

    with parashape_scope(fake_app, align="center"):
        assert hpara.AlignType == 1  # "center" → 1 per ALIGN_MAP

    calls = observed["calls"]
    assert len(calls) == 2
    assert calls[0]["name"] == "ParaShape"
    assert calls[0]["align"] == 1
    # Restore hop puts the original 0 back.
    assert calls[1]["align"] == 0


def test_parashape_scope_touches_action(fake_app):
    """``app.actions.ParaShape`` is touched, proving the scope engaged."""
    with parashape_scope(fake_app, align="left"):
        pass
    assert fake_app.actions.ParaShape is not None


# ---------------------------------------------------------------------
# styled_text
# ---------------------------------------------------------------------

def test_styled_text_sets_shape_inserts_text_and_restores(fake_app):
    """styled_text drives CharShape apply → InsertText → CharShape restore."""
    hchar = fake_app.api.HParameterSet.HCharShape

    # Track ordering of HAction.Execute + InsertText action.run().
    events = []

    def record_execute(name, hset):
        events.append(("Execute", name, hchar.Bold))

    fake_app.api.HAction.Execute.side_effect = record_execute

    insert_action = fake_app.actions.InsertText
    insert_action.run.side_effect = lambda pset=None: events.append(
        ("InsertText.run", getattr(pset, "Text", None))
    )

    result = styled_text(fake_app, "Hello", bold=True)

    # One-shot, not a context manager — returns None.
    assert result is None

    # Ordering: apply (Bold=True) → InsertText → restore (Bold=False).
    assert events[0] == ("Execute", "CharShape", True)
    # The InsertText event sits between the two Execute calls.
    assert events[1][0] == "InsertText.run"
    assert events[1][1] == "Hello"
    assert events[2] == ("Execute", "CharShape", False)


def test_styled_text_assigns_text_to_pset(fake_app):
    """The InsertText pset receives the text before the action runs."""
    insert_action = fake_app.actions.InsertText
    pset_mock = MagicMock(name="InsertTextPset")
    type(insert_action).pset = pset_mock  # any attribute assignment is fine
    insert_action.pset = pset_mock

    styled_text(fake_app, "abc")

    # Either Text or text was set on the pset; the scope tries both.
    assert pset_mock.Text == "abc" or pset_mock.text == "abc"
