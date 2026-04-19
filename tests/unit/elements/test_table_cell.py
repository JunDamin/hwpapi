"""Unit tests for :class:`hwpapi.collections.tables.Table` / :class:`Cell`."""
from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

from hwpapi.collections.tables import Cell, Table


def _make_ctrl(rows=0, cols=0, user_desc="", ctrl_id="tbl "):
    """Build a MagicMock shaped like an HWP table control."""
    ctrl = MagicMock()
    ctrl.CtrlID = ctrl_id
    ctrl.UserDesc = user_desc
    ctrl.Next = None
    props = MagicMock()
    props.Rows = rows
    props.Cols = cols
    ctrl.Properties = props
    return ctrl


def _app(impl=None):
    impl = impl or MagicMock()
    return SimpleNamespace(engine=SimpleNamespace(impl=impl))


def test_table_rows_cols_from_ctrl_pset():
    ctrl = _make_ctrl(rows=5, cols=3)
    t = Table(_app(), ctrl, index=0)
    assert t.rows == 5
    assert t.cols == 3


def test_table_rows_cols_default_to_zero_when_unknown():
    ctrl = MagicMock()
    ctrl.CtrlID = "tbl "
    # No Properties attribute — fallback to 0.
    del ctrl.Properties
    del ctrl.Param
    t = Table(_app(), ctrl, index=0)
    assert t.rows == 0
    assert t.cols == 0


def test_cell_returns_cell_with_correct_coords():
    ctrl = _make_ctrl(rows=3, cols=3)
    t = Table(_app(), ctrl, index=0)
    c = t.cell(1, 2)
    assert isinstance(c, Cell)
    assert c.row == 1
    assert c.col == 2


def test_cell_address_is_a1_style():
    ctrl = _make_ctrl(rows=3, cols=3)
    t = Table(_app(), ctrl, index=0)
    assert t.cell(0, 0).address == "A1"
    assert t.cell(2, 1).address == "B3"
    assert t.cell(0, 25).address == "Z1"
    # AA — 26th column (0-indexed 26)
    assert t.cell(0, 26).address == "AA1"


def test_cell_select_calls_set_cell_addr_when_available():
    impl = MagicMock()
    impl.SetCellAddr = MagicMock(return_value=True)
    ctrl = _make_ctrl(rows=3, cols=3)
    app = _app(impl)
    t = Table(app, ctrl, index=0)
    c = t.cell(1, 2)
    assert c.select() is True
    impl.SelectCtrl.assert_called_once_with(ctrl)
    impl.SetCellAddr.assert_called_once_with("C2")


def test_cell_select_returns_false_when_table_select_fails():
    impl = MagicMock()
    impl.SelectCtrl.side_effect = Exception("boom")
    ctrl = _make_ctrl(rows=3, cols=3)
    app = _app(impl)
    t = Table(app, ctrl, index=0)
    assert t.cell(0, 0).select() is False


def test_cell_select_falls_back_to_navigation_without_set_cell_addr():
    impl = MagicMock(spec=["SelectCtrl", "Run"])
    # No SetCellAddr — fallback path exercised.
    ctrl = _make_ctrl(rows=3, cols=3)
    app = _app(impl)
    t = Table(app, ctrl, index=0)
    c = t.cell(2, 1)
    assert c.select() is True
    # Should issue TableColBegin once + 2 TableLowerCell + 1 TableRightCell
    calls = [call.args for call in impl.Run.call_args_list]
    assert ("TableColBegin",) in calls
    assert calls.count(("TableLowerCell",)) == 2
    assert calls.count(("TableRightCell",)) == 1


def test_cell_text_returns_saveblock_text():
    impl = MagicMock()
    impl.SetCellAddr = MagicMock(return_value=True)
    impl.GetTextFile = MagicMock(return_value="매출액\n")
    ctrl = _make_ctrl(rows=2, cols=2)
    app = _app(impl)
    t = Table(app, ctrl, index=0)
    assert t.cell(0, 0).text == "매출액"


def test_cell_text_empty_when_select_fails():
    impl = MagicMock()
    impl.SelectCtrl.side_effect = Exception("boom")
    ctrl = _make_ctrl(rows=2, cols=2)
    app = _app(impl)
    t = Table(app, ctrl, index=0)
    assert t.cell(0, 0).text == ""


def test_cell_repr_mentions_address_and_table_index():
    ctrl = _make_ctrl(rows=2, cols=2)
    t = Table(_app(), ctrl, index=5)
    c = t.cell(0, 1)
    r = repr(c)
    assert "B1" in r
    assert "#5" in r
