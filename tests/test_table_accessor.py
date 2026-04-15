"""Domain-grouped tests: table_accessor."""
from hwpapi.classes.debug import Debug
from unittest.mock import MagicMock, patch
import pytest


@pytest.fixture
def dbg_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.api.GetPos.return_value = (5, 10, 17)
    app.current_page = 3
    app.page_count = 10
    app.selection = "hello world"
    cs = MagicMock()
    cs.fontsize = 1100
    cs.bold = True
    cs.italic = False
    cs.text_color = "#000000"
    cs.shade_color = None
    app.charshape = cs
    app.in_table.return_value = False
    app.documents = [MagicMock(), MagicMock()]
    app.visible = True
    app.version = "13.0.0"
    app.get_filepath.return_value = "test.hwp"
    return app


@pytest.fixture
def tbl_app():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = True
    app.api.Run.return_value = True
    return app


def test_header_row_needs_table():
    app = MagicMock()
    app.logger = MagicMock()
    app.in_table.return_value = False
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(app).header_row(bold=True)
    app.logger.warning.assert_called()


def test_header_row_applies_bg(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).header_row(bg="#E8F4F8")
    # CellFill CreateAction attempted
    tbl_app.api.CreateAction.assert_any_call("CellFill")


def test_footer_row_applies_bg(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).footer_row(bg="#EEEEEE")
    tbl_app.api.CreateAction.assert_any_call("CellFill")


def test_current_row_applies(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).current_row(bg="#FFFF00")
    tbl_app.api.CreateAction.assert_any_call("CellFill")


def test_align_center(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).align(horz="center", scope="current_cell")
    calls = [c.args[0] for c in tbl_app.api.Run.call_args_list if c.args]
    assert "ParagraphAlignCenter" in calls


def test_align_vertical_middle(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).align(vert="middle", scope="all")
    calls = [c.args[0] for c in tbl_app.api.Run.call_args_list if c.args]
    assert "CellAlignCenter" in calls


def test_align_scope_col(tbl_app):
    from hwpapi.classes.accessors import TableAccessor
    TableAccessor(tbl_app).align(horz="right", scope="current_col")
    calls = [c.args[0] for c in tbl_app.api.Run.call_args_list if c.args]
    assert "TableCellBlockCol" in calls
    assert "ParagraphAlignRight" in calls


def test_batch_chain(tbl_app):
    """Methods return self for chaining."""
    from hwpapi.classes.accessors import TableAccessor
    t = TableAccessor(tbl_app)
    r = t.header_row(bg="#CCC").align(horz="center", scope="all")
    assert r is t


