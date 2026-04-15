"""Domain-grouped tests: debug."""
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


def test_debug_state_has_cursor(dbg_app):
    s = Debug(dbg_app).state()
    assert s["cursor"] == {"doc_id": 5, "para": 10, "pos": 17}


def test_debug_state_has_page(dbg_app):
    s = Debug(dbg_app).state()
    assert s["page"] == {"current": 3, "total": 10}


def test_debug_state_has_charshape_summary(dbg_app):
    s = Debug(dbg_app).state()
    assert s["charshape_summary"]["fontsize"] == 1100
    assert s["charshape_summary"]["bold"] is True


def test_debug_state_has_in_table(dbg_app):
    s = Debug(dbg_app).state()
    assert s["in_table"] is False


def test_debug_state_has_documents_open(dbg_app):
    s = Debug(dbg_app).state()
    assert s["documents_open"] == 2


def test_debug_state_has_version(dbg_app):
    s = Debug(dbg_app).state()
    assert s["version"] == "13.0.0"


def test_debug_state_has_filepath(dbg_app):
    s = Debug(dbg_app).state()
    assert s["filepath"] == "test.hwp"


def test_debug_state_error_handled():
    """Broken props should not crash state()."""
    app = MagicMock()
    app.logger = MagicMock()
    app.api.GetPos.side_effect = Exception("boom")
    app.current_page = property(lambda s: (_ for _ in ()).throw(Exception("boom")))
    app.selection = ""
    app.charshape = None
    app.in_table.side_effect = Exception("boom")
    # Should not raise
    s = Debug(app).state()
    assert isinstance(s, dict)


def test_debug_print_does_not_crash(dbg_app, capsys):
    Debug(dbg_app).print()
    out = capsys.readouterr().out
    assert "hwpapi debug state" in out
    assert "cursor" in out


def test_debug_timing_success(dbg_app):
    def work():
        return 42

    r = Debug(dbg_app).timing(work)
    assert r["result"] == 42
    assert r["success"] is True
    assert "elapsed_ms" in r
    assert r["exception"] is None


def test_debug_timing_failure(dbg_app):
    def boom():
        raise ValueError("x")

    r = Debug(dbg_app).timing(boom)
    assert r["success"] is False
    assert isinstance(r["exception"], ValueError)


def test_debug_trace_wraps_run(capsys):
    """trace() should override api.Run temporarily."""
    from hwpapi.classes.debug import Debug

    class FakeApi:
        def Run(self, name):
            return True

    class FakeApp:
        api = FakeApi()

    app = FakeApp()
    d = Debug(app)
    with d.trace(verbose=True):
        app.api.Run("InsertText")
    out = capsys.readouterr().out
    assert "Run('InsertText')" in out
    # After exit, further Run calls should not emit trace output
    app.api.Run("NotTraced")
    after = capsys.readouterr().out
    assert "NotTraced" not in after


