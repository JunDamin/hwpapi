"""Domain-grouped tests: template."""
from hwpapi.classes.lint import Config, LintReport, Linter, Template
from pathlib import Path
from unittest.mock import MagicMock
import json
import pytest
import tempfile


@pytest.fixture
def lint_app(monkeypatch):
    app = MagicMock()
    app.logger = MagicMock()
    return app


@pytest.fixture
def tmpl_app():
    app = MagicMock()
    app.logger = MagicMock()
    cs = MagicMock()
    cs.fontsize = 1100
    cs.bold = True
    cs.italic = False
    cs.text_color = "#000000"
    cs.facename_hangul = "맑은 고딕"
    cs.facename_latin = "Arial"
    app.charshape = cs
    ps = MagicMock()
    ps.align_type = 0
    ps.indent = 0
    ps.line_spacing = 160
    ps.space_before = 0
    ps.space_after = 0
    app.parashape = ps
    return app


def test_template_save_writes_json(tmpl_app, tmp_path):
    path = tmp_path / "tmpl.json"
    data = Template(tmpl_app).save(str(path))
    assert path.exists()
    loaded = json.loads(path.read_text(encoding="utf-8"))
    assert loaded["format"] == "hwpapi-template"
    assert loaded["charshape"]["bold"] is True
    assert loaded["charshape"]["facename_hangul"] == "맑은 고딕"


def test_template_save_strips_nones(tmpl_app, tmp_path):
    tmpl_app.charshape.bold = None
    path = tmp_path / "tmpl.json"
    Template(tmpl_app).save(str(path))
    loaded = json.loads(path.read_text(encoding="utf-8"))
    assert "bold" not in loaded["charshape"]


def test_template_apply_calls_set(tmpl_app, tmp_path):
    tmpl = {
        "format": "hwpapi-template",
        "version": "1.0",
        "charshape": {"fontsize": 1400, "bold": True},
        "parashape": {"indent": 100, "align_type": 1},
        "page": {},
    }
    path = tmp_path / "tmpl.json"
    path.write_text(json.dumps(tmpl), encoding="utf-8")
    Template(tmpl_app).apply(str(path))
    # set_charshape called with fontsize+bold
    tmpl_app.set_charshape.assert_called()
    tmpl_app.set_parashape.assert_called()


def test_template_apply_missing_file(tmpl_app):
    result = Template(tmpl_app).apply("/nonexistent/path/template.json")
    # Should not crash; return empty
    assert result == {}


