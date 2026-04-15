"""Domain-grouped tests: config."""
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


def test_config_default():
    app = MagicMock()
    c = Config(app)
    assert c.default_font is None
    assert c.palette == {}


def test_config_set_get():
    app = MagicMock()
    c = Config(app)
    c.default_font = "함초롬바탕"
    assert c.default_font == "함초롬바탕"


def test_config_update():
    app = MagicMock()
    c = Config(app)
    c.update(default_font="함초롬", default_size=11)
    assert c.default_font == "함초롬"
    assert c.default_size == 11


def test_config_reset():
    app = MagicMock()
    c = Config(app)
    c.default_font = "X"
    c.reset()
    assert c.default_font is None


def test_config_save_load_roundtrip(tmp_path):
    app = MagicMock()
    app.logger = MagicMock()
    c = Config(app)
    c.default_font = "함초롬"
    c.default_size = 11
    c.palette = {"primary": "#003366"}

    path = tmp_path / "hwpapirc.json"
    c.save(str(path))
    assert path.exists()

    c2 = Config(app).load(str(path))
    assert c2.default_font == "함초롬"
    assert c2.default_size == 11
    assert c2.palette == {"primary": "#003366"}


def test_config_load_missing_file():
    app = MagicMock()
    app.logger = MagicMock()
    # Should not raise
    Config(app).load("/nonexistent/hwpapirc.json")


def test_config_to_dict():
    app = MagicMock()
    c = Config(app).update(default_font="X", default_size=12)
    d = c.to_dict()
    assert d["default_font"] == "X"
    assert d["default_size"] == 12


def test_config_unknown_attr_raises():
    app = MagicMock()
    with pytest.raises(AttributeError):
        Config(app).not_a_real_setting


