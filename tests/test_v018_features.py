"""Test v0.0.18 — Linter, Template, Config."""
import json
import tempfile
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from hwpapi.classes.lint import Config, LintReport, Linter, Template


# ═════════════════════════════════════════════════════════════════
# Linter
# ═════════════════════════════════════════════════════════════════

@pytest.fixture
def lint_app(monkeypatch):
    app = MagicMock()
    app.logger = MagicMock()
    return app


def test_lint_empty_document(lint_app):
    lint_app.text = ""
    r = Linter(lint_app)()
    assert r.total_chars == 0
    # split("") → [""] → 1 empty paragraph
    assert r.total_paragraphs == 1
    assert not r.long_sentences


def test_lint_detects_long_sentence(lint_app):
    lint_app.text = "이것은 " + "아주 " * 50 + "긴 문장입니다."
    r = Linter(lint_app)()
    assert r.long_sentences
    assert r.long_sentences[0][0] == 0   # paragraph 0


def test_lint_detects_empty_paragraphs(lint_app):
    lint_app.text = "첫 문단\n\n세 번째 문단\n"
    r = Linter(lint_app)()
    assert 1 in r.empty_paragraphs or 3 in r.empty_paragraphs


def test_lint_detects_double_spaces(lint_app):
    lint_app.text = "여기에  이중 공백이 있음"
    r = Linter(lint_app)()
    assert r.double_spaces


def test_lint_detects_trailing_whitespace(lint_app):
    lint_app.text = "공백 뒤따름   "
    r = Linter(lint_app)()
    assert r.trailing_whitespace


def test_lint_long_paragraph(lint_app):
    lint_app.text = "가" * 600
    r = Linter(lint_app)()
    assert r.long_paragraphs
    assert r.long_paragraphs[0][1] == 600


def test_lint_report_has_issues(lint_app):
    lint_app.text = "여기에  double space"
    r = Linter(lint_app)()
    assert r.has_issues is True


def test_lint_report_issue_count(lint_app):
    lint_app.text = "여기에  double.\n\n"
    r = Linter(lint_app)()
    assert r.issue_count >= 1


def test_lint_report_summary(lint_app):
    lint_app.text = "Simple text."
    r = Linter(lint_app)()
    assert "Document lint" in r.summary
    assert "paragraphs" in r.summary


def test_lint_custom_threshold(lint_app):
    lint_app.text = "이것은 테스트 문장."
    r = Linter(lint_app)(long_sentence_threshold=5)
    # 5자 임계 → "이것은 테스트 문장" (9자) 초과
    assert r.long_sentences


# ═════════════════════════════════════════════════════════════════
# Template
# ═════════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════════
# Config
# ═════════════════════════════════════════════════════════════════

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
