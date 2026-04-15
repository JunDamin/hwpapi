"""Domain-grouped tests: lint."""
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


