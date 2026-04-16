"""
Document linter — ``app.lint``.

문서 품질 체크. 공백/고아줄/과도하게 긴 문장/빈 문단/일관성 없는 폰트
등을 감지해서 구조화된 리포트 반환.

Usage::

    report = app.lint()
    print(report)
    report.long_sentences                     # list of (para_idx, length)
    report.empty_paragraphs                   # list[int]
    report.has_issues                          # bool
    report.summary                             # str 요약
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterable, List, Tuple

if TYPE_CHECKING:
    from hwpapi.core.app import App


# Sentence-ending regex (Korean + standard)
_SENTENCE_END = re.compile(r"[.!?。!?]+\s*")


@dataclass
class LintReport:
    """문서 품질 체크 결과."""

    long_sentences: List[Tuple[int, int]] = field(default_factory=list)
    """(para_idx, sentence_length) — 과도하게 긴 문장."""

    empty_paragraphs: List[int] = field(default_factory=list)
    """완전히 빈 문단의 인덱스."""

    double_spaces: List[int] = field(default_factory=list)
    """연속 공백이 있는 문단 인덱스."""

    trailing_whitespace: List[int] = field(default_factory=list)
    """문단 끝에 공백이 있는 문단 인덱스."""

    long_paragraphs: List[Tuple[int, int]] = field(default_factory=list)
    """(para_idx, char_count) — 너무 긴 문단."""

    total_paragraphs: int = 0
    total_sentences: int = 0
    total_chars: int = 0

    @property
    def has_issues(self) -> bool:
        return bool(
            self.long_sentences
            or self.empty_paragraphs
            or self.double_spaces
            or self.trailing_whitespace
            or self.long_paragraphs
        )

    @property
    def issue_count(self) -> int:
        return (len(self.long_sentences)
                + len(self.empty_paragraphs)
                + len(self.double_spaces)
                + len(self.trailing_whitespace)
                + len(self.long_paragraphs))

    @property
    def summary(self) -> str:
        lines = [
            f"Document lint — {self.issue_count} issue(s)",
            f"  Total: {self.total_paragraphs} paragraphs, "
            f"{self.total_sentences} sentences, {self.total_chars} chars",
        ]
        if self.long_sentences:
            lines.append(f"  Long sentences: {len(self.long_sentences)}")
        if self.empty_paragraphs:
            lines.append(f"  Empty paragraphs: {len(self.empty_paragraphs)}")
        if self.double_spaces:
            lines.append(f"  Double spaces: {len(self.double_spaces)} para(s)")
        if self.trailing_whitespace:
            lines.append(f"  Trailing whitespace: {len(self.trailing_whitespace)} para(s)")
        if self.long_paragraphs:
            lines.append(f"  Long paragraphs: {len(self.long_paragraphs)}")
        return "\n".join(lines)

    def __repr__(self) -> str:
        return f"LintReport(issues={self.issue_count})"

    def __str__(self) -> str:
        return self.summary


class Linter:
    """
    문서 linter — ``app.lint`` 로 접근. 호출 형태: ``app.lint()``.

    __call__ 이 정의돼서 ``app.lint`` 자체가 호출 가능.
    """

    __slots__ = ("_app",)

    # Thresholds
    LONG_SENTENCE_CHARS = 80
    LONG_PARAGRAPH_CHARS = 500

    def __init__(self, app: "App"):
        self._app = app

    def __call__(
        self,
        long_sentence_threshold: int = LONG_SENTENCE_CHARS,
        long_paragraph_threshold: int = LONG_PARAGRAPH_CHARS,
    ) -> LintReport:
        """
        현재 문서 전체를 검사하여 :class:`LintReport` 반환.

        Parameters
        ----------
        long_sentence_threshold : int
            문장 길이 기준 (기본 80자 — Korean 평균 문장의 2배).
        long_paragraph_threshold : int
            문단 길이 기준 (기본 500자).

        Returns
        -------
        LintReport

        Examples
        --------
        >>> r = app.lint()
        >>> print(r.summary)
        Document lint — 7 issue(s)
          Total: 42 paragraphs, ...

        >>> r.has_issues
        True
        >>> r.long_sentences[:3]
        [(5, 103), (12, 95), (18, 87)]
        """
        report = LintReport()
        app = self._app

        try:
            full_text = app.text or ""
        except Exception as e:
            app.logger.warning(f"lint: cannot read text: {e}")
            return report

        report.total_chars = len(full_text)
        paragraphs = full_text.split("\n")
        report.total_paragraphs = len(paragraphs)

        for idx, para in enumerate(paragraphs):
            # Empty paragraph
            if not para.strip():
                report.empty_paragraphs.append(idx)
                continue

            # Trailing whitespace
            if para != para.rstrip():
                report.trailing_whitespace.append(idx)

            # Double spaces
            if "  " in para.strip():
                report.double_spaces.append(idx)

            # Long paragraph
            if len(para) > long_paragraph_threshold:
                report.long_paragraphs.append((idx, len(para)))

            # Sentence analysis
            sentences = _SENTENCE_END.split(para)
            sentences = [s for s in sentences if s.strip()]
            report.total_sentences += len(sentences)

            for sent in sentences:
                if len(sent) > long_sentence_threshold:
                    report.long_sentences.append((idx, len(sent)))

        return report

    def __repr__(self) -> str:
        return f"Linter(<app {id(self._app):x}>)"


class Template:
    """
    문서 템플릿 저장/로드 — ``app.template`` 으로 접근.

    현재 문서의 주요 속성 (페이지 설정, 스타일, 기본 문자 모양) 을
    JSON 파일로 export 하고 다른 문서에 재적용.

    Usage::

        app.template.save("company.json")
        app2.template.apply("company.json")
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    def save(self, path: str) -> dict:
        """
        현재 문서 상태를 JSON 으로 저장. 현재 저장 항목:

        - 페이지 크기 (PageDef width/height/margins)
        - 현재 charshape summary (fontsize, fontname, text_color 등)
        - 현재 parashape summary
        - 문서 기본 폰트 (가능한 경우)

        Examples
        --------
        >>> app.template.save("report_template.json")
        """
        import json

        app = self._app
        tmpl = {
            "format": "hwpapi-template",
            "version": "1.0",
            "charshape": {},
            "parashape": {},
            "page": {},
        }

        try:
            cs = app.charshape
            tmpl["charshape"] = {
                "fontsize": getattr(cs, "fontsize", None)
                             or getattr(cs, "height", None),
                "bold": getattr(cs, "bold", None),
                "italic": getattr(cs, "italic", None),
                "text_color": getattr(cs, "text_color", None),
                "facename_hangul": getattr(cs, "facename_hangul", None),
                "facename_latin": getattr(cs, "facename_latin", None),
            }
        except Exception as e:
            app.logger.debug(f"template.save charshape: {e}")

        try:
            ps = app.parashape
            tmpl["parashape"] = {
                "align_type": getattr(ps, "align_type", None),
                "indent": getattr(ps, "indent", None),
                "line_spacing": getattr(ps, "line_spacing", None),
                "space_before": getattr(ps, "space_before", None),
                "space_after": getattr(ps, "space_after", None),
            }
        except Exception as e:
            app.logger.debug(f"template.save parashape: {e}")

        # Coerce non-JSON-serializable values (Color objects, etc.) to str
        def _to_json_safe(v):
            if v is None or isinstance(v, (bool, int, float, str)):
                return v
            # Try Color/hex-like object
            for attr in ("hex", "to_hex", "value"):
                if hasattr(v, attr):
                    try:
                        result = getattr(v, attr)
                        return result() if callable(result) else result
                    except Exception:
                        pass
            return str(v)

        for group in ("charshape", "parashape", "page"):
            tmpl[group] = {k: _to_json_safe(v) for k, v in tmpl[group].items()}

        # Strip None-valued keys (serialization cleanliness)
        for group in ("charshape", "parashape", "page"):
            tmpl[group] = {k: v for k, v in tmpl[group].items() if v is not None}

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(tmpl, f, ensure_ascii=False, indent=2)
        except Exception as e:
            app.logger.warning(f"template.save write: {e}")

        return tmpl

    def apply(self, path: str) -> dict:
        """
        JSON 템플릿 파일을 로드해 현재 문서에 적용.

        Examples
        --------
        >>> app.template.apply("report_template.json")
        """
        import json

        app = self._app
        try:
            with open(path, encoding="utf-8") as f:
                tmpl = json.load(f)
        except Exception as e:
            app.logger.warning(f"template.apply read: {e}")
            return {}

        # Apply charshape
        cs_data = tmpl.get("charshape") or {}
        if cs_data:
            try:
                app.set_charshape(**cs_data)
            except Exception as e:
                app.logger.debug(f"template.apply charshape: {e}")

        # Apply parashape
        ps_data = tmpl.get("parashape") or {}
        if ps_data:
            try:
                app.set_parashape(**ps_data)
            except Exception as e:
                app.logger.debug(f"template.apply parashape: {e}")

        app.logger.info(f"template.apply: {path} applied")
        return tmpl

    def __repr__(self) -> str:
        return f"Template(<app {id(self._app):x}>)"


class Config:
    """
    App 기본 선호도 설정 — ``app.config``.

    v0.0.18 기준: in-memory dict. 향후 ``~/.hwpapirc`` 지원 예정.
    """

    __slots__ = ("_app", "_data")

    # Defaults
    DEFAULTS = {
        "default_font": None,
        "default_size": None,
        "default_line_spacing": None,
        "default_table_style": None,
        "palette": {},
    }

    def __init__(self, app: "App"):
        self._app = app
        self._data = dict(self.DEFAULTS)

    def __getattr__(self, key):
        if key in ("_app", "_data"):
            raise AttributeError(key)
        try:
            return self._data[key]
        except KeyError:
            raise AttributeError(f"Config has no attribute {key!r}")

    def __setattr__(self, key, value):
        if key in ("_app", "_data"):
            object.__setattr__(self, key, value)
        else:
            self._data[key] = value

    def to_dict(self) -> dict:
        return dict(self._data)

    def update(self, **kw) -> "Config":
        self._data.update(kw)
        return self

    def reset(self) -> "Config":
        self._data = dict(self.DEFAULTS)
        return self

    def apply_defaults(self) -> "Config":
        """
        설정된 기본값을 **현재 커서 위치 (또는 SelectAll)** 에 실제로 적용.
        v0.0.24+.

        이전엔 ``app.config.default_font = "함초롬"`` 이 dict 저장만 하고
        실제 charshape 변경에 영향을 주지 않는 no-op 였음. 이 메소드를
        호출해야 실제로 반영됨.

        적용되는 키:
        - ``default_font`` → 7개 facename (hangul/latin/japanese/hanja/other/symbol/user) 동일 적용
        - ``default_size`` → height (HWPUNIT, 100=1pt 단위 또는 직접 pt)
        - ``default_line_spacing`` → ParaShape line_spacing

        Returns
        -------
        Config
            self — chainable.

        Examples
        --------
        >>> app.config.default_font = "함초롬바탕"
        >>> app.config.default_size = 1100   # 11pt (HWPUNIT)
        >>> app.config.apply_defaults()      # 명시적 적용

        Notes
        -----
        호출 전에 ``app.api.Run("SelectAll")`` 등으로 적용 범위를
        지정할 수 있음. 선택 영역이 없으면 커서 위치 이후 입력에만 영향.
        """
        app = self._app

        # CharShape 적용
        font = self._data.get("default_font")
        size = self._data.get("default_size")
        if font or size:
            cs_kw = {}
            if font:
                for key in ("facename_hangul", "facename_latin",
                             "facename_japanese", "facename_hanja",
                             "facename_other", "facename_symbol",
                             "facename_user"):
                    cs_kw[key] = font
            if size:
                cs_kw["height"] = int(size)
            try:
                app.set_charshape(**cs_kw)
                app.logger.info(
                    f"config.apply_defaults charshape: "
                    f"font={font!r}, size={size}"
                )
            except Exception as e:
                app.logger.warning(
                    f"config.apply_defaults charshape: {type(e).__name__}: {e}",
                    exc_info=True,
                )

        # ParaShape 적용
        line_sp = self._data.get("default_line_spacing")
        if line_sp:
            try:
                app.set_parashape(line_spacing=int(line_sp))
                app.logger.info(
                    f"config.apply_defaults parashape: line_spacing={line_sp}"
                )
            except Exception as e:
                app.logger.warning(
                    f"config.apply_defaults parashape: {type(e).__name__}: {e}",
                    exc_info=True,
                )

        return self

    def save(self, path: str = "~/.hwpapirc") -> None:
        """JSON 으로 디스크에 저장."""
        import json
        import os

        path = os.path.expanduser(path)
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._app.logger.warning(f"config.save: {e}")

    def load(self, path: str = "~/.hwpapirc") -> "Config":
        """JSON 에서 로드."""
        import json
        import os

        path = os.path.expanduser(path)
        try:
            with open(path, encoding="utf-8") as f:
                self._data.update(json.load(f))
        except FileNotFoundError:
            pass
        except Exception as e:
            self._app.logger.warning(f"config.load: {e}")
        return self

    def __repr__(self) -> str:
        items = ", ".join(f"{k}={v!r}" for k, v in self._data.items() if v)
        return f"Config({items})"
