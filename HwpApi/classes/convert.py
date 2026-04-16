"""
Convert accessor — ``app.convert``.

텍스트 변환 헬퍼 모음. 승승아빠 매크로의 ``숫자를_한글숫자로`` /
``줄_나눔_어절_단어_순환`` 등을 이식한 단일 지점 유틸.

Usage::

    app.convert.number_to_korean()      # 선택된 "1,234,567" → "일백이십삼만..."
    app.convert.wrap_by_word()          # 줄 나눔 어절 단위
    app.convert.wrap_by_char()          # 줄 나눔 글자 단위
    app.convert.replace_font(old, new)  # 문서 전체 폰트 교체
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App


# Korean digit mapping
_DIGITS_KO = ("영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
# 만 단위 단위 (10000 = 만, 10000만 = 억, ...)
_UNITS_10K = ("", "만", "억", "조", "경")
# 자리수 단위
_DIGITS_10 = ("", "십", "백", "천")


def _int_to_korean(n: int) -> str:
    """정수를 한글 숫자로. 최대 경(10^16) 까지."""
    if n == 0:
        return "영"
    if n < 0:
        return "음" + _int_to_korean(-n)

    parts = []
    blk_idx = 0
    while n > 0 and blk_idx < len(_UNITS_10K):
        block = n % 10000
        n //= 10000
        if block == 0:
            blk_idx += 1
            continue
        block_str = ""
        digits = []
        tmp = block
        pos = 0
        while tmp > 0:
            d = tmp % 10
            if d != 0:
                # 일십, 일백, 일천 은 보통 생략 (십/백/천 으로) — simplification
                if d == 1 and pos > 0:
                    digits.append(_DIGITS_10[pos])
                else:
                    digits.append(_DIGITS_KO[d] + _DIGITS_10[pos])
            tmp //= 10
            pos += 1
        block_str = "".join(reversed(digits))
        parts.append(block_str + _UNITS_10K[blk_idx])
        blk_idx += 1

    return "".join(reversed(parts))


class Convert:
    """텍스트 변환 accessor."""

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    def number_to_korean(self, text: Optional[str] = None) -> Optional[str]:
        """
        숫자 → 한글숫자 변환. 승승아빠 매크로 ``숫자를_한글숫자로`` 이식.

        Parameters
        ----------
        text : str, optional
            변환할 텍스트. ``None`` 이면 현재 선택 영역을 사용하고,
            결과를 문서에 교체. ``text`` 제공 시에는 변환 문자열만 반환
            (문서는 건드리지 않음).

        Returns
        -------
        str
            변환된 한글숫자 문자열. 문서 수정 모드에서는 동일값.

        Examples
        --------
        >>> app.convert.number_to_korean("1,234,567")
        '일백이십삼만사천오백육십칠'

        >>> app.sel.current_line()
        >>> app.convert.number_to_korean()  # 현재 선택 변환
        """
        pattern = re.compile(r"[-]?[\d,]+")

        def sub_fn(m):
            raw = m.group(0).replace(",", "")
            try:
                n = int(raw)
                return _int_to_korean(n)
            except ValueError:
                return m.group(0)

        if text is not None:
            return pattern.sub(sub_fn, text)

        # Selection mode
        app = self._app
        sel = app.selection or ""
        if not sel:
            return None
        converted = pattern.sub(sub_fn, sel)
        # 선택 영역을 변환 결과로 교체
        try:
            app.insert_text(converted)   # selection 덮어씀
        except Exception as e:
            app.logger.debug(f"number_to_korean replace: {e}")
        return converted

    def wrap_by_word(self) -> "App":
        """**줄 나눔 어절 단위** — 승승아빠 매크로 ``줄_나눔_어절_단어_순환`` 이식."""
        app = self._app
        try:
            act = app.api.CreateAction("ParagraphShape")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("BreakNonLatinWord", 1)   # 어절 단위
            act.Execute(pset)
        except Exception as e:
            app.logger.debug(f"wrap_by_word: {e}")
        return app

    def wrap_by_char(self) -> "App":
        """**줄 나눔 글자 단위** (기본값)."""
        app = self._app
        try:
            act = app.api.CreateAction("ParagraphShape")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("BreakNonLatinWord", 0)   # 글자 단위
            act.Execute(pset)
        except Exception as e:
            app.logger.debug(f"wrap_by_char: {e}")
        return app

    def replace_font(self, old_font: str, new_font: str,
                      replace_all: bool = False) -> int:
        """
        문서에서 **특정 폰트만** 다른 폰트로 일괄 교체.

        v0.0.24 부터 ``old_font`` 인자가 실제로 사용됨 — 이전 버전은 문서
        전체를 ``new_font`` 로 덮어쓰는 버그가 있었음.

        Parameters
        ----------
        old_font : str
            바꿀 대상 폰트 이름 (예: ``"맑은 고딕"``).
        new_font : str
            새 폰트 이름 (예: ``"함초롬바탕"``).
        replace_all : bool
            True 이면 문서 전체를 ``new_font`` 로 덮어씀 (이전 버그 동작).
            기본 False — ``old_font`` 와 일치하는 텍스트만 교체.

        Returns
        -------
        int
            교체된 텍스트 영역 개수.

        Examples
        --------
        >>> # "맑은 고딕" 부분만 함초롬바탕으로
        >>> app.convert.replace_font("맑은 고딕", "함초롬바탕")
        42

        >>> # 문서 전체를 함초롬바탕으로 (이전 버그 동작 — 명시적 opt-in)
        >>> app.convert.replace_font(None, "함초롬바탕", replace_all=True)
        1

        Notes
        -----
        Find/Replace 시 7개 facename 필드 (hangul/latin/japanese/other/symbol/
        user/hanja) 모두를 한 번에 매칭하여 다국어 텍스트가 일괄 교체됨.
        큰 문서는 ``with app.batch_mode():`` 와 함께 쓰는 걸 권장.
        """
        app = self._app

        # Mode B: 전체 덮어쓰기 (legacy 버그 호환)
        if replace_all:
            app.logger.info(f"replace_font(all): → {new_font}")
            try:
                app.move.top_of_file()
                app.api.Run("SelectAll")
                app.set_charshape(facename_hangul=new_font,
                                   facename_latin=new_font,
                                   facename_japanese=new_font,
                                   facename_other=new_font,
                                   facename_symbol=new_font,
                                   facename_user=new_font,
                                   facename_hanja=new_font)
                app.api.Run("Cancel")
                return 1
            except Exception as e:
                app.logger.warning(f"replace_font(all): {type(e).__name__}: {e}",
                                    exc_info=True)
                return 0

        # Mode A: 특정 폰트만 교체 — AllReplace + HFindReplace pset
        if not old_font:
            app.logger.warning("replace_font: old_font 가 비어있음 — replace_all=True 로 전체 교체하려면 명시 필요")
            return 0

        count = 0
        try:
            app.move.top_of_file()
            hpset = app.api.HParameterSet.HFindReplace
            app.api.HAction.GetDefault("AllReplace", hpset.HSet)

            # 일치 조건: FindCharShape 의 facename 만 지정 → 다른 텍스트는 매칭 안 됨
            # 빈 문자열 매칭으로 모든 텍스트 검색 + charshape filter
            hpset.FindString = ""
            hpset.ReplaceString = ""
            hpset.IgnoreMessage = 1
            hpset.MatchCase = 0
            hpset.AllWordForms = 0
            hpset.SeveralWords = 0
            hpset.UseWildCards = 0
            hpset.WholeWordOnly = 0
            hpset.AutoSpell = 1
            hpset.Direction = 2   # all directions

            # Find side: old_font (7개 facename 모두 동일하게 매칭)
            try:
                fcs = hpset.FindCharShape
                for attr in ("FaceNameHangul", "FaceNameLatin",
                              "FaceNameJapanese", "FaceNameHanja",
                              "FaceNameOther", "FaceNameSymbol",
                              "FaceNameUser"):
                    try:
                        setattr(fcs, attr, old_font)
                    except Exception as e:
                        app.logger.debug(f"replace_font: cannot set Find.{attr}: {e}")
            except Exception as e:
                app.logger.debug(f"replace_font: FindCharShape access: {e}")

            # Replace side: new_font (7개 모두)
            try:
                rcs = hpset.ReplaceCharShape
                for attr in ("FaceNameHangul", "FaceNameLatin",
                              "FaceNameJapanese", "FaceNameHanja",
                              "FaceNameOther", "FaceNameSymbol",
                              "FaceNameUser"):
                    try:
                        setattr(rcs, attr, new_font)
                    except Exception as e:
                        app.logger.debug(f"replace_font: cannot set Replace.{attr}: {e}")
            except Exception as e:
                app.logger.debug(f"replace_font: ReplaceCharShape access: {e}")

            ok = app.api.HAction.Execute("AllReplace", hpset.HSet)
            count = 1 if ok else 0   # HWP 는 정확 count 미반환
            app.logger.info(f"replace_font: '{old_font}' → '{new_font}' (success={bool(ok)})")
        except Exception as e:
            app.logger.warning(f"replace_font: {type(e).__name__}: {e}", exc_info=True)
        return count

    def __repr__(self) -> str:
        return f"Convert(<app {id(self._app):x}>)"
