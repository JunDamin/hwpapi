"""
Debug accessor — ``app.debug``.

개발/디버깅 보조용. 현재 문서 상태, 커서 위치, 문자/문단 모양, 열린 문서 수
등을 한 번에 덤프해서 "왜 이상하지?" 를 빨리 파악할 수 있게 함.

Usage::

    app.debug.state()                   # 현재 상태 dict
    app.debug.print()                   # 상태 예쁘게 출력
    app.debug.timing(fn)                # 함수 호출 시간 측정
    with app.debug.trace(): ...         # 블록 내부 COM 호출 로그
"""
from __future__ import annotations

import time
from contextlib import contextmanager
from typing import TYPE_CHECKING, Any, Callable, Dict, Optional

if TYPE_CHECKING:
    from hwpapi.core.app import App


class Debug:
    """
    디버깅 헬퍼 accessor.

    모든 메소드는 실패 시 조용히 ``None`` 또는 빈 값 반환 — debug tool이
    디버그 대상 문제로 깨지면 본말전도.
    """

    __slots__ = ("_app",)

    def __init__(self, app: "App"):
        self._app = app

    def state(self) -> Dict[str, Any]:
        """
        현재 문서·커서·선택 상태를 dict 로 반환.

        Returns
        -------
        dict
            ``cursor``, ``page``, ``selection``, ``charshape_summary``,
            ``in_table``, ``documents_open``, ``visible``, ``version`` 키 포함.
        """
        app = self._app
        result = {}

        # Cursor position (para, pos)
        try:
            # GetPos 는 (ID, para, pos) 반환하는 경우가 있음
            pos = app.api.GetPos()
            if isinstance(pos, tuple) and len(pos) >= 3:
                result["cursor"] = {"doc_id": pos[0], "para": pos[1], "pos": pos[2]}
            else:
                result["cursor"] = str(pos)
        except Exception as e:
            result["cursor"] = f"<err: {e}>"

        # Page info
        try:
            result["page"] = {
                "current": app.current_page,
                "total": app.page_count,
            }
        except Exception as e:
            result["page"] = f"<err: {e}>"

        # Selection
        try:
            sel = app.selection or ""
            result["selection"] = {
                "text": sel[:60] + ("…" if len(sel) > 60 else ""),
                "length": len(sel),
            }
        except Exception as e:
            result["selection"] = f"<err: {e}>"

        # CharShape summary (key props only)
        try:
            cs = app.charshape
            result["charshape_summary"] = {
                "fontsize": getattr(cs, "fontsize", None) or getattr(cs, "height", None),
                "bold": getattr(cs, "bold", None),
                "italic": getattr(cs, "italic", None),
                "text_color": getattr(cs, "text_color", None),
                "shade_color": getattr(cs, "shade_color", None),
            }
        except Exception as e:
            result["charshape_summary"] = f"<err: {e}>"

        # In table?
        try:
            result["in_table"] = app.in_table()
        except Exception:
            result["in_table"] = None

        # Documents open
        try:
            result["documents_open"] = len(app.documents)
        except Exception:
            result["documents_open"] = None

        # Visibility & version
        try:
            result["visible"] = getattr(app, "visible", None) or \
                                 bool(app.api.XHwpWindows.Active_XHwpWindow.Visible)
        except Exception:
            result["visible"] = None

        try:
            result["version"] = app.version
        except Exception:
            result["version"] = None

        # File path
        try:
            result["filepath"] = app.get_filepath() or "(unsaved)"
        except Exception:
            result["filepath"] = None

        return result

    def print(self) -> None:
        """:meth:`state` 를 예쁘게 표 형태로 출력."""
        s = self.state()
        print("=" * 60)
        print(f" hwpapi debug state @ {time.strftime('%H:%M:%S')}")
        print("=" * 60)
        for k, v in s.items():
            if isinstance(v, dict):
                print(f"  {k}:")
                for kk, vv in v.items():
                    print(f"    {kk:20s} = {vv!r}")
            else:
                print(f"  {k:22s} = {v!r}")
        print("=" * 60)

    def timing(self, fn: Callable, *args, **kwargs) -> Dict[str, Any]:
        """
        함수 한 번 호출 시간을 측정. 반환값 + 걸린 시간(ms) 반환.

        Examples
        --------
        >>> r = app.debug.timing(app.insert_text, "테스트")
        >>> print(r['elapsed_ms'])
        12.34
        """
        t0 = time.perf_counter()
        exc = None
        result = None
        try:
            result = fn(*args, **kwargs)
        except Exception as e:
            exc = e
        elapsed_ms = (time.perf_counter() - t0) * 1000
        return {
            "result": result,
            "elapsed_ms": round(elapsed_ms, 3),
            "exception": exc,
            "success": exc is None,
        }

    @contextmanager
    def trace(self, verbose: bool = True):
        """
        **Context manager** — 블록 내부의 모든 COM ``Run`` 호출을 로그.

        Examples
        --------
        >>> with app.debug.trace():
        ...     app.insert_text("hello")
        [trace] Run("InsertText") → True (0.8ms)
        [trace] Run("...") → ...
        """
        app = self._app
        try:
            original_run = app.api.Run
        except Exception:
            yield
            return

        call_count = [0]

        def traced_run(name, *args, **kwargs):
            t0 = time.perf_counter()
            try:
                r = original_run(name, *args, **kwargs)
                elapsed = (time.perf_counter() - t0) * 1000
                if verbose:
                    print(f"[trace {call_count[0]:3d}] Run({name!r}) → {r} ({elapsed:.2f}ms)")
                call_count[0] += 1
                return r
            except Exception as e:
                elapsed = (time.perf_counter() - t0) * 1000
                if verbose:
                    print(f"[trace {call_count[0]:3d}] Run({name!r}) → EXC {e} ({elapsed:.2f}ms)")
                call_count[0] += 1
                raise

        try:
            app.api.Run = traced_run
        except Exception:
            # COM object may refuse attribute assignment — give up gracefully
            yield
            return

        try:
            yield
        finally:
            try:
                app.api.Run = original_run
            except Exception:
                pass
            if verbose:
                print(f"[trace] total: {call_count[0]} Run() calls")

    def __repr__(self) -> str:
        return f"Debug(<app {id(self._app):x}>)"
