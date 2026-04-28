"""
:mod:`hwpapi.core.app` — v2 slim :class:`App` facade (Phase 2).

Clean-cut rewrite from the v1 3,313-line monolith to the audited
**≤15 public members** surface. Every document-scoped operation now
lives behind :attr:`App.doc` (see :mod:`hwpapi.document`). Every
collection accessor moves to :mod:`hwpapi.collections` in Phase 3.
Every element accessor moves to :mod:`hwpapi.elements` in Phase 4.
The v1 surface is **deliberately removed**, not shimmed — callers
of removed members now hit :class:`AttributeError`.

Authoritative references
------------------------
- ``docs/design/app-member-audit.md`` — member disposition matrix.
- ``.omc/plans/hwpapi_v2_redesign.md`` Phase 2 (lines 240-256).
- ``.omc/prd.json`` stories P2-002 and P2-004.

Public surface (14 kept members + ``actions``)::

    __init__  __enter__  __exit__
    api  close  doc  engine  new  open  quit  reload  save  save_as  visible
    actions   # low-level escape hatch per audit §1.5
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import warnings

from hwpapi.low.actions import _Actions
from hwpapi.low.engine import Engine, Engines, Apps
from hwpapi.functions import check_dll, get_absolute_path
from hwpapi.logging import get_logger

__all__ = ["Engine", "Engines", "Apps", "App"]


class App:
    """
    High-level v2 façade for HWP (한글) automation.

    The slim v2 :class:`App` owns an :class:`~hwpapi.low.engine.Engine`
    and exposes only lifecycle + window-level concerns. All
    document-scoped operations live on :attr:`doc`.

    Parameters
    ----------
    new_app : bool, optional
        If ``True`` spawn a fresh engine; otherwise reuse the most
        recently opened one (or create one if none exist).
    is_visible : bool, optional
        Initial visibility of the HWP main window.
    dll_path : str | None, optional
        Path to ``FilePathCheckerModuleExample.dll``. Auto-discovered
        when omitted.
    engine : Engine | None, optional
        Pre-built engine to wrap. Rare — primarily for tests.

    Examples
    --------
    >>> with App() as app:
    ...     app.open("report.hwp")
    ...     app.doc.text                 # document text (→ Document)
    ...     app.save_as("report.pdf")
    """

    # ------------------------------------------------------------------
    # Construction & context-manager protocol
    # ------------------------------------------------------------------

    def __init__(
        self,
        new_app: bool = False,
        is_visible: bool = True,
        dll_path: Optional[str] = None,
        engine: Optional[Engine] = None,
    ) -> None:
        self._logger = get_logger("core")
        self._logger.debug(
            f"Initializing App (new_app={new_app}, is_visible={is_visible}, "
            f"dll_path={dll_path})"
        )

        self._load(new_app=new_app, dll_path=dll_path, engine=engine)

        # Low-level action dispatcher — escape hatch per audit §1.5.
        self.actions = _Actions(self)

        # Window visibility via the canonical property setter.
        self.visible = is_visible
        self._logger.info(f"App window visibility set to: {is_visible}")

        # Lazy Document facade cache (see `.doc`).
        self._doc_cache = None

        self._logger.info("App initialized (v2 slim facade)")

    @classmethod
    def new(cls, is_visible: bool = True) -> "App":
        """
        Create an :class:`App` backed by a brand-new HWP document.

        v2 rename of v1 ``App.new_document(is_tab=True)``. Replaces the
        document-level helper with a constructor-shaped factory so the
        returned object is the usual App handle.

        Parameters
        ----------
        is_visible : bool, optional
            Initial window visibility. Default ``True``.

        Returns
        -------
        App
            Fresh :class:`App` bound to a brand-new engine + document.
        """
        app = cls(new_app=True, is_visible=is_visible)
        try:
            # Ensure a blank document exists (``FileNew`` is idempotent).
            app.api.Run("FileNew")
        except Exception as exc:  # pragma: no cover — logged, non-fatal
            app._logger.warning(f"new(): FileNew failed: {exc}")
        return app

    def __enter__(self) -> "App":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        # Close the document cleanly; propagate any exception raised
        # inside the `with` block.
        try:
            self.quit()
        except Exception as close_exc:  # pragma: no cover
            self._logger.warning(f"__exit__: quit() raised: {close_exc}")

    # ------------------------------------------------------------------
    # Internal plumbing — not part of the public surface.
    # ------------------------------------------------------------------

    def _load(
        self,
        new_app: bool = False,
        engine: Optional[Engine] = None,
        dll_path: Optional[str] = None,
    ) -> None:
        """
        Bind the engine and register the FilePathChecker DLL.

        Extracted verbatim from v1 ``App._load`` so behaviour matches.
        """
        if new_app:
            engine = Engine()
            self._logger.debug("Created new engine for new_app")
        if not engine:
            engines = Engines()
            engine = engines[-1] if len(engines) > 0 else Engine()
            self._logger.debug("Selected engine from engines collection")

        self.engine = engine
        self._logger.info("Engine loaded successfully")

        if dll_path is None:
            from hwpapi.functions import get_hwp_dll_path

            dll_path = get_hwp_dll_path()

        if dll_path is not None:
            self._logger.info(f"Registering DLL: {dll_path}")
            check_dll(dll_path)
        else:
            self._logger.warning(
                "No DLL file found - some functionality may be limited"
            )

        try:
            self.api.RegisterModule(
                "FilePathCheckDLL", "FilePathCheckerModule"
            )
            self._logger.debug("Registered FilePathCheckDLL module")
        except Exception as exc:
            self._logger.warning(
                f"Failed to register FilePathCheckDLL module: {exc}"
            )
            warnings.warn(
                f"Failed to register FilePathCheckDLL module: {exc}"
            )

    # ------------------------------------------------------------------
    # Properties
    # ------------------------------------------------------------------

    @property
    def api(self):
        """Read-only COM handle — thin alias for ``self.engine.impl``."""
        return self.engine.impl

    @property
    def doc(self):
        """
        Cached :class:`hwpapi.document.Document` façade bound to this App.

        Same instance on every access. Primary v2 surface for per-document
        operations (text, fields, selection, collections …).
        """
        if self._doc_cache is None:
            # Lazy import breaks an otherwise-circular module cycle.
            from hwpapi.document import Document

            self._doc_cache = Document(self)
        return self._doc_cache

    @property
    def visible(self) -> bool:
        """HWP main-window visibility — read/write."""
        try:
            return bool(self.api.XHwpWindows.Item(0).Visible)
        except Exception:
            return False

    @visible.setter
    def visible(self, value: bool) -> None:
        self.api.XHwpWindows.Item(0).Visible = bool(value)

    # ------------------------------------------------------------------
    # Lifecycle methods
    # ------------------------------------------------------------------

    def open(self, path, format: Optional[str] = None, arg: str = "") -> str:
        """
        Open a document in HWP. Returns the absolute path opened.

        Parameters
        ----------
        path : str | Path
            File path (relative paths are resolved against cwd).
        format, arg : str, optional
            Forwarded to ``HwpObject.Open`` when provided (HWP's
            extended-open signature). Omit for the default behaviour.
        """
        self._logger.debug(f"open({path!r})")
        name = get_absolute_path(path)
        if format:
            self.api.Open(name, format, arg)
        else:
            self.api.Open(name)
        self._logger.info(f"open: {name}")
        return name

    def save(self, path=None, format: Optional[str] = None, arg: str = ""):
        """
        Save the active document.

        Without ``path`` performs an in-place save (``HwpObject.Save``).
        With ``path`` performs a ``SaveAs`` using the file-extension-
        derived format — ``.hwp``, ``.pdf``, ``.hwpx``, ``.hml``, ``.png``,
        ``.txt``, ``.docx``, ``.html`` / ``.htm`` are recognised.
        """
        self._logger.debug(f"save({path!r})")
        if not path:
            self.api.Save()
            current = self._current_filepath()
            self._logger.info(f"save (in-place): {current}")
            return current

        name = get_absolute_path(path)
        fmt = format or _format_from_suffix(Path(name).suffix)
        self.api.SaveAs(name, fmt)
        self._logger.info(f"save: {name} (format={fmt})")
        return name

    def save_as(self, path, format: Optional[str] = None, arg: str = ""):
        """
        v2 rename of v1 ``save_block``: save the active document
        selection/block to ``path``.

        Extension-sniffed format set matches v1: ``.hwp``, ``.pdf``,
        ``.hwpx`` (HWPML2X), ``.png``.
        """
        self._logger.debug(f"save_as({path!r})")
        name = get_absolute_path(path)
        fmt = format or _BLOCK_FORMAT_MAP.get(Path(name).suffix)
        action = self.actions.SaveBlockAction
        pset = action.pset
        pset.filename = name
        pset.Format = fmt
        action.run(pset)
        return name if Path(name).exists() else None

    def close(self, save: bool = False) -> bool:
        """
        Close the active document (``FileClose`` command).

        Parameters
        ----------
        save : bool, optional
            If ``True`` save before closing. HWP prompts by default; the
            caller can set this to force the save path.
        """
        self._logger.debug(f"close(save={save})")
        if save:
            try:
                self.api.Save()
            except Exception as exc:
                self._logger.warning(f"close(): pre-save failed: {exc}")
        return bool(self.api.Run("FileClose"))

    def quit(self) -> None:
        """Terminate the HWP engine (``FileQuit`` command)."""
        self._logger.debug("quit()")
        self.api.Run("FileQuit")

    def reload(self, new_app: bool = False, dll_path: Optional[str] = None) -> None:
        """
        Rebind this :class:`App` to a fresh/updated engine.

        Useful after DLL reinstall or when the HWP process has exited
        out from under the Python binding. Also invalidates the cached
        :attr:`doc`.
        """
        self._logger.debug(f"reload(new_app={new_app})")
        self._load(new_app=new_app, dll_path=dll_path)
        # Reset Document cache so `.doc` rebinds to the new engine.
        self._doc_cache = None

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _current_filepath(self) -> str:
        """Full path of the currently active document, or '' on error."""
        try:
            return str(self.api.XHwpDocuments.Active_XHwpDocument.FullName)
        except Exception:
            return ""

    def __repr__(self) -> str:
        try:
            path = self._current_filepath()
            tag = f" {path!r}" if path else ""
        except Exception:
            tag = ""
        return f"<hwpapi.App{tag}>"


# ----------------------------------------------------------------------
# Module-level helpers (not part of the App public surface)
# ----------------------------------------------------------------------

_SAVE_FORMAT_MAP = {
    ".hwp": "HWP",
    ".pdf": "PDF",
    ".hwpx": "HWPX",
    ".hml": "HWPML2X",
    ".png": "PNG",
    ".txt": "TEXT",
    ".docx": "MSWORD",
    ".html": "HTML+",
    ".htm": "HTML",
}

_BLOCK_FORMAT_MAP = {
    ".hwp": "HWP",
    ".pdf": "PDF",
    ".hwpx": "HWPML2X",
    ".png": "PNG",
}


def _format_from_suffix(suffix: str) -> Optional[str]:
    """Resolve an HWP save-format string from a file extension."""
    return _SAVE_FORMAT_MAP.get(suffix.lower())
