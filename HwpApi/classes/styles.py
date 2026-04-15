"""
Styles accessor — ``app.styles``.

HWP 문서의 **문단 스타일(ParaStyle)** 관리. 리스트 조회, 적용, 추가/삭제,
파일 내보내기·가져오기.

Extracted as a Phase D deliverable (PYHWPX_COMPARISON.md roadmap).
"""
from __future__ import annotations
from typing import Union, List, Optional

__all__ = ["StylesAccessor", "Style"]


class Style:
    """문단 스타일 하나를 감싼 값 객체.

    Attributes
    ----------
    index : int
        스타일 인덱스 (0-based).
    name : str
        스타일 이름 (예: '바탕글', '제목 1').
    """

    __slots__ = ("index", "name", "_app")

    def __init__(self, index: int, name: str, app=None):
        self.index = int(index)
        self.name = str(name)
        self._app = app

    def apply(self) -> bool:
        """현재 커서 위치(또는 선택 영역) 문단에 이 스타일 적용."""
        if self._app is None:
            return False
        return bool(self._app.styles.apply(self.index))

    def delete(self, replace_with: Union[int, str] = 0) -> bool:
        """이 스타일을 삭제하고 사용 부분은 ``replace_with`` 스타일로 대체."""
        if self._app is None:
            return False
        return bool(self._app.styles.delete(self.index, replace_with))

    def __repr__(self):
        return f"Style(index={self.index}, name={self.name!r})"

    def __eq__(self, other):
        if isinstance(other, Style):
            return self.index == other.index and self.name == other.name
        if isinstance(other, str):
            return self.name == other
        if isinstance(other, int):
            return self.index == other
        return NotImplemented

    def __hash__(self):
        return hash((self.index, self.name))


class StylesAccessor:
    """
    ``app.styles`` — 문단 스타일 관리 접근자.

    Examples
    --------
    >>> len(app.styles)                    # 스타일 개수
    >>> for s in app.styles:               # 순회
    ...     print(s.index, s.name)
    >>> app.styles.names                   # 이름 리스트
    ['바탕글', '본문', '개요 1', ...]

    >>> "바탕글" in app.styles              # 존재 확인
    True
    >>> s = app.styles["제목 1"]            # 이름으로 조회
    >>> s.apply()                          # 커서 문단에 적용

    >>> app.styles.apply("제목 1")          # 단축 — 이름으로 바로 적용
    >>> app.styles.current                 # 현재 문단의 스타일

    >>> app.styles.export("my_styles.sty")
    >>> app.styles.import_from("my_styles.sty")
    >>> app.styles.remove_unused()         # 사용 안 된 스타일 제거
    """

    def __init__(self, app):
        self._app = app

    # ── 조회 ────────────────────────────────────────────────

    @property
    def names(self) -> List[str]:
        """모든 스타일 이름 리스트 (index 순)."""
        return [s.name for s in self._iter_styles()]

    def _iter_styles(self):
        """HWPML 내보내기 + XML 파싱 — pyhwpx 패턴."""
        # HStyle pset 의 Apply 필드 값들을 iterate 하는 직접 방법이 없으므로
        # HWPML 로 블록 내보내기 → XML 에서 <STYLE> 태그 파싱.
        import tempfile, os, re
        try:
            with tempfile.NamedTemporaryFile(
                suffix=".xml", delete=False
            ) as tmp:
                xml_path = tmp.name

            # Ensure selection is minimal (1 char or nothing)
            try:
                prev_pos = self._app.api.GetPos()
            except Exception:
                prev_pos = None

            saved = False
            try:
                # HWP needs a SELECTION to use SaveBlockAs; select 1 char.
                try:
                    self._app.api.Run("MoveSelRight")
                    saved = self._app.api.SaveBlockAction(
                        xml_path, "HWPML2X"
                    ) if hasattr(self._app.api, "SaveBlockAction") else False
                    if not saved:
                        saved = self._app.api.SaveBlockAs(
                            xml_path, "HWPML2X"
                        )
                except Exception:
                    pass
            finally:
                try:
                    self._app.api.Run("Cancel")
                    if prev_pos is not None:
                        self._app.api.SetPos(*prev_pos)
                except Exception:
                    pass

            if not (saved and os.path.isfile(xml_path)):
                return
            with open(xml_path, encoding="utf-8") as f:
                xml = f.read()
            # <STYLE ... Name="..." ...>
            for i, m in enumerate(
                re.finditer(r'<STYLE[^>]+Name="([^"]+)"', xml)
            ):
                yield Style(i, m.group(1), app=self._app)
            try:
                os.remove(xml_path)
            except Exception:
                pass
        except Exception:
            return

    def __iter__(self):
        return iter(list(self._iter_styles()))

    def __len__(self) -> int:
        return len(self.names)

    def __contains__(self, name_or_index) -> bool:
        if isinstance(name_or_index, int):
            return 0 <= name_or_index < len(self)
        return str(name_or_index) in self.names

    def __getitem__(self, key) -> Optional[Style]:
        """이름 또는 인덱스로 :class:`Style` 반환."""
        styles = list(self._iter_styles())
        if isinstance(key, int):
            if -len(styles) <= key < len(styles):
                return styles[key]
            return None
        if isinstance(key, str):
            for s in styles:
                if s.name == key:
                    return s
            return None
        raise TypeError(f"indices must be int or str, not {type(key).__name__}")

    @property
    def current(self) -> Optional[Style]:
        """현재 커서 문단의 스타일."""
        try:
            pset = self._app.api.HParameterSet.HStyle
            self._app.api.HAction.GetDefault("Style", pset.HSet)
            idx = int(pset.Apply)
        except Exception:
            return None
        styles = list(self._iter_styles())
        if 0 <= idx < len(styles):
            return styles[idx]
        return None

    # ── 편집 ────────────────────────────────────────────────

    def _resolve_index(self, name_or_index) -> Optional[int]:
        """이름/인덱스 → 정수 인덱스."""
        if isinstance(name_or_index, Style):
            return name_or_index.index
        if isinstance(name_or_index, int):
            return int(name_or_index)
        if isinstance(name_or_index, str):
            names = self.names
            if name_or_index in names:
                return names.index(name_or_index)
        return None

    def apply(self, name_or_index) -> bool:
        """
        현재 커서 문단에 스타일 적용.

        Parameters
        ----------
        name_or_index : str | int | Style
            스타일 이름 또는 인덱스.
        """
        idx = self._resolve_index(name_or_index)
        if idx is None:
            return False
        try:
            pset = self._app.api.HParameterSet.HStyle
            self._app.api.HAction.GetDefault("Style", pset.HSet)
            pset.Apply = int(idx)
            return bool(self._app.api.HAction.Execute("Style", pset.HSet))
        except Exception:
            return False

    def delete(self, name_or_index, replace_with: Union[int, str] = 0) -> bool:
        """
        스타일 삭제 — 기존에 해당 스타일이 적용된 부분은 ``replace_with`` 로 변경.
        """
        idx = self._resolve_index(name_or_index)
        alt = self._resolve_index(replace_with)
        if idx is None or alt is None:
            return False
        try:
            pset = self._app.api.HParameterSet.HStyleDelete
            self._app.api.HAction.GetDefault("StyleDelete", pset.HSet)
            pset.HSet.SetItem("Target", int(idx))
            pset.HSet.SetItem("Alternation", int(alt))
            return bool(
                self._app.api.HAction.Execute("StyleDelete", pset.HSet)
            )
        except Exception:
            return False

    def export(self, path: str) -> bool:
        """
        현재 문서의 스타일을 ``.sty`` 파일로 export.

        Examples
        --------
        >>> app.styles.export("house_style.sty")
        """
        try:
            return bool(self._app.api.ExportStyle(str(path)))
        except Exception:
            return False

    def import_from(self, path: str) -> bool:
        """``.sty`` 파일에서 스타일 가져오기."""
        try:
            return bool(self._app.api.ImportStyle(str(path)))
        except Exception:
            return False

    def remove_unused(self, replace_with: Union[int, str] = 0) -> int:
        """
        문서에서 **사용되지 않은** 스타일 모두 삭제. 삭제된 개수 반환.
        """
        all_names = set(self.names)
        # Detect 'used' styles by iterating paragraphs and reading HStyle.Apply
        # This is expensive; for now use a best-effort approach — read the XML
        # and look for Apply="..." attribute in <PARA> tags.
        used = set()
        try:
            import tempfile, os, re
            with tempfile.NamedTemporaryFile(
                suffix=".xml", delete=False
            ) as tmp:
                xml_path = tmp.name
            self._app.api.SelectAll()
            try:
                self._app.api.SaveBlockAs(xml_path, "HWPML2X")
            finally:
                self._app.api.Run("Cancel")
            if os.path.isfile(xml_path):
                with open(xml_path, encoding="utf-8") as f:
                    xml = f.read()
                names = self.names
                # <P Style="N"> or similar
                for m in re.finditer(r'Style="(\d+)"', xml):
                    idx = int(m.group(1))
                    if 0 <= idx < len(names):
                        used.add(names[idx])
                os.remove(xml_path)
        except Exception:
            pass

        unused = all_names - used
        count = 0
        # Delete in reverse-index order so earlier indices stay valid
        for name in sorted(unused, key=lambda n: -self.names.index(n)
                            if n in self.names else 0):
            if self.delete(name, replace_with):
                count += 1
        return count

    def __repr__(self):
        return f"<StylesAccessor count={len(self)}>"
