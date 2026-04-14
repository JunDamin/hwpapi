"""
Pythonic accessor classes for navigating HWP documents.

These classes wrap the HWP COM API to provide clean, discoverable methods
for common operations:

- ``MoveAccessor``: Caret navigation (top/bottom of file, paragraph,
  line, word, next/prev, etc.)
- ``CellAccessor``: Table cell operations (content, formatting, merge)
- ``TableAccessor``: Table structure operations (create, resize, merge
  cells, insert rows/columns)
- ``PageAccessor``: Page-level operations (margins, breaks, numbering)

Attached to :class:`hwpapi.core.App` as ``app.move``, ``app.cell``,
``app.table``, ``app.page``.
"""
__all__ = ['MoveAccessor', 'CellAccessor', 'TableAccessor', 'PageAccessor', 'Character', 'CharShape', 'Paragraph', 'ParaShape',
           'PageShape']

import inspect
from dataclasses import asdict, dataclass, fields
from typing import Optional, Union

from hwpapi.constants import (
    chinese_fonts,
    english_fonts,
    japanese_fonts,
    korean_fonts,
    other_fonts,
    symbol_fonts,
    user_fonts,
)
from hwpapi.functions import (
    convert2int,
    convert_hwp_color_to_hex,
    convert_to_hwp_color,
    get_key,
    get_value,
    mili2unit,
    point2unit,
    unit2mili,
    unit2point,
)
import hwpapi.constants as const
from hwpapi.logging import get_logger




class MoveAccessor:
    """
    Wrap MovePos for easier use.
    """
    def __init__(self, app):
        self._app = app
        self.logger = get_logger('classes.MoveAccessor')

    def __call__(self, key=const.MoveId.ScanPos.value, para=None, pos=None):
        """
        지정된 키에 따라 캐럿 위치를 이동합니다.

        매개변수
        ----------
        key : MoveId, optional
            `MoveId` Enum에 정의된 이동 옵션. 기본값은 MoveId.ScanPos입니다.
        para : int, optional
            해당하는 경우 이동할 문단 번호. 기본값은 None입니다.
        pos : int, optional
            해당하는 경우 문단 내 위치. 기본값은 None입니다.
        
        반환값
        -------
        bool
            이동이 성공하면 True, 그렇지 않으면 False.

        사용 예시
        --------
        >>> app = App()
        >>> move(app, key=MoveId.TopOfFile)
        """
        move_id = key.value if isinstance(key, const.MoveId) else const.MoveId[key].value
        self.logger.debug(f"Moving cursor with moveID={move_id}, para={para}, pos={pos}")
        result = self._app.api.MovePos(moveID=move_id, Para=para, pos=pos)
        self.logger.debug(f"Move operation result: {result}")
        return result

    def main(self, para, pos):
        """루트 리스트의 특정 위치(para, pos로 위치 지정)"""
        return self._app.api.MovePos(moveID=const.MoveId.Main.value, Para=para, pos=pos)

    def current_list(self, para, pos):
        """현재 리스트의 특정 위치.(para pos로 위치 지정)"""
        return self._app.api.MovePos(moveID=const.MoveId.CurList.value, Para=para, pos=pos)

    def top_of_file(self):
        """문서의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.TopOfFile.value)

    def bottom_of_file(self):
        """문서의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.BottomOfFile.value)

    def top_of_list(self):
        """현재 리스트의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.TopOfList.value)

    def bottom_of_list(self):
        """현재 리스트의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.BottomOfList.value)

    def start_of_para(self):
        """현재 위치한 문단의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.StartOfPara.value)

    def end_of_para(self):
        """현재 위치한 문단의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.EndOfPara.value)

    def start_of_word(self):
        """현재 위치한 단어의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.StartOfWord.value)

    def end_of_word(self):
        """현재 위치한 단어의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.EndOfWord.value)

    def next_para(self):
        """다음 문단의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.NextPara.value)

    def prev_para(self):
        """앞 문단의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.PrevPara.value)

    def next_pos(self):
        """한 글자 뒤로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.NextPos.value)

    def prev_pos(self):
        """한 글자 앞으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.PrevPos.value)

    def next_pos_ex(self):
        """한 글자 뒤로 이동 (머리말/꼬리말 포함)"""
        return self._app.api.MovePos(moveID=const.MoveId.NextPosEx.value)

    def prev_pos_ex(self):
        """한 글자 앞으로 이동 (머리말/꼬리말 포함)"""
        return self._app.api.MovePos(moveID=const.MoveId.PrevPosEx.value)

    def next_char(self):
        """한 글자 뒤로 이동 (현재 리스트만을 대상으로 동작)"""
        return self._app.api.MovePos(moveID=const.MoveId.NextChar.value)

    def prev_char(self):
        """한 글자 앞으로 이동 (현재 리스트만을 대상으로 동작)"""
        return self._app.api.MovePos(moveID=const.MoveId.PrevChar.value)

    def next_word(self):
        """한 단어 뒤로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.NextWord.value)

    def prev_word(self):
        """한 단어 앞으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.PrevWord.value)

    def next_line(self):
        """한 줄 아래로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.NextLine.value)

    def prev_line(self):
        """한 줄 위로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.PrevLine.value)

    def start_of_line(self):
        """현재 위치한 줄의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.StartOfLine.value)

    def end_of_line(self):
        """현재 위치한 줄의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.EndOfLine.value)

    def parent_list(self):
        """한 레벨 상위로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.ParentList.value)

    def top_level_list(self):
        """탑레벨 리스트로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.TopLevelList.value)

    def root_list(self):
        """루트 리스트로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.RootList.value)

    def current_caret(self):
        """현재 캐럿이 위치한 곳으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.CurrentCaret.value)

    def left_of_cell(self):
        """현재 캐럿이 위치한 셀의 왼쪽으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.LeftOfCell.value)

    def right_of_cell(self):
        """현재 캐럿이 위치한 셀의 오른쪽으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.RightOfCell.value)

    def up_of_cell(self):
        """현재 캐럿이 위치한 셀의 위쪽으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.UpOfCell.value)

    def down_of_cell(self):
        """현재 캐럿이 위치한 셀의 아래쪽으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.DownOfCell.value)

    def start_of_cell(self):
        """현재 셀에서 행(row)의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.StartOfCell.value)

    def end_of_cell(self):
        """현재 셀에서 행(row)의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.EndOfCell.value)

    def top_of_cell(self):
        """현재 셀에서 열(column)의 시작으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.TopOfCell.value)

    def bottom_of_cell(self):
        """현재 셀에서 열(column)의 끝으로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.BottomOfCell.value)

    def scr_pos(self):
        """screen 좌표로 위치를 설정"""
        return self._app.api.MovePos(moveID=const.MoveId.ScrPos.value)

    def scan_pos(self):
        """GetText() 실행 후 위치로 이동"""
        return self._app.api.MovePos(moveID=const.MoveId.ScanPos.value)


class CellAccessor:
    """
    셀의 속성을 접근하고 수정할 수 있는 클래스입니다.

    이 클래스는 HWP 문서의 테이블 셀과 관련된 속성에 접근하고, 해당 속성을 읽거나 수정할 수 있도록 기능을 제공합니다. 
    셀의 너비와 높이를 포함한 다양한 속성을 다룰 수 있습니다.

    속성(Property)
    -------------
    width : float
        셀의 너비를 밀리미터 단위로 반환하거나 설정합니다.
    height : float
        셀의 높이를 밀리미터 단위로 반환하거나 설정합니다.

    메서드(Method)
    -------------
    _get_cell_property() -> dict
        셀과 관련된 속성들을 딕셔너리 형태로 반환합니다.

    사용 예시
    --------
    >>> app = App()  # HWP API 객체
    >>> cell = CellAccessor(app)
    >>> print(cell.width)  # 셀의 현재 너비 출력
    >>> cell.width = 50  # 셀의 너비를 50mm로 설정
    >>> print(cell.height)  # 셀의 현재 높이 출력
    >>> cell.height = 20  # 셀의 높이를 20mm로 설정
    """

    def __init__(self, app):
        """
        CellAccessor 클래스의 인스턴스를 초기화합니다.

        매개변수
        ----------
        app : App
            HWP API 객체로, 문서 및 테이블 속성에 접근하기 위해 필요합니다.
        """
        self._app = app
        self.logger = get_logger('classes.CellAccessor')


    def __call__(self):
        app = self._app
        action = app.actions.TablePropertyDialog
        return action.pset.shape_table_cell
    
    @property
    def width(self):
        """
        셀의 너비를 가져옵니다.

        반환값
        -------
        float
            셀의 너비 (밀리미터 단위).
        """
        
        app = self._app
        action = app.actions.TablePropertyDialog
        return action.pset.shape_table_cell.width

    @property
    def height(self):
        """
        셀의 높이를 가져옵니다.

        반환값
        -------
        float
            셀의 높이 (밀리미터 단위).
        """
        app = self._app
        action = app.actions.TablePropertyDialog
        return action.pset.shape_table_cell.height
        
    @width.setter
    def width(self, width):
        """
        셀의 너비를 설정합니다.

        매개변수
        ----------
        width : float
            셀의 새로운 너비 (밀리미터 단위).

        반환값
        -------
        bool
            너비 설정이 성공했는지 여부.
        """
        self.logger.debug(f"Setting cell width to: {width}mm")
        app = self._app
        action = app.actions.TablePropertyDialog
        action.pset.shape_table_cell.width = width
        action.run()
        result = app.cell.width == width
        self.logger.info(f"Cell width set to {width}mm, success: {result}")
        return result

    @height.setter
    def height(self, height):
        """
        셀의 높이를 설정합니다.

        매개변수
        ----------
        height : float
            셀의 새로운 높이 (밀리미터 단위).

        반환값
        -------
        bool
            높이 설정이 성공했는지 여부.
        """
        self.logger.debug(f"Setting cell height to: {height}mm")
        app = self._app
        action = app.actions.TablePropertyDialog
        action.pset.shape_table_cell.height = height
        action.run()
        result = app.cell.height == height
        self.logger.info(f"Cell height set to {height}mm, success: {result}")
        return result


class TableAccessor:
    """
    테이블 속성에 접근하고 조작할 수 있는 클래스입니다.

    이 클래스는 HWP 문서의 테이블과 관련된 속성에 접근하거나 수정할 수 있는 기능을 제공합니다.
    """

    miliunits = ["Width", "Height", "LayoutWidth", "LayoutHeight", "OutsideMarginLeft", "OutsideMarginRight", "OutsideMarginTop", "OutsideMarginBottom",]

    def __init__(self, app):
        """
        TableAccessor 클래스의 인스턴스를 초기화합니다.

        매개변수
        ----------
        app : App
            HWP API 객체로, 테이블 속성에 접근하거나 수정하기 위해 필요합니다.
        """
        self._app = app
        self.logger = get_logger('classes.TableAccessor')

    def _get_shape_properties(self):
        """
        테이블의 모양과 관련된 속성을 가져옵니다.

        반환값
        -------
        dict
            테이블의 속성 정보를 포함하는 딕셔너리입니다.
        """
        app = self._app
        property_names = (
            "TreatAsChar", "AffectsLine", "VertRelTo", "VertAlign", "VertOffset", "HorzRelTo",
            "HorzAlign", "HorzOffset", "FlowWithText", "AllowOverlap", "WidthRelTo", "Width",
            "HeightRelTo", "Height", "ProtectSize", "TextWrap", "TextFlow", "OutsideMarginLeft",
            "OutsideMarginRight", "OutsideMarginTop", "OutsideMarginBottom", "NumberingType",
            "LayoutWidth", "LayoutHeight", "Lock", "HoldAnchorObj", "PageNumber", "AdjustSelection",
            "AdjustTextBox", "AdjustPrevObjAttr"
        )
        action = app.actions.TablePropertyDialog
        pset = action._create_pset()
        return {name: unit2mili(pset.Item(name)) if name in TableAccessor.miliunits else pset.Item(name) for name in property_names }

    def _set_shape_property(self, name, value):
        """
        테이블의 특정 속성을 설정합니다.

        매개변수
        ----------
        name : str
            설정하려는 속성 이름.
        value : Any
            설정하려는 값.
        """
        app = self._app
        action = app.actions.TablePropertyDialog
        pset = action._create_pset()
        pset.SetItem(name, mili2unit(value) if name in TableAccessor.miliunits else value)
        action.run(pset)

    def __call__(self):
        """
        테이블 속성 정보를 호출합니다.

        반환값
        -------
        dict
            테이블의 속성 정보를 포함하는 딕셔너리입니다.
        """
        return self._get_shape_properties()

    # Property methods for all property names
    @property
    def treat_as_char(self):
        """테이블을 문자로 간주할지 여부를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("TreatAsChar")

    @treat_as_char.setter
    def treat_as_char(self, value):
        """테이블을 문자로 간주할지 여부를 설정합니다."""
        self._set_shape_property("TreatAsChar", value)

    @property
    def affects_line(self):
        """테이블이 줄에 영향을 미치는지 여부를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("AffectsLine")

    @affects_line.setter
    def affects_line(self, value):
        """테이블이 줄에 영향을 미치는지 여부를 설정합니다."""
        self._set_shape_property("AffectsLine", value)

    @property
    def vert_rel_to(self):
        """테이블의 수직 기준을 반환하거나 설정합니다."""
        return self._get_shape_properties().get("VertRelTo")

    @vert_rel_to.setter
    def vert_rel_to(self, value):
        """테이블의 수직 기준을 설정합니다."""
        self._set_shape_property("VertRelTo", value)

    @property
    def vert_align(self):
        """테이블의 수직 정렬 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("VertAlign")

    @vert_align.setter
    def vert_align(self, value):
        """테이블의 수직 정렬 상태를 설정합니다."""
        self._set_shape_property("VertAlign", value)

    @property
    def vert_offset(self):
        """테이블의 수직 오프셋을 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("VertOffset"))

    @vert_offset.setter
    def vert_offset(self, value):
        """테이블의 수직 오프셋을 설정합니다."""
        self._set_shape_property("VertOffset", mili2unit(value))

    @property
    def horz_rel_to(self):
        """테이블의 수평 기준을 반환하거나 설정합니다."""
        return self._get_shape_properties().get("HorzRelTo")

    @horz_rel_to.setter
    def horz_rel_to(self, value):
        """테이블의 수평 기준을 설정합니다."""
        self._set_shape_property("HorzRelTo", value)

    @property
    def horz_align(self):
        """테이블의 수평 정렬 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("HorzAlign")

    @horz_align.setter
    def horz_align(self, value):
        """테이블의 수평 정렬 상태를 설정합니다."""
        self._set_shape_property("HorzAlign", value)

    @property
    def horz_offset(self):
        """테이블의 수평 오프셋을 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("HorzOffset"))

    @horz_offset.setter
    def horz_offset(self, value):
        """테이블의 수평 오프셋을 설정합니다."""
        self._set_shape_property("HorzOffset", mili2unit(value))

    @property
    def flow_with_text(self):
        """테이블의 텍스트와 함께 이동 여부를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("FlowWithText")

    @flow_with_text.setter
    def flow_with_text(self, value):
        """테이블의 텍스트와 함께 이동 여부를 설정합니다."""
        self._set_shape_property("FlowWithText", value)

    @property
    def allow_overlap(self):
        """테이블의 겹치기 허용 여부를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("AllowOverlap")

    @allow_overlap.setter
    def allow_overlap(self, value):
        """테이블의 겹치기 허용 여부를 설정합니다."""
        self._set_shape_property("AllowOverlap", value)

    @property
    def width_rel_to(self):
        """테이블의 너비 기준을 반환하거나 설정합니다."""
        return self._get_shape_properties().get("WidthRelTo")

    @width_rel_to.setter
    def width_rel_to(self, value):
        """테이블의 너비 기준을 설정합니다."""
        self._set_shape_property("WidthRelTo", value)

    @property
    def width(self):
        """테이블의 너비를 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("Width"))

    @width.setter
    def width(self, value):
        """테이블의 너비를 설정합니다."""
        self._set_shape_property("Width", mili2unit(value))

    @property
    def height_rel_to(self):
        """테이블의 높이 기준을 반환하거나 설정합니다."""
        return self._get_shape_properties().get("HeightRelTo")

    @height_rel_to.setter
    def height_rel_to(self, value):
        """테이블의 높이 기준을 설정합니다."""
        self._set_shape_property("HeightRelTo", value)

    @property
    def height(self):
        """테이블의 높이를 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("Height"))

    @height.setter
    def height(self, value):
        """테이블의 높이를 설정합니다."""
        self._set_shape_property("Height", mili2unit(value))

    @property
    def protect_size(self):
        """테이블의 크기 보호 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("ProtectSize")

    @protect_size.setter
    def protect_size(self, value):
        """테이블의 크기 보호 상태를 설정합니다."""
        self._set_shape_property("ProtectSize", value)

    @property
    def text_wrap(self):
        """테이블의 텍스트 줄바꿈 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("TextWrap")

    @text_wrap.setter
    def text_wrap(self, value):
        """테이블의 텍스트 줄바꿈 상태를 설정합니다."""
        self._set_shape_property("TextWrap", value)

    @property
    def text_flow(self):
        """테이블의 텍스트 흐름 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("TextFlow")

    @text_flow.setter
    def text_flow(self, value):
        """테이블의 텍스트 흐름 상태를 설정합니다."""
        self._set_shape_property("TextFlow", value)

    @property
    def outside_margin_left(self):
        """테이블의 왼쪽 외부 여백을 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("OutsideMarginLeft"))

    @outside_margin_left.setter
    def outside_margin_left(self, value):
        """테이블의 왼쪽 외부 여백을 설정합니다."""
        self._set_shape_property("OutsideMarginLeft", mili2unit(value))

    @property
    def outside_margin_right(self):
        """테이블의 오른쪽 외부 여백을 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("OutsideMarginRight"))

    @outside_margin_right.setter
    def outside_margin_right(self, value):
        """테이블의 오른쪽 외부 여백을 설정합니다."""
        self._set_shape_property("OutsideMarginRight", mili2unit(value))

    @property
    def outside_margin_top(self):
        """테이블의 위쪽 외부 여백을 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("OutsideMarginTop"))

    @outside_margin_top.setter
    def outside_margin_top(self, value):
        """테이블의 위쪽 외부 여백을 설정합니다."""
        self._set_shape_property("OutsideMarginTop", mili2unit(value))

    @property
    def outside_margin_bottom(self):
        """테이블의 아래쪽 외부 여백을 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("OutsideMarginBottom"))

    @outside_margin_bottom.setter
    def outside_margin_bottom(self, value):
        """테이블의 아래쪽 외부 여백을 설정합니다."""
        self._set_shape_property("OutsideMarginBottom", mili2unit(value))

    @property
    def numbering_type(self):
        """테이블의 번호 매기기 유형을 반환하거나 설정합니다."""
        return self._get_shape_properties().get("NumberingType")

    @numbering_type.setter
    def numbering_type(self, value):
        """테이블의 번호 매기기 유형을 설정합니다."""
        self._set_shape_property("NumberingType", value)

    @property
    def layout_width(self):
        """테이블의 레이아웃 너비를 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("LayoutWidth"))

    @layout_width.setter
    def layout_width(self, value):
        """테이블의 레이아웃 너비를 설정합니다."""
        self._set_shape_property("LayoutWidth", mili2unit(value))

    @property
    def layout_height(self):
        """테이블의 레이아웃 높이를 반환하거나 설정합니다."""
        return unit2mili(self._get_shape_properties().get("LayoutHeight"))

    @layout_height.setter
    def layout_height(self, value):
        """테이블의 레이아웃 높이를 설정합니다."""
        self._set_shape_property("LayoutHeight", mili2unit(value))

    @property
    def lock(self):
        """테이블의 잠금 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("Lock")

    @lock.setter
    def lock(self, value):
        """테이블의 잠금 상태를 설정합니다."""
        self._set_shape_property("Lock", value)

    @property
    def hold_anchor_obj(self):
        """테이블의 고정 앵커 개체 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("HoldAnchorObj")

    @hold_anchor_obj.setter
    def hold_anchor_obj(self, value):
        """테이블의 고정 앵커 개체 상태를 설정합니다."""
        self._set_shape_property("HoldAnchorObj", value)

    @property
    def page_number(self):
        """테이블의 페이지 번호를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("PageNumber")

    @page_number.setter
    def page_number(self, value):
        """테이블의 페이지 번호를 설정합니다."""
        self._set_shape_property("PageNumber", value)

    @property
    def adjust_selection(self):
        """테이블 선택 조정 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("AdjustSelection")

    @adjust_selection.setter
    def adjust_selection(self, value):
        """테이블 선택 조정 상태를 설정합니다."""
        self._set_shape_property("AdjustSelection", value)

    @property
    def adjust_text_box(self):
        """테이블 텍스트 상자 조정 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("AdjustTextBox")

    @adjust_text_box.setter
    def adjust_text_box(self, value):
        """테이블 텍스트 상자 조정 상태를 설정합니다."""
        self._set_shape_property("AdjustTextBox", value)

    @property
    def adjust_prev_obj_attr(self):
        """이전 개체 속성 조정 상태를 반환하거나 설정합니다."""
        return self._get_shape_properties().get("AdjustPrevObjAttr")

    @adjust_prev_obj_attr.setter
    def adjust_prev_obj_attr(self, value):
        """이전 개체 속성 조정 상태를 설정합니다."""
        self._set_shape_property("AdjustPrevObjAttr", value)


class PageAccessor:
    
    def __init__(self, app):
        self._app = app
        self.logger = get_logger('classes.PageAccessor')

    def _get_miliproperty(self, name):
        action = self._app.actions.PageSetup
        pset = action._create_pset()
        value = pset.Item("PageDef").Item(name)
        return unit2mili(value)

    def _set_miliproperty(self, name, value):
        action = self._app.actions.PageSetup
        pset = action._create_pset()
        pset.Item("PageDef").SetItem(name, mili2unit(value))
        return action.run(pset)

    def __call__(self):
        return self._get_properties

    @property
    def inner_width(self):
        action = self._app.actions.PageSetup
        pset = action._create_pset()
        paper_width = pset.Item("PageDef").Item("PaperWidth")
        left_margin = pset.Item("PageDef").Item("LeftMargin")
        right_margin = pset.Item("PageDef").Item("RightMargin")
        return unit2mili(paper_width - (left_margin + right_margin))

    @property
    def inner_height(self):
        action = self._app.actions.PageSetup
        pset = action._create_pset()
        paper_height = pset.Item("PageDef").Item("PaperHeight")
        top_margin = pset.Item("PageDef").Item("TopMargin")
        bottom_margin = pset.Item("PageDef").Item("BottomMargin")
        return unit2mili(paper_height - (top_margin + bottom_margin))
    
    
    @property
    def paper_height(self):
        return self._get_miliproperty("PaperHeight")
    
    @paper_height.setter
    def papaer_height(self, value: float):
        return self._set_miliproperty("PaperHeight", value)

    @property
    def paper_width(self):
        return self._get_miliproperty("PaperWidth")
    
    @paper_width.setter
    def paper_width(self, value):
        return self._set_miliproperty("PaperWidth", value)
    
    @property
    def top_margin(self):
        return self._get_miliproperty("TopMargin")
    
    @top_margin.setter
    def top_margin(self, value):
        return self._set_miliproperty("TopMargin", value)
    
    @property
    def bottom_margin(self):
        return self._get_miliproperty("BottomMargin")
    
    @bottom_margin.setter
    def bottom_margin(self, value):
        return self._set_miliproperty("BottomMargin", value)
    
    @property
    def left_margin(self):
        return self._get_miliproperty("LeftMargin")

    @left_margin.setter
    def left_margin(self, value):
        return self._set_miliproperty("LeftMargin", value)

    @property
    def right_margin(self):
        return self._get_miliproperty("RightMargin")
    
    @right_margin.setter
    def right_margin(self, value):
        return self._set_miliproperty("RightMargin", value)
        
    @property
    def header(self):
        return self._get_miliproperty("HeaderLen")
    
    @header.setter
    def header(self, value):
        return self._set_miliproperty("HeaderLen", value)
        
        
    @property
    def footer(self):
        return self._get_miliproperty("FooterLen")
    
    @header.setter
    def footer(self, value):
        return self._set_miliproperty("FooterLen", value)
        
        
    @property
    def gutter(self):
        return self._get_miliproperty("GutterLen")
    
    @gutter.setter
    def gutter(self, value):
        return self._set_miliproperty("GutterLen", value)
        
    @property
    def _get_properties(self):
        action = self._app.actions.PageSetup
        pset = action._create_pset()
        hwpunit_property_names, property_names = ("PaperWidth", "PaperHeight", "LeftMargin", "RightMargin", "TopMargin", "BottomMargin", "HeaderLen", "FooterLen", "GutterLen"), ("Landscape", "GutterType", "ApplyTo", "ApplyClass")
        properties = {name: unit2mili(pset.Item("PageDef").Item(name)) for name in hwpunit_property_names}
        properties.update({name: pset.Item("PageDef").Item(name) for name in property_names})
        return properties

