# AUTOGENERATED! DO NOT EDIT! File to edit: ../nbs/02_api/03_dataclasses.ipynb.

# %% auto 0
__all__ = ['Character', 'CharShape', 'Paragraph', 'ParaShape']

# %% ../nbs/02_api/03_dataclasses.ipynb 3
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

# %% ../nbs/02_api/03_dataclasses.ipynb 4
@dataclass
class Character:
    Bold: Optional[int] = None
    DiacSymMark: Optional[int] = None
    Emboss: Optional[int] = None
    Engrave: Optional[int] = None
    FaceNameHangul: Optional[str] = None
    FaceNameHanja: Optional[str] = None
    FaceNameJapanese: Optional[str] = None
    FaceNameLatin: Optional[str] = None
    FaceNameOther: Optional[str] = None
    FaceNameSymbol: Optional[str] = None
    FaceNameUser: Optional[str] = None
    FontTypeHangul: Optional[int] = None
    FontTypeHanja: Optional[int] = None
    FontTypeJapanese: Optional[int] = None
    FontTypeLatin: Optional[int] = None
    FontTypeOther: Optional[int] = None
    FontTypeSymbol: Optional[int] = None
    FontTypeUser: Optional[int] = None
    Height: Optional[int] = None
    Italic: Optional[int] = None
    OffsetHangul: Optional[int] = None
    OffsetHanja: Optional[int] = None
    OffsetJapanese: Optional[int] = None
    OffsetLatin: Optional[int] = None
    OffsetOther: Optional[int] = None
    OffsetSymbol: Optional[int] = None
    OffsetUser: Optional[int] = None
    OutLineType: Optional[int] = None
    RatioHangul: Optional[int] = None
    RatioHanja: Optional[int] = None
    RatioJapanese: Optional[int] = None
    RatioLatin: Optional[int] = None
    RatioOther: Optional[int] = None
    RatioSymbol: Optional[int] = None
    RatioUser: Optional[int] = None
    ShadeColor: Optional[int] = None
    ShadowColor: Optional[int] = None
    ShadowOffsetX: Optional[int] = None
    ShadowOffsetY: Optional[int] = None
    ShadowType: Optional[int] = None
    SizeHangul: Optional[int] = None
    SizeHanja: Optional[int] = None
    SizeJapanese: Optional[int] = None
    SizeLatin: Optional[int] = None
    SizeOther: Optional[int] = None
    SizeSymbol: Optional[int] = None
    SizeUser: Optional[int] = None
    SmallCaps: Optional[int] = None
    SpacingHangul: Optional[int] = None
    SpacingHanja: Optional[int] = None
    SpacingJapanese: Optional[int] = None
    SpacingLatin: Optional[int] = None
    SpacingOther: Optional[int] = None
    SpacingSymbol: Optional[int] = None
    SpacingUser: Optional[int] = None
    StrikeOutColor: Optional[int] = None
    StrikeOutShape: Optional[int] = None
    StrikeOutType: Optional[int] = None
    SubScript: Optional[int] = None
    SuperScript: Optional[int] = None
    TextColor: Optional[int] = None
    UnderlineColor: Optional[int] = None
    UnderlineShape: Optional[int] = None
    UnderlineType: Optional[int] = None
    UseFontSpace: Optional[int] = None
    UseKerning: Optional[int] = None

# %% ../nbs/02_api/03_dataclasses.ipynb 5
import inspect


class CharShape:
    """
    CharShape 클래스는 문자 모양을 다룹니다. CharShape는 폰트, 색상, 크기 및 스타일과 같이 문자에 적용된 스타일링 속성의 전체 세트를 나타냅니다.

    메서드
    -------
    __init__(char_pset=None, **kwargs):
        CharShape 클래스의 새 인스턴스를 초기화합니다. 키워드 인수를 사용하여 다양한 속성의 초기 값을 설정할 수 있습니다.

    __str__():
        CharShape 인스턴스의 모든 속성과 그들의 현재 값이 나열된 문자열을 반환합니다.

    __repr__():
        __str__() 메서드와 같은 역할을 합니다.

    _set_font(attr_name, value, font_dict):
        문자 모양의 폰트를 설정하기 위한 보조 메서드입니다.

    _get_font(attr_name):
        문자 모양의 폰트를 가져오기 위한 보조 메서드입니다.

    @property 와 @<property_name>.setter 쌍들:
        문자 모양의 다양한 속성에 대한 getter 및 setter 메서드를 정의합니다.

    todict():
        CharShape 인스턴스를 딕셔너리으로 변환합니다.

    fromdict(values: dict):
        값을 담고 있는 딕셔너리를 받아 CharShape 값을 채웁니다.

    get_from_pset(pset):
        주어진 아래아 한글 파라미터셋에서 CharShape 값을 채웁니다.
    """
    
    attributes = [
            "hangul_font",
            "latin_font",
            "text_color",
            "fontsize",
            "bold",
            "italic",
            "super_script",
            "sub_script",
            "offset",
            "spacing",
            "ratio",
            "shade_color",
            "shadow_color",
            "shadow_offset_x",
            "shadow_offset_y",
            "strike_out_type",
            "strike_out_color",
            "underline_type",
            "underline_shape",
            "underline_color",
            "out_line_type",
        ]

    def __init__(self, char_pset=None, **kwargs):
        self.p = Character()

        if char_pset:
            self.get_from_pset(char_pset)

        for key, value in kwargs.items():
            setattr(self, key, value)

        self.__repr__ = self.__str__

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        
        values = [getattr(self, attr) for attr in CharShape.attributes]
        return f"<CharShape: {', '.join([f'{attr}={value}' for attr, value in zip(CharShape.attributes, values)])}>"

    def _set_font(self, attr_name, value, font_dict):
        setattr(self.p, attr_name, value)
        setattr(self.p, f"FontType{attr_name[8:]}", 2 if value in font_dict else 1)

    def _get_font(self, attr_name):
        return getattr(self.p, attr_name)

    # font
    #
    # hangul
    @property
    def hangul_font(self):
        return self._get_font("FaceNameHangul")

    @hangul_font.setter
    def hangul_font(self, value):
        self._set_font("FaceNameHangul", value, korean_fonts)

    # hanja
    @property
    def hanja_font(self):
        return self._get_font("FaceNameHanja")

    @hanja_font.setter
    def hanja_font(self, value):
        self._set_font("FaceNameHanja", value, chinese_fonts)

    # japanese
    @property
    def japanese_font(self):
        return self._get_font("FaceNameJapanese")

    @japanese_font.setter
    def japanese_font(self, value):
        self._set_font("FaceNameJapanese", value, japanese_fonts)

    # latin
    @property
    def latin_font(self):
        return self._get_font("FaceNameLatin")

    @latin_font.setter
    def latin_font(self, value):
        self._set_font("FaceNameLatin", value, english_fonts)

    # other
    @property
    def other_font(self):
        return self._get_font("FaceNameOther")

    @other_font.setter
    def other_font(self, value):
        self._set_font("FaceNameOther", value, other_fonts)

    # symbol
    @property
    def symbol_font(self):
        return self._get_font("FaceNameSymbol")

    @symbol_font.setter
    def symbol_font(self, value):
        self._set_font("FaceNameSymbol", value, symbol_fonts)

    # user
    @property
    def user_font(self):
        return self._get_font("FaceNameUser")

    @user_font.setter
    def user_font(self, value):
        self._set_font("FaceNameUser", value, user_fonts)

    @property
    def font(self):
        fonts = [
            ("한글", self.hangul_font),
            ("영어", self.latin_font),
            ("한자", self.hanja_font),
            ("일어", self.japanese_font),
            ("심볼", self.symbol_font),
            ("유저", self.user_font),
            ("기타", self.other_font),
        ]
        return f"<{', '.join([f'{key}: {value}' for key, value in fonts])}>"

    @font.setter
    def font(self, value):
        self.hangul_font = value
        self.latin_font = value
        self.hanja_font = value
        self.japanese_font = value
        self.symbol_font = value
        self.user_font = value
        self.other_font = value

    # text color
    @property
    def text_color(self):
        return convert_hwp_color_to_hex(self.p.TextColor)

    @text_color.setter
    def text_color(self, color: Union[int, str, tuple]):
        value = convert_to_hwp_color(color)
        self.p.TextColor = value

    # font size
    @property
    def fontsize(self):
        return unit2point(self.p.Height)

    @fontsize.setter
    def fontsize(self, value):
        self.p.Height = point2unit(value)

    # style
    #
    # bold
    @property
    def bold(self):
        return self.p.Bold

    @bold.setter
    def bold(self, value):
        self.p.Bold = value

    # italic
    @property
    def italic(self):
        return self.p.Italic

    @italic.setter
    def italic(self):
        self.p.Italic = value

    # super_script
    @property
    def super_script(self):
        return self.p.SuperScript

    @super_script.setter
    def super_script(self, value):
        self.p.SuperScript = value

    # subscript
    @property
    def sub_script(self):
        return self.p.SubScript

    @sub_script.setter
    def sub_script(self, value):
        self.p.SubrScript = value

    # offset
    @property
    def offset(self):
        return self.p.OffsetHangul

    @offset.setter
    def offset(self, value):
        self.p.OffsetHangul = value
        self.p.OffsetHanja = value
        self.p.OffsetJapanese = value
        self.p.OffsetLatin = value
        self.p.OffsetOther = value
        self.p.OffsetSymbol = value
        self.p.OffsetUser = value

    # ratio
    @property
    def ratio(self):
        return self.p.RatioHangul

    @ratio.setter
    def ratio(self, value):
        self.p.RatioHangul = value
        self.p.RatioHanja = value
        self.p.RatioJapanese = value
        self.p.RatioLatin = value
        self.p.RatioOther = value
        self.p.RatioSymbol = value
        self.p.RatioUser = value

    # spacing
    @property
    def spacing(self):
        return self.p.SpacingHangul

    @spacing.setter
    def spacing(self, value):
        self.p.SpacingHangul = value
        self.p.SpacingHanja = value
        self.p.SpacingJapanese = value
        self.p.SpacingLatin = value
        self.p.SpacingOther = value
        self.p.SpacingSymbol = value
        self.p.SpacingUser = value

    # outline
    @property
    def out_line_type(self):
        return self.p.OutLineType

    @out_line_type.setter
    def out_line_type(self, value):
        self.p.OutLineType = value

    # shade
    #
    #
    @property
    def shade_color(self):
        return convert_hwp_color_to_hex(self.p.ShadeColor)

    @shade_color.setter
    def shade_color(self, color: Union[int, str, tuple]):
        value = convert_to_hwp_color(color)
        self.p.ShadeColor = value

    # shadow
    #
    #
    @property
    def shadow_color(self):
        return convert_hwp_color_to_hex(self.p.ShadowColor)

    @shadow_color.setter
    def shadow_color(self, color: Union[int, str, tuple]):
        value = convert_to_hwp_color(color)
        self.p.ShadowColor = value

    @property
    def shadow_offset_x(self):
        return self.p.ShadowOffsetX

    @shadow_offset_x.setter
    def shadow_offset_x(self, value):
        self.p.ShadowOffsetX = value

    @property
    def shadow_offset_y(self):
        return self.p.ShadowOffsetY

    @shadow_offset_y.setter
    def shadow_offset_y(self, value):
        self.p.ShadowOffsetY = value

    @property
    def shadow_type(self):
        return self.p.ShadowType

    @shadow_type.setter
    def shadow_type(self, value):
        self.p.ShadowType = value

    # StrikeOut
    #
    #
    @property
    def strike_out_color(self):
        return convert_hwp_color_to_hex(self.p.StrikeOutColor)

    @strike_out_color.setter
    def strike_out_color(self, color: Union[int, str, tuple]):
        value = convert_to_hwp_color(color)
        self.p.StrikeOutColor = value

    @property
    def strike_out_type(self):
        return self.p.StrikeOutType

    @strike_out_type.setter
    def strike_out_type(self, value):
        self.p.StrikeOutType = value

    # Underline
    #
    #
    @property
    def underline_shape(self):
        return self.p.UnderlineShape

    @underline_shape.setter
    def underline_shape(self, value):
        self.p.UnderlineShape = value

    @property
    def underline_type(self):
        return self.p.UnderlineType

    @underline_type.setter
    def underline_type(self, value):
        self.p.UnderlineType = value

    @property
    def underline_color(self):
        return convert_hwp_color_to_hex(self.p.UnderlineColor)

    @underline_color.setter
    def strike_out_color(self, color: Union[int, str, tuple]):
        value = convert_to_hwp_color(color)
        self.p.UnderlineColor = value

    def todict(self):
        values = asdict(self.p)

        return {key: value for key, value in values.items() if value is not None}

    def fromdict(self, values: dict):
        for key, value in values.items():
            setattr(self.p, key, value)

    def get_from_pset(self, pset):
        data_fields = fields(self.p)

        for field in data_fields:
            name = field.name
            value = getattr(pset, name)
            setattr(self.p, name, value)

# %% ../nbs/02_api/03_dataclasses.ipynb 7
@dataclass
class Paragraph:
    AlignType: Optional[int] = None
    AutoSpaceEAsianEng: Optional[int] = None
    AutoSpaceEAsianNum: Optional[int] = None
    BorderConnect: Optional[int] = None
    BorderOffsetBottom: Optional[int] = None
    BorderOffsetLeft: Optional[int] = None
    BorderOffsetRight: Optional[int] = None
    BorderOffsetTop: Optional[int] = None
    BorderText: Optional[int] = None
    BreakLatinWord: Optional[int] = None
    BreakNonLatinWord: Optional[int] = None
    Checked: Optional[int] = None
    Condense: Optional[int] = None
    FontLineHeight: Optional[int] = None
    HeadingType: Optional[int] = None
    Indentation: Optional[int] = None
    KeepLinesTogether: Optional[int] = None
    KeepWithNext: Optional[int] = None
    LeftMargin: Optional[int] = None
    Level: Optional[int] = None
    LineSpacing: Optional[int] = None
    LineSpacingType: Optional[int] = None
    LineWrap: Optional[int] = None
    NextSpacing: Optional[int] = None
    PagebreakBefore: Optional[int] = None
    PrevSpacing: Optional[int] = None
    RightMargin: Optional[int] = None
    SnapToGrid: Optional[int] = None
    SuppressLineNum: Optional[int] = None
    TailType: Optional[int] = None
    TextAlignment: Optional[int] = None
    WidowOrphan: Optional[int] = None

# %% ../nbs/02_api/03_dataclasses.ipynb 8
class ParaShape:
    """
    ParaShape 클래스는 문단 모양을 다룹니다. ParaShape는 문단에 적용되는 스타일링 속성을 나타냅니다.

    속성
    ----
    attributes : list
        ParaShape 클래스의 주요 속성 목록입니다.

    align_types, spacing_types, break_latin_words, text_alignments, heading_types : dict
        각각 문단의 정렬 유형, 줄 간격 유형, 라틴어 단어 분리 유형, 텍스트 정렬, 헤딩 유형을 나타내는 딕셔너리 입니다.

    메서드
    ------
    __init__(char_pset=None, **kwargs):
        ParaShape 클래스의 새 인스턴스를 초기화합니다. 키워드 인수를 사용하여 다양한 속성의 초기 값을 설정할 수 있습니다.

    __str__():
        ParaShape 인스턴스의 모든 속성과 그들의 현재 값이 나열된 문자열을 반환합니다.

    __repr__():
        __str__() 메서드와 같은 역할을 합니다.

    @property 와 @<property_name>.setter 쌍들:
        문단 모양의 다양한 속성에 대한 getter 및 setter 메서드를 정의합니다.

    todict():
        ParaShape 인스턴스를 딕셔너리를 변환합니다.

    fromdict(values: dict):
        딕셔너리에서 ParaShape 인스턴스를 채웁니다.

    get_from_pset(pset):
        주어진 hwp parameterset에서 ParaShape 인스턴스의 속성을 채웁니다.
    """

    attributes = [
        "left_margin",
        "right_margin",
        "indentation",
        "prev_spacing",
        "next_spacing",
        "line_spacing_type",
        "line_spacing",
        "align_type",
        "break_latin_word",
        "break_non_latin_word",
        "snap_to_grid",
        "condense",
        "widow_orphan",
        "keep_with_next",
        "page_break_before",
        "text_alignment",
        "font_line_height",
        "heading_type",
        "level",
        "border_connect",
        "border_text",
        "border_offset_left",
        "border_offset_right",
        "border_offset_top",
        "border_offset_bottom",
        "tail_type",
        "line_wrap",
    ]

    align_types = {
        "Both": 0,
        "Left": 1,
        "Right": 2,
        "Center": 3,
        "Distributed": 4,
        "SpaceOnly": 5,
    }

    spacing_types = {
        "Word": 0,
        "Fixed": 1,
        "Margin": 2,
    }

    break_latin_words = {
        "Word": 0,
        "Hyphen": 1,
        "Letter": 2,
    }

    text_alignments = {
        "Font": 0,
        "Upper": 1,
        "Middle": 2,
        "Lower": 3,
    }

    heading_types = {
        "None": 0,
        "Content": 1,
        "Numbering": 2,
        "Bullet": 3,
    }

    def __init__(self, char_pset=None, **kwargs):
        self.p = Paragraph()

        if char_pset:
            self.get_from_pset(char_pset)

        for key, value in kwargs.items():
            setattr(self, key, value)

        self.__repr__ = self.__str__

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        values = [getattr(self, attr) for attr in ParaShape.attributes]
        return f"<ParaShape: {', '.join([f'{attr}={value}' for attr, value in zip(ParaShape.attributes, values)])}>"

    @property
    def left_margin(self):
        return unit2mili(self.p.LeftMargin)

    @left_margin.setter
    def left_margin(self, value):
        self.p.LeftMargin = mili2unit(value)

    @property
    def right_margin(self):
        return unit2mili(self.p.RightMargin)

    @right_margin.setter
    def right_margin(self, value):
        self.p.RightMargin = mili2unit(value)

    @property
    def indentation(self):
        return unit2point(self.p.Indentation)

    @indentation.setter
    def indentation(self, value):
        self.p.Indentation = point2unit(value)

    @property
    def prev_spacing(self):
        return unit2point(self.p.PrevSpacing)

    @prev_spacing.setter
    def prev_spacing(self, value):
        self.p.PrevSpacing = point2unit(value)

    @property
    def next_spacing(self):
        return unit2point(self.p.NextSpacing)

    @next_spacing.setter
    def next_spacing(self, value):
        self.p.NextSpacing = point2unit(value)

    @property
    def line_spacing_type(self):
        return get_key(ParaShape.spacing_types, self.p.LineSpacingType)

    @line_spacing_type.setter
    def line_spacing_type(self, value: Union[str, int]):
        self.p.LineSpacingType = convert2int(ParaShape.spacing_types, value)

    @property
    def line_spacing(self):
        return (
            unit2point(self.p.LineSpacing)
            if self.p.LineSpacingType != 0
            else self.p.LineSpacing
        )

    @line_spacing.setter
    def line_spacing(self, value):
        self.p.LineSpacing = point2unit(value) if self.p.LineSpacingType != 0 else value

    @property
    def align_type(self):
        return get_key(ParaShape.align_types, self.p.AlignType)

    @align_type.setter
    def align_type(self, value):
        self.p.AlignType = convert2int(ParaShape.align_types, value)

    @property
    def break_latin_word(self):
        return get_key(ParaShape.break_latin_words, self.p.BreakLatinWord)

    @break_latin_word.setter
    def break_latin_word(self, value):
        self.p.BreakLatinWord = convert2int(ParaShape.break_latin_words, value)

    @property
    def break_non_latin_word(self):
        return self.p.BreakNonLatinWord

    @break_non_latin_word.setter
    def break_non_latin_word(self, value):
        self.p.BreakNonLatinWord = value

    @property
    def snap_to_grid(self):
        return self.p.SnapToGrid

    @snap_to_grid.setter
    def snap_to_grid(self, value):
        self.p.SnapToGrid = value

    @property
    def condense(self):
        return self.p.Condense

    @condense.setter
    def condense(self, value):
        self.p.Condense = value

    @property
    def widow_orphan(self):
        return self.p.WidowOrphan

    @widow_orphan.setter
    def widow_orphan(self, value):
        self.p.WidowOrphan = value

    @property
    def keep_with_next(self):
        return self.p.KeepWithNext

    @keep_with_next.setter
    def keep_with_next(self, value):
        self.p.KeepWithNext = value

    @property
    def page_break_before(self):
        return self.p.PagebreakBefore

    @page_break_before.setter
    def page_break_before(self, value):
        self.p.PageBreakBefore = value

    @property
    def text_alignment(self):
        return get_key(ParaShape.text_alignments, self.p.TextAlignment)

    @text_alignment.setter
    def text_alignment(self, value):
        self.p.TextAlignment = convert2int(ParaShape.text_alignments, value)

    @property
    def font_line_height(self):
        return self.p.FontLineHeight

    @font_line_height.setter
    def font_line_height(self, value):
        self.p.FontLineHeight = value

    @property
    def heading_type(self):
        return get_key(ParaShape.heading_types, self.p.HeadingType)

    @heading_type.setter
    def heading_type(self, value):
        self.p.HeadingType = convert2int(ParaShape.heading_types, value)

    @property
    def level(self):
        return self.p.Level

    @level.setter
    def level(self, value):
        self.p.Level = value

    @property
    def border_connect(self):
        return self.p.BorderConnect

    @border_connect.setter
    def border_connect(self, value):
        self.p.BorderConnect = value

    @property
    def border_text(self):
        return self.p.BorderText

    @border_text.setter
    def border_text(self, value):
        self.p.BorderText = value

    @property
    def border_offset_left(self):
        return unit2mili(self.p.BorderOffsetLeft)

    @border_offset_left.setter
    def border_offset_left(self, value):
        self.p.BorderOffsetLeft = mili2unit(value)

    @property
    def border_offset_right(self):
        return unit2mili(self.p.BorderOffsetRight)

    @border_offset_right.setter
    def border_offset_right(self, value):
        self.p.BorderOffsetRight = mili2unit(value)

    @property
    def border_offset_top(self):
        return unit2mili(self.p.BorderOffsetTop)

    @border_offset_top.setter
    def border_offset_top(self, value):
        self.p.BorderOffsetTop = mili2unit(value)

    @property
    def border_offset_bottom(self):
        return unit2mili(self.p.BorderOffsetBottom)

    @border_offset_bottom.setter
    def border_offset_bottom(self, value):
        self.p.BorderOffsetBottom = mili2unit(value)

    @property
    def tail_type(self):
        return self.p.TailType

    @tail_type.setter
    def tail_type(self, value):
        self.p.TailType = value

    @property
    def line_wrap(self):
        return self.p.LineWrap

    @line_wrap.setter
    def line_wrap(self, value):
        self.p.LineWrap = value

    def todict(self):
        values = asdict(self.p)

        return {key: value for key, value in values.items() if value is not None}

    def fromdict(self, values: dict):
        for key, value in values.items():
            setattr(self.p, key, value)

    def get_from_pset(self, pset):
        data_fields = fields(self.p)

        for field in data_fields:
            name = field.name
            value = getattr(pset, name)
            setattr(self.p, name, value)