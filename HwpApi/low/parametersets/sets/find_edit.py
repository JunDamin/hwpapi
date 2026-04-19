"""
Find / edit / navigation ParameterSet classes (13 classes).

Covers find/replace and editing operations:
- FindReplace (23 properties, uses CharShape/ParaShape from primitives)
- DocFindInfo, DocFilters, GotoE, SelectionOpt
- BookMark, IndexMark, MakeContents
- DeleteCtrls, ActionCrossRef, ExchangeFootnoteEndNote
- SaveFootnote, RevisionDef
"""
from __future__ import annotations
from hwpapi.low.parametersets import (
    ParameterSet, PropertyDescriptor, IntProperty, BoolProperty, StringProperty,
    ColorProperty, UnitProperty, MappedProperty, TypedProperty, NestedProperty,
    ArrayProperty, ListProperty,
)
# Python-level dependency: FindReplace references CharShape/ParaShape via NestedProperty
from hwpapi.low.parametersets.sets.primitives import CharShape, ParaShape


class FindReplace(ParameterSet):
    """
    ### FindReplace

    Find and replace text with character and paragraph formatting.

    This ParameterSet uses NestedProperty for auto-creating nested parameter sets.
    Access shape properties directly - they auto-create on first access!

    Example:
        >>> fr = FindReplace(action.CreateSet())
        >>> fr.FindString = "Python"
        >>> fr.MatchCase = True
        >>> # Auto-creates CharShape on first access - no create_itemset needed!
        >>> fr.FindCharShape.Bold = True
        >>> fr.FindCharShape.Size = 1200  # 12pt
    """
    # String properties
    FindString = StringProperty("FindString", "Text to find")
    ReplaceString = StringProperty("ReplaceString", "Text to replace")

    # Boolean properties
    MatchCase = BoolProperty("MatchCase", "Case sensitive matching")
    AllWordForms = BoolProperty("AllWordForms", "Match all word forms")
    SeveralWords = BoolProperty("SeveralWords", "Find multiple words")
    UseWildCards = BoolProperty("UseWildCards", "Use wildcards")
    WholeWordOnly = BoolProperty("WholeWordOnly", "Match whole words only")
    AutoSpell = BoolProperty("AutoSpell", "Auto spelling check")
    ReplaceMode = BoolProperty("ReplaceMode", "Replace mode enabled")
    IgnoreFindString = BoolProperty("IgnoreFindString", "Ignore find string")
    IgnoreReplaceString = BoolProperty("IgnoreReplaceString", "Ignore replace string")
    IgnoreMessage = BoolProperty("IgnoreMessage", "Suppress message boxes")
    HanjaFromHangul = BoolProperty("HanjaFromHangul", "Hanja from Hangul conversion")
    FindJaso = BoolProperty("FindJaso", "Find by Jaso (Korean characters)")
    FindRegExp = BoolProperty("FindRegExp", "Use regular expressions")

    # Integer/Mapped properties
    Direction = MappedProperty("Direction", {
        "down": 0, "up": 1, "all": 2
    }, "Search direction")
    FindType = BoolProperty("FindType", "Find type (True=find, False=find and select)")

    # Auto-creating nested shape properties using NestedProperty!
    FindCharShape = NestedProperty("FindCharShape", "CharShape", CharShape, "Find character shape")
    FindParaShape = NestedProperty("FindParaShape", "ParaShape", ParaShape, "Find paragraph shape")
    ReplaceCharShape = NestedProperty("ReplaceCharShape", "CharShape", CharShape, "Replace character shape")
    ReplaceParaShape = NestedProperty("ReplaceParaShape", "ParaShape", ParaShape, "Replace paragraph shape")

    # Style properties
    FindStyle = StringProperty("FindStyle", "Find style name")
    ReplaceStyle = StringProperty("ReplaceStyle", "Replace style name")



class ActionCrossRef(ParameterSet):
    """ActionCrossRef ParameterSet."""
    Command = PropertyDescriptor("Command", r"""※command string 참조 (한글2007에 새로 추가)""")


class BookMark(ParameterSet):
    """BookMark ParameterSet."""
    Name = PropertyDescriptor("Name", r"""책갈피 이름""")
    Type = PropertyDescriptor("Type", r"""책갈피 종류 : 0 \= 일반 책갈피, 1 \= 블록 책갈피""")
    Command = PropertyDescriptor("Command", r"""책갈피 명령의 종류 :0 \= 책갈피 생성, 1 \= 책갈피로 이동, 2 \= 책갈피 수정""")


class DeleteCtrls(ParameterSet):
    """DeleteCtrls ParameterSet."""
    DeleteCtrlType = PropertyDescriptor("DeleteCtrlType", r"""지울 개체 목록""")


class DocFilters(ParameterSet):
    """DocFilters ParameterSet."""
    DocFilters = PropertyDescriptor("DocFilters", r"""Parameter property""")
    Format = PropertyDescriptor("Format", r"""필터 리스트를 string array형태로 가져옴 (Export시에만)""")
    Type = PropertyDescriptor("Type", r"""Import 리스트를 가져올 것인지 Export 리스트를 가져올 것인지의 관한 타입.Import \= 1Export \= 0 (on input)""")


class DocFindInfo(ParameterSet):
    """DocFindInfo ParameterSet."""
    ListID = PropertyDescriptor("ListID", r"""현재 위치한 리스트""")
    ParaID = PropertyDescriptor("ParaID", r"""현재 위치한 문단""")
    Pos = PropertyDescriptor("Pos", r"""현재 위치한 글자""")


class ExchangeFootnoteEndNote(ParameterSet):
    """ExchangeFootnoteEndNote ParameterSet."""
    Flag = PropertyDescriptor("Flag", r"""옵션 0 : 모든 각주를 미주로 바꾸기1 : 모든 미주를 각주로 바꾸기2 : 각주와 미주를 서로 바꾸기""")


class GotoE(ParameterSet):
    """GotoE ParameterSet."""
    SetSelectionIndex = PropertyDescriptor("SetSelectionIndex", r"""현재 선택되어 있는 라디오 값 (한글2007에 새로 추가)""")
    DialogResult = PropertyDescriptor("DialogResult", r"""대화상자의 반환값 (한글2007에 새로 추가)""")


class IndexMark(ParameterSet):
    """IndexMark ParameterSet."""
    First = PropertyDescriptor("First", r"""첫 번째 키""")
    Second = PropertyDescriptor("Second", r"""두 번째 키""")


class MakeContents(ParameterSet):
    """MakeContents ParameterSet."""
    Make = PropertyDescriptor("Make", r"""생성할 차례의 종류, 다음의 값들의 조합이다0x01: 제목 차례0x02: 표 차례0x04: 그림 차례0x08: 수식 차례제목 차례를 지정한 경우에는 다음의 값을 추가로 지정할 수 있다.0x10: 개요 문단으로 모으기0x20: 스타일로 모으기0x40: 차례코드로 모으기""")
    Level = PropertyDescriptor("Level", r"""개요 수준""")
    AutoTabRight = PropertyDescriptor("AutoTabRight", r"""문단 오른쪽 끝 자동 탭 여부 : 0 \= 자동 탭 사용안함, 1 \= 자동 탭 사용""")
    Leader = PropertyDescriptor("Leader", r"""오른쪽 끝 탭 채울 모양(선 종류)""")
    OutlineNumber = PropertyDescriptor("OutlineNumber", r"""개요 문단 모으기""")
    Styles = PropertyDescriptor("Styles", r"""모을 스타일 목록""")
    StyleName = PropertyDescriptor("StyleName", r"""모을 스타일 이름들""")
    OutFileName = PropertyDescriptor("OutFileName", r"""만들 파일 이름. ""이면 현재 문서에 생성""")
    Position = PropertyDescriptor("Position", r"""만들 위치. 반드시 0이어야 한다. (한글 컨트롤은 탭이 없으므로)(한글2007에 새로 추가)0 \= 현재 문서1 \= 새 탭으로""")


class RevisionDef(ParameterSet):
    """RevisionDef ParameterSet."""
    SignType = PropertyDescriptor("SignType", r"""교정부호 종류 :0 \= 교정부호 없음1 \= 띄움표2 \= 줄 바꿈표3 \= 줄 비움표4 \= 메모형 고침표5 \= 지움표6 \= 붙임표7 \= 뺌표8 \= 줄 이음표9 \= 줄 붙임표10 \= 톱니표11 \= 생각표12 \= 칭찬표 13 \= 줄표14 \= 부호 넣음표15 \= 넣음표16 \= 고침표17 \= 자리 바꿈표18 \= 오른자리 옮김표19 \= 자료연결20 \= 왼자리 옮김표21 \= 부분자리 옮김표22 \= 줄 서로 바꿈표23 \= 자리바꿈 나눔표(내부용)24 \= 줄 서로 바꿈 나눔표(내부용)""")
    SubText = PropertyDescriptor("SubText", r"""교정 문자열교정 문자열을 가질 수 있는 교정부호만 적용. 나머지는 무시""")
    Margin = PropertyDescriptor("Margin", r"""여백(HWPUNIT). 오른자리 옮김표와 왼자리 옮김표일 경우에만 적용.""")
    BeginPos = PropertyDescriptor("BeginPos", r"""시작위치(HWPUNIT). 오른자리 옮김표와 왼자리 옮김표일 경우에만 적용.""")


class SaveFootnote(ParameterSet):
    """SaveFootnote ParameterSet."""
    FileName = PropertyDescriptor("FileName", r"""파일 이름""")
    Flag = PropertyDescriptor("Flag", r"""옵션1 : 각주 저장2 : 미주 저장3 : 각주/미주 저장""")


class SelectionOpt(ParameterSet):
    """SelectionOpt ParameterSet - 선택 옵션 (Copy/Cut/Paste/Erase 등)."""

