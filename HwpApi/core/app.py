"""
The :class:`App` class — main user-facing API for HWP automation.

Wraps an :class:`~hwpapi.core.engine.Engine` and provides high-level
methods for all common operations:

- **File I/O**: :meth:`open`, :meth:`save`, :meth:`save_block`,
  :meth:`insert_file`, :meth:`close`, :meth:`quit`
- **Text**: :meth:`insert_text`, :meth:`get_text`, :meth:`find_text`,
  :meth:`replace_all`, :meth:`get_selected_text`, :meth:`select_text`, :meth:`scan`
- **Formatting**: :meth:`charshape`, :meth:`get_charshape`,
  :meth:`set_charshape`, :meth:`get_parashape`, :meth:`set_parashape`
- **Images/Tables**: :meth:`insert_picture`, :meth:`set_cell_border`,
  :meth:`set_cell_color`
- **Page**: :meth:`setup_page`, :meth:`get_filepath`, :meth:`get_hwnd`
- **Window**: :meth:`set_visible`, :meth:`reload`
- **Actions/Fonts**: :meth:`create_action`, :meth:`create_parameterset`,
  :meth:`get_font_list`

Accessors attached as attributes (see :mod:`hwpapi.classes.accessors`):
- ``app.actions`` — dynamic action dispatcher
- ``app.move`` — caret navigation
- ``app.cell`` — table cell ops
- ``app.table`` — table structure ops
- ``app.page`` — page ops
"""
__all__ = ['Engine', 'Engines', 'Apps', 'App', 'move_to_line']

from contextlib import contextmanager
from pathlib import Path
import numbers
import hwpapi.constants as const
import sys
import warnings
from hwpapi.logging import get_logger

from hwpapi.actions import _Action, _Actions
from hwpapi.parametersets import ParaShape
import hwpapi.parametersets as parametersets
from hwpapi.classes import MoveAccessor, CellAccessor, TableAccessor, PageAccessor
from .engine import Engine, Engines, Apps
from hwpapi.functions import (
    check_dll,
    get_hwp_objects,
    dispatch,
    get_absolute_path,
    get_rgb_tuple,
    set_pset,
)



class App:
    """
    `App` 클래스는 한컴오피스의 한/글 프로그램과 상호작용하기 위한 인터페이스를 제공합니다.

    이 클래스는 한/글 프로그램의 COM 객체와의 연동을 관리하고, 해당 객체에 대한 작업을 수행하는 메서드들을 포함합니다.

    Attributes
    ----------
    engine : Engine, optional
        사용할 엔진 객체입니다.
    is_visible : bool, optional
        한/글 프로그램의 가시성을 설정합니다.
    dll_path : str, optional
        사용할 DLL 파일의 경로입니다.

    Methods
    -------
    __init__(engine=None, is_visible=True, dll_path=None)
        `App` 클래스의 인스턴스를 초기화합니다.
    api()
        현재 사용 중인 엔진의 구현체를 반환합니다.
    __str__()
        `App` 인스턴스의 문자열 표현을 반환합니다.
    """

    def __init__(
        self,
        new_app=False,
        is_visible=True,
        dll_path=None,
        engine=None,
    ):
        """
        `App` 클래스의 인스턴스를 초기화합니다.

        Parameters
        ----------
        new_app: bool, optional
            별도 앱을 열지를 결정합니다.
        is_visible : bool, optional
            한/글 프로그램의 가시성 설정. 기본값은 True입니다.
        dll_path : str, optional
            사용할 DLL 파일의 경로. 기본값은 None입니다.
        engine : Engine, optional
            한/글 프로그램과의 상호작용을 위한 엔진. 기본값이 없을 경우, 새로운 엔진 인스턴스가 생성됩니다.
        Notes
        -----
        engine이 제공되지 않은 경우, `Engines`를 통해 엔진을 생성하거나 선택합니다.
        """
        self.logger = get_logger('core')
        self.logger.debug(f"Initializing App with new_app={new_app}, is_visible={is_visible}, dll_path={dll_path}")

        self._load(new_app=new_app, dll_path=dll_path, engine=engine)
        self.actions = _Actions(self)
        self.parameters = self.api.HParameterSet
        self.set_visible(is_visible)
        self.logger.info(f"App window visibility set to: {is_visible}")

        self.move = MoveAccessor(self)
        self.cell = CellAccessor(self)
        self.table = TableAccessor(self)
        self.page = PageAccessor(self)
        self.logger.info("App initialized successfully with all accessors")

    def _load(self, new_app=False, engine=None, dll_path=None):
        self.logger.debug(f"Loading App with new_app={new_app}, engine={engine}, dll_path={dll_path}")

        if new_app:
            engine = Engine()
            self.logger.debug("Created new engine for new_app")
        if not engine:
            engines = Engines()
            engine = engines[-1] if len(engines) > 0 else Engine()
            self.logger.debug(f"Selected engine from engines collection")

        self.engine = engine
        self.logger.info(f"Engine loaded successfully")

        # Get DLL path (uses stable appdata location)
        if dll_path is None:
            from hwpapi.functions import get_hwp_dll_path
            dll_path = get_hwp_dll_path()

        # Register DLL if found
        if dll_path is not None:
            self.logger.info(f"Registering DLL: {dll_path}")
            check_dll(dll_path)
        else:
            self.logger.warning("No DLL file found - some functionality may be limited")

        try:
            self.api.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            self.logger.debug("Successfully registered FilePathCheckDLL module")
        except Exception as e:
            self.logger.warning(f"Failed to register FilePathCheckDLL module: {e}")
            warnings.warn(f"Failed to register FilePathCheckDLL module: {e}")

    @property
    def api(self):
        """
        현재 사용 중인 엔진의 구현체를 반환합니다.

        Returns
        -------
        object
            엔진의 구현체.
        """
        return self.engine.impl

    def __str__(self):
        """
        `App` 인스턴스의 문자열 표현을 반환합니다.

        Returns
        -------
        str
            `App` 인스턴스를 설명하는 문자열.
        """
        return f"<Hwp App: {self.get_filepath()}>"

    __repr__ = __str__

    def set_visible(self, is_visible=True, window_i=0):
        """
        한컴오피스 Hwp 프로그램 창의 가시성을 설정합니다.

        이 함수는 `is_visible` 매개변수에 따라 한컴오피스 Hwp 프로그램의 창을 표시하거나 숨기는 데 사용됩니다.
        `window_i` 매개변수는 표시하거나 숨길 창의 인덱스를 지정합니다.

        매개변수
        ----------
        is_visible : bool, optional
            프로그램 창의 가시성을 나타내는 불린 값.
            True이면 창이 표시되고, False이면 숨겨집니다. 기본값은 True입니다.
        window_i : int, optional
            수정할 창의 인덱스. 기본값은 0입니다.

        사용 예시
        --------
        >>> app = App()
        >>> app.set_visible(is_visible=True, window_i=0)
        >>> app.set_visible(is_visible=False, window_i=1)
        """
        logger = get_logger('core')
        logger.debug(f"Calling set_visible")
        self.api.XHwpWindows.Item(window_i).Visible = is_visible

    def create_parameterset(self, key: str):
        logger = get_logger('core')
        logger.debug(f"Calling create_parameterset")
        return getattr(self.api.HParameterSet, f"H{key}")

    def reload(self, new_app=False, dll_path=None):
        """
        새로운 HWPFrame.HwpObject로 `App` 인스턴스를 다시 로드하고 가시성과 DLL 경로를 재설정합니다.

        이 함수는 `app` 객체의 API를 새로운 HWPFrame.HwpObject 인스턴스로 다시 초기화하도록 설계되었습니다.
        또한 제공된 경우 지정된 DLL 파일을 확인하고 등록합니다.

        매개변수
        ----------
        dll_path : str, optional
            확인하고 등록할 DLL 파일의 경로. None인 경우 DLL 등록이 수행되지 않습니다.

        주의사항
        -----
        `reload` 함수는 HWPFrame.HwpObject의 상태를 재설정해야 하거나
        사용 중인 DLL을 변경할 때 유용합니다.

        사용 예시
        --------
        >>> app = App()
        >>> app.reload(dll_path="path/to/dll")
        """
        logger = get_logger('core')
        logger.debug(f"Calling reload")
        self._load(new_app=new_app, dll_path=dll_path)

    def get_filepath(self):
        """
        현재 열려있는 한컴오피스 Hwp 문서의 파일 경로를 검색합니다.

        이 함수는 `App` 인스턴스와 연결된 한컴오피스 Hwp 프로그램의 활성 문서에 접근하여
        전체 파일 경로를 반환합니다.

        반환값
        -------
        str
            현재 활성화된 한컴오피스 Hwp 문서의 전체 파일 경로.

        사용 예시
        --------
        >>> app = App()
        >>> filepath = app.get_filepath()
        >>> print(filepath)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_filepath")
        doc = self.api.XHwpDocuments.Active_XHwpDocument
        return doc.FullName

    def create_action(self, action_key: str):
        """
        `_Action` 클래스의 인스턴스를 생성하고 반환합니다.

        이 함수는 주어진 `action_key`와 연결된 새로운 `_Action` 인스턴스를 생성합니다.
        `action_key`는 애플리케이션에 대해 생성할 특정 액션의 유형을 지정합니다.

        매개변수
        ----------
        action_key : str
            생성할 특정 액션을 나타내는 키.

        반환값
        -------
        _Action
            제공된 애플리케이션 객체와 액션 키로 초기화된 `_Action` 클래스의 인스턴스.

        사용 예시
        --------
        >>> app = App()
        >>> action = app.create_action('some_action_key')
        >>> print(action)
        """
        logger = get_logger('core')
        logger.debug(f"Calling create_action")
        return _Action(self, action_key)

    def open(self, path: str):
        """
        제공된 파일 경로를 사용하여 한컴오피스 Hwp 프로그램에서 파일을 엽니다.

        이 함수는 먼저 제공된 파일 경로를 `get_absolute_path` 함수를 사용하여 절대 경로로 변환합니다.
        그런 다음 `api.Open` 메서드를 사용하여 한컴오피스 Hwp 프로그램에서 파일을 열고 열린 파일의 절대 경로를 반환합니다.

        매개변수
        ----------
        path : str
            열릴 문서의 파일 경로.

        반환값
        -------
        str
            열린 파일의 절대 경로.

        사용 예시
        --------
        >>> app = App()
        >>> opened_file_path = app.open('path/to/document.hwp')
        >>> print(opened_file_path)
        """
        logger = get_logger('core')
        logger.debug(f"Calling open")
        name = get_absolute_path(path)
        self.api.Open(name)
        return name

    def get_hwnd(self):
        """
        Retrieves the window handle (HWND) of the active window in the Hancom Office Hwp program.

        This function accesses the active window in the Hancom Office Hwp program linked to the `App` instance
        and returns its window handle. The window handle can be used in scenarios where direct manipulation
        or interaction with the window at the OS level is required.

        Returns
        -------
        int
            The window handle (HWND) of the active Hancom Office Hwp window.

        Examples
        --------
        >>> app = App()
        >>> hwnd = app.get_hwnd()
        >>> print(hwnd)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_hwnd")
        return self.api.XHwpWindows.Active_XHwpWindow.WindowHandle

    def save(self, path=None):
        """
        Saves the current document in the Hancom Office Hwp program.

        If a path is provided, the document is saved to that location in the appropriate format based on the file extension.
        Supported formats include HWP, PDF, HWPML2X (HWPX), and PNG. If no path is provided, the document is saved
        in its current location.

        Parameters
        ----------
        path : str, optional
            The file path where the document will be saved. If None, the document is saved in its current location.

        Returns
        -------
        str
            The file path where the document has been saved.

        Examples
        --------
        >>> app = App()
        >>> saved_file_path = app.save('path/to/document.hwp')
        >>> print(saved_file_path)
        """
        logger = get_logger('core')
        logger.debug(f"Calling save")
        if not path:
            self.api.Save()
            return self.get_filepath()

        name = get_absolute_path(path)
        extension = Path(name).suffix
        format_ = {
            ".hwp": "HWP",
            ".pdf": "PDF",
            ".hwpx": "HWPX",
            ".hml": "HWPML2X",
            ".png": "PNG",
            ".txt": "TEXT",
            ".docx": "MSWORD",
        }.get(extension)

        self.api.SaveAs(name, format_)
        return name

    def save_block(self, path: Path):
        """
        Saves a block of content in the Hancom Office Hwp program and returns the file path.

        This function saves a specified block from the Hancom Office Hwp program to a file in a given format.
        The format is determined by the file extension. Supported formats include HWP, PDF, HWPML2X (HWPX), and PNG.
        It returns the file path if the save operation is successful or None if it fails.

        Parameters
        ----------
        path : Path
            The file path where the block will be saved.

        Returns
        -------
        Path or None
            The path to the saved file if the save operation is successful, or None if it fails.

        Examples
        --------
        >>> app = App()
        >>> saved_path = app.save_block(Path('path/to/save/block.hwp'))
        >>> print(saved_path)
        """
        logger = get_logger('core')
        logger.debug(f"Calling save_block")

        name = get_absolute_path(path)
        extension = Path(name).suffix
        format_ = {
            ".hwp": "HWP",
            ".pdf": "PDF",
            ".hwpx": "HWPML2X",
            ".png": "PNG",
        }.get(extension)

        action = self.actions.SaveBlockAction
        p = action.pset

        p.filename = name
        p.Format = format_
        action.run(p)
        return name if Path(name).exists() else None

    def close(self):
        """
        Closes the currently open document in the Hancom Office Hwp program.

        This function triggers the 'FileClose' command within the Hancom Office Hwp program to close the current document.
        It's useful for programmatically managing documents within the application.

        Examples
        --------
        >>> app = App()
        >>> # Open and manipulate the document
        >>> app.close()
        """
        logger = get_logger('core')
        logger.debug(f"Calling close")
        self.api.Run("FileClose")

    def quit(self):
        """
        Terminates the Hancom Office Hwp program instance associated with the `App`.

        This function invokes the 'FileQuit' command within the Hancom Office Hwp program to close the application.
        It's useful for programmatically controlling the lifecycle of the application.

        Examples
        --------
        >>> app = App()
        >>> # Perform actions with the app
        >>> app.quit()
        """
        logger = get_logger('core')
        logger.debug(f"Calling quit")
        self.api.Run("FileQuit")

    def get_font_list(self):
        """
        Retrieves the list of fonts that used in the documnents from the current application.

        This method accesses the font list in the Hancom Office Hwp program linked to the `App` instance.
        It extracts fonts that used and return in list for further use or analysis.

        Returns
        -------
        List
            return font list that used in the hwp document.

        Examples
        --------
        >>> app = App()
        >>> font_list = app.get_font_list()
        >>> print(font_list)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_font_list")
        self.api.ScanFont()
        result = self.api.GetFontList()
        output = list(map(lambda txt: txt.split(",")[0],
            result.split("\x02")
            ))
        return output

    def charshape(self,
        facename=None,
        fonttype=None,
        size=None,
        ratio=None,
        spacing=None,
        offset=None,
        bold=None,
        italic=None,
        small_caps=None,
        emboss=None,
        engrave=None,
        superscript=None,
        subscript=None,
        underline_type=None,
        underline_shape=None,
        outline_type=None,
        shadow_type=None,
        text_color=None,
        shade_color=None,
        underline_color=None,
        shadow_color=None,
        shadow_offset_x=None,
        shadow_offset_y=None,
        strikeout_color=None,
        strikeout_type=None,
        strikeout_shape=None,
        diac_sym_mark=None,
        use_font_space=None,
        use_kerning=None,
        height=None,
        border_fill=None):
        """return null charshape
        """

        logger = get_logger('core')
        logger.debug(f"Calling charshape")

        charshape_pset = self.create_parameterset("CharShape")
        value_set = {
            "facename": facename,
            "fonttype": fonttype,
            "size": size,
            "ratio": ratio,
            "spacing": spacing,
            "offset": offset,
            "bold": bold,
            "italic": italic,
            "small_caps": small_caps,
            "emboss": emboss,
            "engrave": engrave,
            "superscript": superscript,
            "subscript": subscript,
            "underline_type": underline_type,
            "underline_shape": underline_shape,
            "outline_type": outline_type,
            "shadow_type": shadow_type,
            "text_color": text_color,
            "shade_color": shade_color,
            "underline_color": underline_color,
            "shadow_color": shadow_color,
            "shadow_offset_x": shadow_offset_x,
            "shadow_offset_y": shadow_offset_y,
            "strikeout_color": strikeout_color,
            "strikeout_type": strikeout_type,
            "strikeout_shape": strikeout_shape,
            "diac_sym_mark": diac_sym_mark,
            "use_font_space": use_font_space,
            "use_kerning": use_kerning,
            "height": height,
            "border_fil": border_fill}
        for key, value in value_set.items():
            if not key:
                continue
            setattr(charshape_pset, key, value)

        return charshape_pset

    def get_charshape(self):
        """
        Retrieves the character shape settings from the current application.

        This method accesses the character shape settings in the Hancom Office Hwp program linked to the `App` instance.
        It extracts these settings and encapsulates them in a `CharShape` object for further use or analysis.

        Returns
        -------
        CharShape
            An object representing the character shape settings of the application.

        Examples
        --------
        >>> app = App()
        >>> char_shape = app.get_charshape()
        >>> print(char_shape)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_charshape")
        return self.actions.CharShape.pset

    def set_charshape(self, charshape: parametersets.CharShape=None, **kwargs):
        """
        Sets the character shape in the current paragraph of the Hancom Office Hwp application using the provided `CharShape`.

        If `charshape` is None, a default instance of `CharShape` is created. The function can also accept additional keyword
        arguments (`kwargs`), which are assigned as properties of `charshape`. This modified `charshape` is then used to change
        the current paragraph shape in the Hancom Office Hwp document.

        Parameters
        ----------
        charshape : CharShape, optional
            The `CharShape` object to be used for setting character shapes in Hancom Office Hwp. The default is None.
        **kwargs
            Additional keyword arguments that are assigned as properties of `charshape`.

        Returns
        -------
        bool
            A boolean value indicating the success of the `set_charshape` operation.

        Examples
        --------
        >>> app = App()
        >>> char_shape = app.get_charshape()
        >>> success = app.set_charshape(charshape=char_shape, fontName='Arial', fontSize=10)
        >>> print(success)
        """
        logger = get_logger('core')
        logger.debug(f"Calling set_charshape")
        action = self.actions.CharShape
        pset = action.pset
        # charshape를 전달하면 반영해야 함
        if charshape:
            pset.update_from(charshape)

        for key, value in kwargs.items():
            setattr(pset, key, value)

        if action.run():
            return pset
        return None

    def get_parashape(self):
        """
        Retrieves the paragraph shape settings from the current application.

        This method accesses the paragraph shape settings in the Hancom Office Hwp program linked to the `App` instance.
        It extracts these settings and encapsulates them in a `ParaShape` object for further use or analysis.

        Returns
        -------
        ParaShape
            An object representing the paragraph shape settings of the application.

        Examples
        --------
        >>> app = App()
        >>> para_shape = app.get_parashape()
        >>> print(para_shape)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_parashape")
        return self.actions.ParagraphShape.pset

    def set_parashape(self, parashape: ParaShape = None, **kwargs):
        """
        Sets the paragraph shape in the Hancom Office Hwp application using the provided `ParaShape`.

        If `parashape` is None, a default instance of `ParaShape` is created. The function can also accept additional keyword
        arguments (`kwargs`), which are assigned as properties of `parashape`. This modified `parashape` is then used to change
        the current paragraph shape in the Hancom Office Hwp document.

        Parameters
        ----------
        parashape : ParaShape, optional
            The `ParaShape` object to be used for setting paragraph shapes in Hancom Office Hwp. Defaults to `None`.
        **kwargs
            Additional keyword arguments assigned as properties of `parashape`.
        if parashape:
            pset.update(parashape)

        bool
            A boolean value indicating the success of the `set_parashape` operation.

        Examples
        --------
        >>> app = App()
        >>> success = app.set_parashape(parashape=para_shape, align='Left', indent=10)
        >>> print(success)
        """
        logger = get_logger('core')
        logger.debug(f"Calling set_parashape")
        action = self.actions.ParagraphShape
        pset = action.pset
        if parashape:
            pset.update_from(parashape)

        for key, value in kwargs.items():
            setattr(pset, key, value)

        return action.run()

    def insert_text(
        self,
        text: str,
        charshape: parametersets.CharShape = None,
        **kwargs,
    ):
        """
        Inserts text into the Hancom Office Hwp document with specified character shape settings.

        This function inserts the given text into the document associated with the `App` instance.
        It allows for optional character shape settings through a `CharShape` object.
        Additional character shape attributes can be specified via keyword arguments.

        Parameters
        ----------
        text : str
            The text to be inserted into the document.
        charshape : CharShape, optional
            An optional `CharShape` object to specify the character shape settings for the inserted text. Defaults to `None`.
        **kwargs
            Additional character shape attributes to be set on the `charshape` object.

        Examples
        --------
        >>> app = App()
        >>> app.insert_text("Hello World", fontName="Arial", fontSize=12)
        """


        logger = get_logger('core')
        logger.debug(f"Calling insert_text")

        if not charshape:
            charshape = self.get_charshape()
        for key, value in kwargs.items():
            setattr(charshape, key, value)
        self.set_charshape(charshape)

        insert_text = self.actions.InsertText
        p = insert_text.pset
        p.text = text

        insert_text.run(p)
        return

    @contextmanager
    def scan(
        self,
        option=const.MaskOption.All,
        selection=False,
        scan_spos=const.ScanStartPosition.Current,
        scan_epos=const.ScanEndPosition.Document,
        spara=None,
        spos=None,
        epara=None,
        epos=None,
        scan_direction=const.ScanDirection.Forward,
    ):


        logger = get_logger('core')
        logger.debug(f"Calling scan")

        range_ = scan_spos.value + scan_epos.value
        if selection:
            range_ = 0x00FF  # Limit the scanning to the block if selection is True

        range_ += scan_direction.value
        self.api.InitScan(
            option=option.value,
            Range=range_,
            spara=spara,
            spos=spos,
            epara=epara,
            epos=epos,
        )
        yield _get_text(self)
        self.api.ReleaseScan()

    def setup_page(
        self,
        top=20,  # 위쪽 여백 (밀리미터 단위)
        bottom=10,  # 아래쪽 여백 (밀리미터 단위)
        right=20,  # 오른쪽 여백 (밀리미터 단위)
        left=20,  # 왼쪽 여백 (밀리미터 단위)
        header=15,  # 머리글 길이 (밀리미터 단위)
        footer=5,  # 바닥글 길이 (밀리미터 단위)
        gutter=0,  # 제본 여백 (밀리미터 단위)
    ):
        """
        Hancom Office Hwp 문서의 페이지 레이아웃을 설정합니다.

        이 함수는 페이지 여백, 머리글, 바닥글, 제본 여백의 크기를 설정합니다.
        단위는 밀리미터이며, 애플리케이션 내부의 단위 시스템으로 변환됩니다.

        매개변수
        ----------
        top : int, optional
            위쪽 여백 크기 (밀리미터 단위). 기본값은 20mm.
        bottom : int, optional
            아래쪽 여백 크기 (밀리미터 단위). 기본값은 10mm.
        right : int, optional
            오른쪽 여백 크기 (밀리미터 단위). 기본값은 20mm.
        left : int, optional
            왼쪽 여백 크기 (밀리미터 단위). 기본값은 20mm.
        header : int, optional
            머리글 길이 (밀리미터 단위). 기본값은 15mm.
        footer : int, optional
            바닥글 길이 (밀리미터 단위). 기본값은 5mm.
        gutter : int, optional
            제본 여백 크기 (밀리미터 단위). 기본값은 0mm.

        반환값
        -------
        bool
            페이지 설정이 성공하면 True, 실패하면 False.

        사용 예시
        --------
        >>> app = App()
        >>> app.setup_page(top=25, bottom=15, right=20, left=20, header=10, footer=5, gutter=5)
        """


        logger = get_logger('core')
        logger.debug(f"Calling setup_page")

        action = self.actions.PageSetup
        p = action.pset

        # 각 여백 및 길이를 밀리미터 단위에서 한컴오피스 내부 단위로 변환하여 설정
        p.PageDef.TopMargin = self.api.MiliToHwpUnit(top)
        p.PageDef.HeaderLen = self.api.MiliToHwpUnit(header)
        p.PageDef.RightMargin = self.api.MiliToHwpUnit(right)
        p.PageDef.BottomMargin = self.api.MiliToHwpUnit(bottom)
        p.PageDef.FooterLen = self.api.MiliToHwpUnit(footer)
        p.PageDef.LeftMargin = self.api.MiliToHwpUnit(left)
        p.PageDef.GutterLen = self.api.MiliToHwpUnit(gutter)

        return action.run()  # 페이지 설정 실행

    def insert_picture(
        self,
        fpath,
        width=None,
        height=None,
        size_option=const.SizeOption.RealSize,
        reverse=False,
        watermark=False,
        effect=const.Effect.RealPicture,
    ):
        """
        지정된 크기와 효과 옵션으로 문서에 이미지를 삽입합니다.

        매개변수
        ----------
        fpath : str
            삽입할 이미지 파일 경로.
        width : int, optional
            이미지의 너비. None이면 크기 옵션에 따라 결정됩니다. 기본값은 None.
        height : int, optional
            이미지의 높이. None이면 크기 옵션에 따라 결정됩니다. 기본값은 None.
        size_option : SizeOption, optional
            이미지의 크기 옵션으로 `SizeOption` Enum에 정의되어 있습니다. 기본값은 SizeOption.RealSize.
        reverse : bool, optional
            True이면 이미지를 반전합니다. 기본값은 False.
        watermark : bool, optional
            True이면 이미지를 워터마크로 처리합니다. 기본값은 False.
        effect : Effect, optional
            이미지의 시각적 효과로 `Effect` Enum에 정의되어 있습니다. 기본값은 Effect.RealPicture.

        반환값
        -------
        bool
            이미지가 성공적으로 삽입되었으면 True, 그렇지 않으면 False.

        사용 예시
        --------
        >>> app = App()
        >>> success = app.insert_picture('path/to/image.jpg', width=100, height=200, size_option=SizeOption.SpecificSize, effect=Effect.GrayScale)
        >>> print(success)
        """


        logger = get_logger('core')
        logger.debug(f"Calling insert_picture")

        path = Path(fpath)  # 이미지 파일 경로 처리
        size_option = size_option.value  # 크기 옵션 값 설정
        effect = effect.value  # 효과 옵션 값 설정

        return self.api.InsertPicture(
            path.absolute().as_posix(),  # 파일의 절대 경로
            Width=width,  # 이미지 너비
            Height=height,  # 이미지 높이
            sizeoption=size_option,  # 크기 옵션
            reverse=reverse,  # 반전 여부
            watermark=watermark,  # 워터마크 여부
            effect=effect,  # 시각적 효과
        )

    def select_text(self, option=const.SelectionOption.Line):
        """
        지정된 옵션에 따라 문서에서 텍스트를 선택합니다.

        매개변수
        ----------
        option : SelectionOption, optional
            선택할 텍스트 단위. SelectionOption Enum에 정의된 옵션을 사용. 기본값은 SelectionOption.Line.

        반환값
        -------
        tuple
            시작 및 끝 이동 작업의 결과를 포함하는 튜플 (둘 다 boolean 값).

        사용 예시
        --------
        >>> app = App()
        >>> app.select_text(option=SelectionOption.Para)
        """

        logger = get_logger('core')
        logger.debug(f"Calling select_text with option={option}")

        # 문자열 입력 처리
        if isinstance(option, str):
            try:
                option = const.SelectionOption[option]  # 문자열을 Enum으로 변환
            except KeyError:
                raise ValueError(f"Invalid option string: {option}. Must be one of {[o.name for o in const.SelectionOption]}")

        # Enum 멤버라면 값 꺼내기
        begin_action_name, end_action_name = option.value
        begin_action = getattr(self.actions, begin_action_name)
        end_action = getattr(self.actions, end_action_name)

        return begin_action(), end_action()

    def get_selected_text(self):
        """
        Hancom Office Hwp 문서에서 현재 선택된 영역의 텍스트를 가져옵니다.

        이 함수는 문서에서 선택된 텍스트를 스캔하여 문자열로 반환합니다.
        현재 강조 표시되거나 선택된 텍스트를 처리하는 작업에 특히 유용합니다.

        반환값
        -------
        str
            문서에서 현재 선택된 영역의 텍스트.

        사용 예시
        --------
        >>> app = App()
        >>> selected_text = app.get_selected_text()
        >>> print(selected_text)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_selected_text")

        with self.scan(selection=True) as scan:  # 선택된 영역 스캔
            text = "\n".join(scan)  # 선택된 텍스트를 줄 단위로 결합
        return text  # 결합된 텍스트 반환

    def get_text(self, spos=const.ScanStartPosition.Line, epos=const.ScanEndPosition.Line):
        """
        지정된 시작 및 끝 위치에 따라 문서에서 텍스트를 가져옵니다.

        매개변수
        ----------
        spos : ScanStartPosition, optional
            텍스트 검색을 시작할 위치. 기본값은 ScanStartPosition.Line.
        epos : ScanEndPosition, optional
            텍스트 검색을 종료할 위치. 기본값은 ScanEndPosition.Line.

        반환값
        -------
        str
            지정된 시작 위치부터 끝 위치까지 가져온 텍스트.

        사용 예시
        --------
        >>> app = App()
        >>> text = app.get_text(spos=ScanStartPosition.Paragraph, epos=ScanEndPosition.Paragraph)
        >>> print(text)
        """
        logger = get_logger('core')
        logger.debug(f"Calling get_text")

        with self.scan(scan_spos=spos, scan_epos=epos) as txts:  # 지정된 위치 범위에서 텍스트 스캔
            text = "".join(txts)  # 스캔한 텍스트를 하나의 문자열로 결합
        return text  # 결합된 텍스트 반환

    def find_text(
        self,
        text="",  # 찾을 텍스트
        ignore_message=True,  # 메시지 무시 여부
        direction=0,  # 검색 방향
        match_case=None,  # 대소문자 구분
        all_word_forms=None,  # 모든 단어 형태 검색
        several_words=None,  # 여러 단어 검색
        use_wild_cards=None,  # 와일드카드 사용
        whole_word_only=None,  # 전체 단어만 검색
        replace_mode=None,  # 찾아 바꾸기 모드
        ignore_find_string=None,  # 찾을 문자열 무시
        ignore_replace_string=None,  # 바꿀 문자열 무시
        find_style="",  # 찾을 스타일
        replace_style="",  # 바꿀 스타일
        find_jaso=None,  # 자소로 검색
        find_regexp=None,  # 정규표현식으로 검색
        find_type=True,  # 마지막 검색(True), 새 검색(False)
        facename=None,
        text_color=None,
        bold = None,
        charshape: parametersets.CharShape = None,

    ):
        """
        문서에서 다양한 옵션을 사용해 특정 텍스트를 검색합니다.

        매개변수
        ----------
        text : str
            검색할 텍스트 문자열.
        [기타 매개변수...]

        반환값
        -------
        bool
            텍스트를 찾았으면 True, 그렇지 않으면 False.

        사용 예시
        --------
        >>> app = App()
        >>> found = app.find_text(text="Hello", direction=Direction.Forward)
        >>> print(found)
        """


        logger = get_logger('core')
        logger.debug(f"Calling find_text")

        # 반복 검색 액션 생성
        action = self.actions.RepeatFind
        pset = action.pset
        # pset.FindCharShape.update_from(self.charshape())

        if charshape:
            pset.FindCharShape
            pset.FindCharShape.update_from(charshape)

        # 옵션 설정
        if text is not None:
            pset.FindString = text
        if ignore_message is not None:
            pset.ignore_message = ignore_message
        if match_case is not None:
            pset.mach_case = match_case
        if all_word_forms is not None:
            pset.all_word_forms = all_word_forms
        if direction is not None:
            pset.direction = direction
        if several_words is not None:
            pset.several_words = several_words
        if use_wild_cards is not None:
            pset.use_wild_cards = use_wild_cards
        if whole_word_only is not None:
            pset.whole_word_only = whole_word_only
        if replace_mode is not None:
            pset.replace_mode = replace_mode
        if ignore_find_string is not None:
            pset.IgnoreFindString = ignore_find_string
        if ignore_replace_string is not None:
            pset.IgnoreReplaceString = ignore_replace_string
        if find_style is not None:
            pset.FindStyle = find_style
        if replace_style is not None:
            pset.ReplaceStyle = replace_style
        if find_jaso is not None:
            pset.FindJaso = find_jaso
        if find_regexp is not None:
            pset.FindRegexp = find_regexp
        if find_type is not None:
            pset.FindType = find_type
        if facename is not None:
            pset.FindCharShape
            pset.FindCharShape.FaceNameHangul = facename
            from hwpapi.constants import korean_fonts, english_fonts
            fonts = set(korean_fonts + english_fonts)
            pset.FindCharShape.FontTypeHangul = 1 if facename not in fonts else 2
        if text_color is not None:
            pset.FindCharShape.TextColor = text_color
        if bold is not None:
            pset.FindCharShape.Bold = bold

        return action.run()

    def replace_all(
        self,
        old_text="",
        new_text="",
        old_charshape=None,  # 기존 문자 모양
        new_charshape=None,  # 새로운 문자 모양
        ignore_message=True,  # 메시지 무시 여부
        direction=None,  # 검색 방향
        match_case=None,  # 대소문자 구분
        all_word_forms=None,  # 모든 단어 형태
        several_words=None,  # 여러 단어 검색
        use_wild_cards=None,  # 와일드카드 사용
        whole_word_only=None,  # 전체 단어만 검색
        auto_spell=None,  # 자동 철자 교정
        replace_mode=None,  # 찾아 바꾸기 모드
        ignore_find_string=None,  # 찾을 문자열 무시
        ignore_replace_string=None,  # 바꿀 문자열 무시
        find_regexp=None,  # 정규표현식으로 검색
        find_style=None,  # 찾을 스타일
        replace_style=None,  # 바꿀 스타일
        find_jaso=None,  # 자소로 검색
        find_reg_exp=None,  # 정규표현식으로 검색
        find_type=None,  # 마지막 검색 사용(True), 새 검색(False)
    ):
        """
        문서에서 특정 텍스트를 새 텍스트로 모두 교체합니다.

        매개변수
        ----------
        old_text : str
            교체할 텍스트 문자열.
        new_text : str
            대체할 새 텍스트 문자열.
        [기타 매개변수...]

        반환값
        -------
        bool
            교체 작업이 성공하면 True, 실패하면 False.

        사용 예시
        --------
        >>> app = App()
        >>> success = app.replace_all(old_text="Hello", new_text="Hi", direction=Direction.All)
        >>> print(success)
        """


        logger = get_logger('core')
        logger.debug(f"Calling replace_all")

        # 반복 검색 액션 생성
        action = self.actions.AllReplace
        pset = action.pset

        # 옵션 설정
        if old_text is not None:
            pset.FindString = old_text
        if new_text is not None:
            pset.ReplaceString = new_text
        if ignore_message is not None:
            pset.ignore_message = ignore_message
        if match_case is not None:
            pset.mach_case = match_case
        if all_word_forms is not None:
            pset.all_word_forms = all_word_forms
        if direction is not None:
            pset.direction = direction
        if several_words is not None:
            pset.several_words = several_words
        if use_wild_cards is not None:
            pset.use_wild_cards = use_wild_cards
        if whole_word_only is not None:
            pset.whole_word_only = whole_word_only
        if replace_mode is not None:
            pset.replace_mode = replace_mode
        if ignore_find_string is not None:
            pset.IgnoreFindString = ignore_find_string
        if ignore_replace_string is not None:
            pset.IgnoreReplaceString = ignore_replace_string
        if find_style is not None:
            pset.FindStyle = find_style
        if replace_style is not None:
            pset.ReplaceStyle = replace_style
        if find_jaso is not None:
            pset.FindJaso = find_jaso
        if find_regexp is not None:
            pset.FindRegexp = find_regexp
        if find_type is not None:
            pset.FindType = find_type
        if auto_spell is not None:
            pset.auto_spell = auto_spell

        # 문자 모양 설정
        if old_charshape:
            pset.FindCharShape.update_from(old_charshape)
        if new_charshape:
            pset.replace_charshape.update_from(new_charshape)
        return action.run()

    def insert_file(
        self,
        fpath,
        keep_charshape=False,
        keep_parashape=False,
        keep_section=False,
        keep_style=False,
    ):
        """
        Inserts the contents of a specified file into the current document.

        This function inserts the contents of another file into the current document at the cursor's position.
        It provides options to retain various formatting attributes of the inserted content.

        Parameters
        ----------
        fpath : str
            The file path of the document to be inserted.
        keep_charshape : bool, optional
            If True, retains the original character shapes from the inserted file. Defaults to False.
        keep_parashape : bool, optional
            If True, retains the original paragraph shapes from the inserted file. Defaults to False.
        keep_section : bool, optional
            If True, retains the original section formatting from the inserted file. Defaults to False.
        keep_style : bool, optional
            If True, retains the original styles from the inserted file. Defaults to False.

        Returns
        -------
        bool
            True if the file was inserted successfully, False otherwise.

        Examples
        --------
        >>> app = App()
        >>> success = app.insert_file('path/to/file.hwp', keep_style=True)
        >>> print(success)
        """


        logger = get_logger('core')
        logger.debug(f"Calling insert_file")


        action = self.actions.InsertFile
        p = action.pset
        p.filename = Path(fpath).absolute().as_posix()
        p.KeepCharshape = keep_charshape
        p.KeepParashape = keep_parashape
        p.KeepSection = keep_section
        p.KeepStyle = keep_style

        return action.run()

    def set_cell_border(
        self,
        top=None,
        right=None,
        left=None,
        bottom=None,
        horizontal=None,
        vertical=None,
        top_width=const.Thickness.NULL,
        right_width=const.Thickness.NULL,
        left_width=const.Thickness.NULL,
        bottom_width=const.Thickness.NULL,
        horizontal_width=const.Thickness.NULL,
        vertical_width=const.Thickness.NULL,
        top_color=None,
        bottom_color=None,
        left_color=None,
        right_color=None,
        horizontal_color=None,
        vertical_color=None,
    ):
        """
        Sets the border properties for cells in a table within a Hancom Office Hwp document.

        This function customizes the border types, widths, and colors for different sides of the cells.
        It allows for detailed customization of cell appearance in tables.

        Parameters
        ----------
        top, right, left, bottom, horizontal, vertical : int, optional
            Types of borders for the respective sides and internal lines of the cell (e.g., solid, dashed).
        top_width, right_width, left_width, bottom_width, horizontal_width, vertical_width : Thickness, optional
            Widths of the borders corresponding to the sides and internal lines of the cell. Specified using the `Thickness` Enum.
        top_color, bottom_color, left_color, right_color, horizontal_color, vertical_color : str, optional
            Colors for the respective borders in a string format (e.g., "#RRGGBB").

        Returns
        -------
        bool
            True if the border settings were successfully applied, False otherwise.

        Examples
        --------
        >>> app = App()
        >>> success = app.set_cell_border(top=1, bottom=1, top_width=Thickness._0_25_mm, bottom_width=Thickness._0_25_mm, top_color="#FF0000", bottom_color="#00FF00")
        >>> print(success)

        Notes
        -----
        The function relies on the `self.actions.CellFill` action to set the border properties. The `Thickness` Enum provides predefined thickness levels for the borders. The color parameters should be provided in hex format.
        """
        attrs = {
            "BorderTypeTop": top,
            "BorderTypeRight": right,
            "BorderTypeLeft": left,
            "BorderTypeBottom": bottom,
            "TypeHorz": horizontal,
            "TypeVert": vertical,
            "BorderWidthTop": top_width.value,
            "BorderWidthRight": right_width.value,
            "BorderWidthLeft": left_width.value,
            "BorderWidthBottom": bottom_width.value,
            "WidthHorz": horizontal_width.value,
            "WidthVert": vertical_width.value,
            "BorderColorTop": self.api.RGBColor(get_rgb_tuple(top_color))
            if top_color
            else None,
            "BorderColorRight": self.api.RGBColor(get_rgb_tuple(right_color))
            if right_color
            else None,
            "BorderColorLeft": self.api.RGBColor(get_rgb_tuple(left_color))
            if left_color
            else None,
            "BorderColorBottom": self.api.RGBColor(get_rgb_tuple(bottom_color))
            if bottom_color
            else None,
            "ColorHorz": self.api.RGBColor(get_rgb_tuple(horizontal_color))
            if horizontal_color
            else None,
            "ColorVert": self.api.RGBColor(get_rgb_tuple(vertical_color))
            if vertical_color
            else None,
        }


        logger = get_logger('core')
        logger.debug(f"Calling set_cell_border")

        action = self.actions.CellFill
        p = action.pset

        for key, value in attrs.items():
            if value is None:
                continue
            setattr(p, key, value)

        return action.run()

    def set_cell_color(
        self, bg_color=None, hatch_color="#000000", hatch_style=6, alpha=None
    ):
        """
        Sets the background color and hatch style for cells in a Hancom Office Hwp document.

        This function allows customization of the cell background, including color and hatching pattern,
        providing options to set transparency and hatch color.

        Parameters
        ----------
        bg_color : str, optional
            Hexadecimal color code for the cell background (e.g., "#RRGGBB"). If not specified, the background is not changed.
        hatch_color : str, optional
            Hexadecimal color code for the hatching. Default is black ("#000000").
        hatch_style : int, optional
            Style of the hatching pattern. Default is 6.
        alpha : int, optional
            Alpha value for the background color's transparency (0-255). If not specified, transparency is not changed.

        Returns
        -------
        bool
            True if the cell color settings were applied successfully, False otherwise.

        Examples
        --------
        >>> app = App()
        >>> success = app.set_cell_color(bg_color="#FF0000", hatch_color="#0000FF", hatch_style=5, alpha=128)
        >>> print(success)

        Notes
        -----
        The function uses the `self.actions.CellBorderFill` action for setting the cell properties.
        Colors are specified in hexadecimal format. The alpha parameter controls the transparency of the background color.
        """


        logger = get_logger('core')
        logger.debug(f"Calling set_cell_color")

        fill_type = windows_brush = None
        if bg_color:
            fill_type = 1
            windows_brush = 1

        attrs = {
            "type": fill_type,
            "WindowsBrush": windows_brush,
            "WinBrushFaceColor": self.api.RGBColor(*get_rgb_tuple(bg_color)) if bg_color else None,
            "WinBrushAlpha": alpha,
            "WinBrushHatchColor": self.api.RGBColor(*get_rgb_tuple(hatch_color)) if hatch_color else None,
            "WinBrushFaceStyle": hatch_style,
        }

        action = self.actions.CellBorderFill
        p = action.pset

        for key, value in attrs.items():
            if value is not None:
                setattr(p.FillAttr, key, value)

        return action.run()


def _get_text(app):
    """스캔한 텍스트 텍스트 제너레이터"""
    flag, text = 2, ""
    while flag not in [0, 1, 101, 102]:
        flag, text = app.api.GetText()
        yield text

def move_to_line(app: App, text):
    """인자로 전달한 텍스트가 있는 줄의 시작지점으로 이동합니다."""
    logger = get_logger('core')
    logger.debug(f"Calling move_to_line")
    with app.scan(scan_spos=const.ScanStartPosition.Line) as scan:
        for line in scan:
            if text in line:
                return app.move.scan_pos()
    return False
