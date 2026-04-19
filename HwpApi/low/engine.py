"""
Engine, Engines, and Apps — low-level HWP application lifecycle management.

- :class:`Engine`: Wraps a single ``HWPFrame.HwpObject`` COM object.
- :class:`Engines`: Collection of ``Engine`` instances; discovers running HWP processes.
- :class:`Apps`: Collection of ``App`` instances (one per Engine).

These are rarely used directly — most users interact with
:class:`~hwpapi.core.App` instead, which wraps an ``Engine``.
"""
__all__ = ['Engine', 'Engines', 'Apps', 'App', 'move_to_line']

from contextlib import contextmanager
from pathlib import Path
import numbers
import hwpapi.constants as const
import sys
import warnings
from hwpapi.logging import get_logger

from hwpapi.low.actions import _Action, _Actions
from hwpapi.low.parametersets import ParaShape
import hwpapi.low.parametersets as parametersets
from hwpapi.classes import MoveAccessor, CellAccessor, TableAccessor, PageAccessor
from hwpapi.functions import (
    check_dll,
    get_hwp_objects,
    dispatch,
    get_absolute_path,
    get_rgb_tuple,
    set_pset,
)



class Engine:
    """
    한컴오피스 Hwp 객체를 캡슐화하는 Engine 클래스입니다.

    이 클래스는 한컴오피스 Hwp 애플리케이션과 상호작용하는 인터페이스를 제공하며,
    Hwp 환경 내에서 작업과 동작을 용이하게 합니다.

    매개변수
    ----------
    hwp_object : object, optional
        Engine에 의해 캡슐화될 Hwp 객체. 제공되지 않은 경우,
        "HWPFrame.HwpObject"를 사용하여 새로운 Hwp 객체를 생성합니다.

    속성
    ----------
    impl : object
        Hwp 객체의 구현체.

    메서드
    -------
    name()
        Hwp 객체의 이름(CLSID)을 반환합니다.

    사용 예시
    --------
    >>> engine = Engine()
    >>> print(engine.name)
    """

    def __init__(self, hwp_object=None):
        """
        한컴 Hwp 객체로 Engine을 초기화합니다.

        `hwp_object`가 제공되지 않으면 새로운 Hwp 객체가 생성됩니다. 생성에 실패하면 `impl`이 None으로 설정됩니다.

        매개변수
        ----------
        hwp_object : object, optional
            Engine에 의해 캡슐화될 Hwp 객체. 기본값은 "HWPFrame.HwpObject"입니다.
        """
        self.logger = get_logger('core')
        try:
            if not hwp_object:
                hwp_object = "HWPFrame.HwpObject"
            self.logger.debug(f"Initializing Engine with hwp_object: {hwp_object}")
            self.impl = dispatch(hwp_object)
            # v0.0.24+: Engine 식별 정보 INFO 로깅 — AI/디버깅 친화
            try:
                import os
                version = self.impl.Version if self.impl else "?"
                clsid = self.impl.CLSID if self.impl else "None"
                self.logger.info(
                    f"Engine ready: PID={os.getpid()} version={version} clsid={clsid}"
                )
            except Exception:
                self.logger.info(
                    f"Engine initialized successfully with CLSID: "
                    f"{self.impl.CLSID if self.impl else 'None'}"
                )
        except Exception as e:
            self.logger.error(f"Failed to initialize Hwp object: {e}")
            self.impl = None

    @property
    def name(self):
        """
        Hwp 객체의 이름(CLSID)을 반환합니다.

        반환값
        -------
        str
            Hwp 객체의 CLSID. 객체가 초기화되지 않은 경우 None을 반환합니다.
        """
        return self.impl.CLSID if self.impl else None

    def __repr__(self):
        """
        Engine 객체의 문자열 표현을 반환합니다.

        반환값
        -------
        str
            Engine의 문자열 표현. 엔진이 초기화되지 않은 경우 'Uninitialized'로 표시됩니다.
        """
        engine_name = self.name if self.name else "Uninitialized"
        return f"<Engine {engine_name}>"


class Engines:
    """
    여러 Engine 인스턴스를 관리하는 컬렉션 매니저입니다.

    이 클래스는 여러 Engine 인스턴스를 관리하며, 이들에 접근하고 반복하는 메서드를 제공합니다.
    여러 한컴오피스 Hwp 객체를 처리하는 데 유용합니다.

    매개변수
    ----------
    dll_path : str, optional
        초기화에 필요한 경우 DLL 파일의 경로.

    속성
    ----------
    active : Engine or None
        현재 활성화된 Engine 인스턴스.
    engines : list of Engine
        이 컬렉션에서 관리하는 Engine 인스턴스 목록.

    메서드
    -------
    add(engine)
        컬렉션에 새로운 Engine 인스턴스를 추가합니다.
    count()
        컬렉션의 Engine 인스턴스 수를 반환합니다.

    사용 예시
    --------
    >>> engines = Engines()
    >>> engines.add(Engine())
    >>> print(engines.count)
    >>> for engine in engines:
    ...     print(engine)

    주의사항
    -----
    `Engines` 클래스는 `get_hwp_objects()` 함수에서 검색된 각 객체에 대해 Engine 인스턴스를 생성하여 초기화됩니다.
    `dll_path`가 제공된 경우 필요한 DLL을 확인합니다.
    """

    def __init__(self, dll_path=None):
        """
        사용 가능한 Hwp 객체로 Engines 컬렉션을 초기화합니다.

        매개변수
        ----------
        dll_path : str, optional
            초기화에 필요한 경우 DLL 파일의 경로.
        """
        self.active = None
        self.engines = [Engine(hwp_object) for hwp_object in get_hwp_objects()]
        check_dll(dll_path)

    def add(self, engine):
        """
        컬렉션에 새로운 Engine 인스턴스를 추가합니다.

        매개변수
        ----------
        engine : Engine
            컬렉션에 추가할 Engine 인스턴스.
        """
        self.engines.append(engine)

    @property
    def count(self):
        """
        컬렉션의 Engine 인스턴스 수를 반환합니다.

        반환값
        -------
        int
            Engine 인스턴스의 수.
        """
        return len(self)

    def __call__(self, ind):
        """
        인덱스를 기준으로 컬렉션에서 Engine 인스턴스를 반환합니다.

        매개변수
        ----------
        ind : int
            검색할 Engine 인스턴스의 인덱스.

        반환값
        -------
        Engine
            지정된 인덱스의 Engine 인스턴스.
        """
        if isinstance(ind, numbers.Number):
            return self.engines[ind]

    def __getitem__(self, ind):
        """
        인덱스를 기준으로 컬렉션에서 Engine 인스턴스를 반환합니다.

        매개변수
        ----------
        ind : int
            검색할 Engine 인스턴스의 인덱스.

        반환값
        -------
        Engine
            지정된 인덱스의 Engine 인스턴스.
        """
        if isinstance(ind, numbers.Number):
            return self.engines[ind]

    def __len__(self):
        """
        컬렉션의 Engine 인스턴스 수를 반환합니다.

        반환값
        -------
        int
            Engine 인스턴스의 수.
        """
        return len(self.engines)

    def __iter__(self):
        """
        컬렉션의 각 Engine 인스턴스를 반환합니다.

        반환값
        ------
        Engine
            컬렉션의 각 Engine 인스턴스.
        """
        for engine in self.engines:
            yield engine

    def __repr__(self):
        """
        Engines 컬렉션의 문자열 표현을 반환합니다.

        반환값
        -------
        str
            Engines 컬렉션의 문자열 표현.
        """
        return f"<HWP Engines activated: {len(self.engines)}>"


class Apps:
    """
    모든 `app <App>` 객체의 컬렉션입니다.

    속성
    ----------
    _apps : list
        `App` 인스턴스를 포함하는 리스트.

    메서드
    -------
    add(**kwargs)
        새로운 App을 생성하고 컬렉션에 추가합니다.
    count()
        컬렉션의 앱 수를 반환합니다.
    cleanup()
        [메서드 설명 필요]
    """

    def __init__(self):
        """
        `Engines`의 각 `engine`에 대해 `App` 인스턴스를 생성하여 `Apps` 컬렉션을 초기화합니다.
        """
        self._apps = [App(engine=engine) for engine in Engines()]

    def add(self, **kwargs):
        """
        새로운 App을 생성합니다. 새로운 App이 활성화됩니다.

        반환값
        -------
        App
            새로 생성된 App 객체.
        """
        app = App(engine=Engine())
        self._apps.append(app)
        return app

    def __call__(self, i):
        """
        `Apps` 인스턴스를 함수처럼 호출할 수 있게 하여, 지정된 인덱스의 앱을 반환합니다.

        매개변수
        ----------
        i : int
            검색할 앱의 인덱스.

        반환값
        -------
        App
            지정된 인덱스의 앱.
        """
        return self[i]

    def __repr__(self):
        """
        `Apps` 인스턴스의 문자열 표현을 반환합니다.

        반환값
        -------
        str
            `Apps` 인스턴스의 문자열 표현.
        """
        return "{}({})".format(
            getattr(self.__class__, "_name", self.__class__.__name__), repr(list(self))
        )

    def __getitem__(self, item):
        """
        지정된 인덱스의 앱을 반환합니다.

        매개변수
        ----------
        item : int
            앱의 인덱스.

        반환값
        -------
        App
            지정된 인덱스의 앱.
        """
        return self._apps[item]

    def __len__(self):
        """
        컬렉션의 앱 수를 반환합니다.

        반환값
        -------
        int
            앱의 수.
        """
        return len(self._apps)

    @property
    def count(self):
        """
        컬렉션의 앱 수를 반환합니다.

        .. versionadded:: 0.9.0

        반환값
        -------
        int
            앱의 수.
        """
        return len(self)

    def cleanup(self):
        """
        [메서드 설명 필요]

        [메서드가 수행하는 작업과 부작용에 대한 설명]
        """


    def __iter__(self):
        """
        앱 컬렉션에 대한 반복을 허용합니다.

        반환값
        ------
        App
            컬렉션의 다음 앱.
        """
        for app in self._apps:
            yield app


