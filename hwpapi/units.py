"""
Unit conversion helpers — ``hwpapi.units``.

HWP 내부 단위는 HWPUNIT (283 HWPUNIT = 1mm, 100 HWPUNIT = 1pt). 이 모듈은
친숙한 단위(mm/cm/in/pt) ↔ HWPUNIT 변환을 제공.

v0.0.23+ — 기존 ``app.mm_to_hwpunit()`` 등 instance method 대신 권장.
이 module function 들은 App 인스턴스 없이도 사용 가능.

Usage::

    from hwpapi import units as U

    hwp = U.mm(210)                 # 210 mm → HWPUNIT
    hwp = U.cm(21)                  # 21 cm → HWPUNIT
    hwp = U.inch(8.27)              # 8.27 in → HWPUNIT
    hwp = U.pt(12)                  # 12 pt → HWPUNIT

    mm_val = U.to_mm(hwp)           # HWPUNIT → mm (float)
    pt_val = U.to_pt(hwp)           # HWPUNIT → pt (float)
    in_val = U.to_inch(hwp)         # HWPUNIT → inch (float)

    # 문자열 입력도 지원
    hwp = U.parse("210mm")          # or "21cm", "8.27in", "12pt"
"""
from __future__ import annotations

from hwpapi.functions import from_hwpunit, to_hwpunit

# Re-export lower-level helpers
__all__ = [
    "mm", "cm", "inch", "pt",
    "to_mm", "to_cm", "to_inch", "to_pt",
    "parse", "from_hwpunit", "to_hwpunit",
]


def mm(value: float) -> int:
    """mm → HWPUNIT (283 HWPUNIT = 1mm)."""
    return to_hwpunit(value, default_unit="mm")


def cm(value: float) -> int:
    """cm → HWPUNIT (2834 HWPUNIT ≈ 1cm)."""
    return to_hwpunit(value, default_unit="cm")


def inch(value: float) -> int:
    """inch → HWPUNIT (7200 HWPUNIT = 1in)."""
    return to_hwpunit(value, default_unit="in")


def pt(value: float) -> int:
    """point → HWPUNIT (100 HWPUNIT = 1pt)."""
    return to_hwpunit(value, default_unit="pt")


def to_mm(hwpunit_value: int) -> float:
    """HWPUNIT → mm."""
    return from_hwpunit(hwpunit_value, target_unit="mm")


def to_cm(hwpunit_value: int) -> float:
    """HWPUNIT → cm."""
    return from_hwpunit(hwpunit_value, target_unit="cm")


def to_inch(hwpunit_value: int) -> float:
    """HWPUNIT → inch."""
    return from_hwpunit(hwpunit_value, target_unit="in")


def to_pt(hwpunit_value: int) -> float:
    """HWPUNIT → point."""
    return from_hwpunit(hwpunit_value, target_unit="pt")


def parse(value) -> int:
    """
    문자열/숫자 → HWPUNIT.

    Examples
    --------
    >>> parse("210mm")   # 283 * 210 ≈ 59544
    59430
    >>> parse("21cm")
    59430
    >>> parse("8.27in")  # 7200 * 8.27
    59544
    >>> parse("12pt")    # 100 * 12
    1200
    >>> parse(210)       # int → assumed mm
    59430
    """
    return to_hwpunit(value, default_unit="mm")
