"""Test hwpapi.units module — HWPUNIT conversion helpers."""
import pytest
from hwpapi import units as U


def test_mm_to_hwpunit():
    """1mm = 283 HWPUNIT (approximately)."""
    assert U.mm(1) == 283
    assert U.mm(210) == 59430   # A4 width


def test_cm_to_hwpunit():
    # 1cm = 10mm = 10 * 283 HWPUNIT = 2830
    assert U.cm(1) == 2830
    assert U.cm(21) == U.mm(210)  # 21cm == 210mm


def test_inch_to_hwpunit():
    assert U.inch(1) == 7200   # 1 inch = 7200 HWPUNIT
    assert U.inch(8.27) == 59544


def test_pt_to_hwpunit():
    assert U.pt(1) == 100      # 1pt = 100 HWPUNIT
    assert U.pt(12) == 1200
    assert U.pt(24) == 2400


def test_to_mm_roundtrip():
    assert U.to_mm(U.mm(210)) == pytest.approx(210.0, abs=0.01)
    assert U.to_mm(U.mm(50)) == pytest.approx(50.0, abs=0.01)


def test_to_pt_roundtrip():
    assert U.to_pt(U.pt(12)) == pytest.approx(12.0, abs=0.01)


def test_to_cm_roundtrip():
    assert U.to_cm(U.cm(21)) == pytest.approx(21.0, abs=0.01)


def test_parse_string_with_unit():
    assert U.parse("210mm") == 59430
    assert U.parse("21cm") == 59430
    assert U.parse("8.27in") == 59544
    assert U.parse("12pt") == 1200


def test_parse_bare_number_defaults_to_mm():
    # Bare numbers default to mm per the V1 plan
    assert U.parse(210) == U.mm(210)


def test_module_public_api():
    """All expected functions exported."""
    expected = {"mm", "cm", "inch", "pt", "to_mm", "to_cm", "to_inch", "to_pt",
                 "parse", "from_hwpunit", "to_hwpunit"}
    assert expected.issubset(set(U.__all__))
