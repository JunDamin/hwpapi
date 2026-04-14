"""
Visual smoke test for hwpapi — captures screenshots of HWP window
after various operations to provide visual artifacts for debugging.

This is NOT a regression test. It only captures screenshots;
it does not compare them to reference images.

Run as:
    python tests/smoke_visual.py

Requires HWP + pyautogui + pygetwindow. Optional, not part of pytest
suite by default (use `pytest tests/smoke_visual.py` if you want).

Artifacts are saved to tests/artifacts/<timestamp>/.
"""
from __future__ import annotations
import os
import sys
import time
import datetime
import tempfile

# Skip if dependencies missing
try:
    import pyautogui
    import pygetwindow as gw
except ImportError:
    print("SKIP: pyautogui/pygetwindow not installed. Run: pip install pyautogui")
    sys.exit(0)

try:
    from hwpapi.core import App
except ImportError as e:
    print(f"SKIP: hwpapi import failed: {e}")
    sys.exit(0)


ARTIFACTS_DIR = os.path.join(
    os.path.dirname(__file__), 'artifacts',
    datetime.datetime.now().strftime('%Y%m%d_%H%M%S'),
)


def screenshot(label: str, region=None):
    """Capture screenshot; save with label, return path."""
    os.makedirs(ARTIFACTS_DIR, exist_ok=True)
    path = os.path.join(ARTIFACTS_DIR, f"{label}.png")
    try:
        img = pyautogui.screenshot(region=region)
        img.save(path)
        print(f"  📷 {path}")
        return path
    except Exception as e:
        print(f"  ⚠ screenshot failed: {e}")
        return None


def find_hwp_window():
    """Try to find HWP window and return its geometry tuple."""
    for title_pattern in ['한글', 'Hangul', 'HWP', '빈 문서']:
        wins = gw.getWindowsWithTitle(title_pattern)
        if wins:
            w = wins[0]
            try:
                if w.width > 100 and w.height > 100:
                    return (w.left, w.top, w.width, w.height)
            except Exception:
                continue
    return None


def test_startup(app):
    """Capture HWP window at startup."""
    print("\n[1] Startup")
    time.sleep(0.5)
    region = find_hwp_window()
    if region:
        screenshot("01_startup_hwp", region=region)
    else:
        screenshot("01_startup_fullscreen")


def test_text_insert(app):
    """Insert text and capture."""
    print("\n[2] Insert text")
    app.insert_text("hwpapi 시각 스모크 테스트\n\n")
    app.insert_text("두 번째 줄입니다.\n")
    app.insert_text("세 번째 줄 — 다양한 문자: abc 123 !@# 한글/English\n")
    time.sleep(0.3)
    region = find_hwp_window()
    screenshot("02_after_text_insert", region=region)


def test_bold_red(app):
    """Apply bold + red to first line."""
    print("\n[3] Apply bold + red title")
    app.move.top_of_file()
    app.select_text()  # current line
    app.set_charshape(bold=True, height=1800, text_color="#FF0000")
    time.sleep(0.3)
    region = find_hwp_window()
    screenshot("03_after_bold_red", region=region)


def test_save(app, path):
    """Save file and capture state."""
    print("\n[4] Save file")
    result = app.save(path)
    print(f"  save() → {result}")
    if os.path.isfile(path):
        size = os.path.getsize(path)
        print(f"  ✅ File saved: {size} bytes")
    else:
        print(f"  ⚠ file not saved at {path}")
    time.sleep(0.3)
    region = find_hwp_window()
    screenshot("04_after_save", region=region)


def test_find_replace(app):
    """Find and replace, capture result."""
    print("\n[5] Find + Replace")
    app.move.top_of_file()
    count = app.replace_all("hwpapi", "HWPAPI")
    print(f"  replace_all → {count} replacements")
    time.sleep(0.3)
    region = find_hwp_window()
    screenshot("05_after_replace", region=region)


def test_introspection(app):
    """Log action/pset metadata (no visual change)."""
    print("\n[6] Introspection")
    total = len(app.actions.list_actions())
    with_pset = len(app.actions.list_actions(with_pset_only=True))
    cs_class = app.actions.get_pset_class('CharShape')
    print(f"  Total actions: {total}")
    print(f"  With ParameterSet: {with_pset}")
    print(f"  CharShape class: {cs_class.__name__} ({len(cs_class._property_registry)} props)")


def test_reopen(app, path):
    """Close and reopen; capture."""
    if not os.path.isfile(path):
        print("\n[7] Skip reopen — file not saved")
        return

    print(f"\n[7] Close + Reopen {os.path.basename(path)}")
    app.close()
    time.sleep(0.5)
    screenshot("07_after_close")

    app.open(path)
    time.sleep(0.5)
    region = find_hwp_window()
    screenshot("08_after_reopen", region=region)


def main():
    print(f"📁 Artifacts: {ARTIFACTS_DIR}")

    app = App(is_visible=True)
    print(f"✅ Connected to HWP")

    tmpdir = tempfile.mkdtemp(prefix='hwpapi_smoke_')
    test_path = os.path.join(tmpdir, 'smoke_test.hwp')

    test_startup(app)
    test_text_insert(app)
    test_bold_red(app)
    test_save(app, test_path)
    test_find_replace(app)
    test_introspection(app)
    test_reopen(app, test_path)

    print(f"\n✅ Smoke test complete. Artifacts in:")
    print(f"   {ARTIFACTS_DIR}")
    if os.path.isfile(test_path):
        print(f"   Test .hwp: {test_path}")


if __name__ == "__main__":
    main()
