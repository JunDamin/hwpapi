"""
Quickstart example for hwpapi.

Shows the most common operations: file I/O, text, formatting, navigation.
Run this as a first test to verify hwpapi works with your HWP installation.
"""
from hwpapi.core import App


def main():
    # Connect to HWP (launches if not running)
    app = App(is_visible=True)
    print(f"Connected: {app}")

    # ── 1. Text insertion ──────────────────────────────────────────────
    app.insert_text("hwpapi 테스트 문서\n\n")
    app.insert_text("이 문서는 hwpapi로 자동 생성되었습니다.\n")

    # ── 2. Formatting ──────────────────────────────────────────────────
    # Select the first line
    app.move.top_of_file()
    app.select_text()  # current line

    # Apply title formatting
    app.set_charshape(bold=True, height=1800)  # 18pt bold

    # ── 3. Navigate and insert more ────────────────────────────────────
    app.move.bottom()
    app.insert_text("\n")

    for i in range(3):
        app.insert_text(f"• 항목 {i + 1}\n")

    # ── 4. Find / Replace ──────────────────────────────────────────────
    app.move.top_of_file()
    count = app.replace_all("항목", "list item")
    print(f"Replaced {count} occurrences")

    # ── 5. Action introspection (no COM calls) ─────────────────────────
    print(f"\nAvailable actions: {len(app.actions.list_actions())}")
    print(f"Actions with ParameterSet: {len(app.actions.list_actions(with_pset_only=True))}")
    print(f"CharShape has {len(app.actions.get_pset_class('CharShape')._property_registry)} properties")

    # ── 6. Raw COM access (escape hatch) ───────────────────────────────
    print(f"\nRaw HWP COM: {app.api.CLSID}")

    print("\n✅ Quickstart complete.")


if __name__ == "__main__":
    main()
