"""
Example: Character formatting with hwpapi.

Demonstrates the modern hwpapi API for manipulating character shapes:

1. Simple bold toggle via App.set_charshape()
2. Per-attribute modification via ParameterSet access
3. Snapshot / restore pattern (clone + update_from)
4. Introspection without COM calls

Requires a running HWP instance (or accessible COM).
"""
from hwpapi.core import App


def example_1_simple_bold(app):
    """Simplest way to toggle bold formatting."""
    print("\n--- Example 1: Simple bold toggle ---")

    app.insert_text("Hello, this is bold test.\n")
    app.move.top_of_file()
    app.select_text()  # select current line

    app.set_charshape(bold=True)
    print("  Applied bold formatting.")


def example_2_per_attribute(app):
    """Modify multiple attributes via ParameterSet."""
    print("\n--- Example 2: Per-attribute modification ---")

    app.insert_text("This will be red, italic, 14pt.\n")
    app.move.top_of_file()
    app.select_text()

    action = app.actions.CharShape
    ps = action.pset

    # snake_case or PascalCase both work
    ps.italic = True
    ps.Bold = False
    ps.Height = 1400            # 14pt (HWPUNIT: 100/pt)
    ps.TextColor = "#FF0000"    # hex → BBGGRR auto-conversion

    action.run()
    print(f"  Applied: italic={ps.italic}, Height={ps.Height}, TextColor={ps.TextColor}")


def example_3_snapshot_restore(app):
    """Save current state, apply changes, restore via clone."""
    print("\n--- Example 3: Snapshot / restore pattern ---")

    app.insert_text("Snapshot test.\n")
    app.move.top_of_file()
    app.select_text()

    action = app.actions.CharShape
    ps = action.pset

    # Snapshot current state (native pset.Clone())
    original = ps.clone()
    print(f"  Snapshot: Bold={original.Bold}, Italic={original.Italic}")

    # Apply temporary changes
    ps.Bold = True
    ps.Italic = True
    ps.TextColor = "#00FF00"
    action.run()
    print("  Applied temporary formatting.")

    # Restore from snapshot
    ps.update_from(original)
    action.run()
    print("  Restored original formatting.")


def example_4_introspection(app):
    """Introspect actions without any COM calls."""
    print("\n--- Example 4: Introspection (no COM) ---")

    # Get ParameterSet class without creating _Action
    cls = app.actions.get_pset_class('CharShape')
    desc = app.actions.get_description('CharShape')
    print(f"  CharShape → {cls.__name__}: {desc}")

    # List attributes
    print(f"  Has {len(cls._property_registry)} properties")
    print(f"  First 5: {list(cls._property_registry.keys())[:5]}")


def main():
    app = App(is_visible=True)
    example_1_simple_bold(app)
    example_2_per_attribute(app)
    example_3_snapshot_restore(app)
    example_4_introspection(app)
    print("\n✅ All examples completed!")


if __name__ == "__main__":
    main()
