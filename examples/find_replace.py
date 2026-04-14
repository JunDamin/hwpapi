"""
Example: Find and Replace with hwpapi.

Demonstrates the modern hwpapi API for find/replace operations:

1. Simple find_text() / replace_all() via App convenience methods
2. Detailed FindReplace ParameterSet access (regex, case, wildcards)
3. Nested formatting constraints (find only bold text, etc.)
4. Snapshot + clone pattern for temporary parameter changes

Requires a running HWP instance (or accessible COM).
"""
from hwpapi.core import App


def example_1_simple(app):
    """Simple find_text / replace_all via App methods."""
    print("\n--- Example 1: Simple find/replace ---")

    # Insert some sample text
    app.insert_text("Hello world. Hello universe. Hello python.\n")
    app.move.top_of_file()

    # Find and jump to first match
    found = app.find_text("Hello")
    print(f"  find_text('Hello') → {found}")

    # Replace all occurrences
    count = app.replace_all("Hello", "Hi")
    print(f"  replace_all('Hello' → 'Hi'): {count} replacements")


def example_2_detailed_pset(app):
    """Detailed FindReplace ParameterSet manipulation."""
    print("\n--- Example 2: Detailed ParameterSet ---")

    app.insert_text("Old text. Some other text. Old text again.\n")
    app.move.top_of_file()

    action = app.actions.AllReplace
    ps = action.pset  # FindReplace ParameterSet

    ps.FindString = "Old"
    ps.ReplaceString = "NEW"
    ps.IgnoreCase = 0            # case-sensitive
    ps.WholeWordOnly = 0
    ps.UseWildCards = 0

    action.run()
    print(f"  Replaced 'Old' → 'NEW' (case-sensitive)")


def example_3_nested_formatting(app):
    """Find only text with specific formatting (bold)."""
    print("\n--- Example 3: Nested formatting constraint ---")

    action = app.actions.FindReplace
    ps = action.pset

    ps.FindString = "important"

    # Auto-creating nested ParameterSet — access creates on first touch
    ps.find_char_shape.Bold = True  # only find BOLD occurrences
    print(f"  FindCharShape.Bold set: {ps.find_char_shape.Bold}")

    # Note: action.run() would actually find the text in HWP; skipped here
    print("  (Would find only bold 'important' if actually run)")


def example_4_snapshot_and_restore(app):
    """Save current FindReplace state, modify, restore via clone."""
    print("\n--- Example 4: Snapshot / restore pattern ---")

    action = app.actions.FindReplace
    ps = action.pset

    # Snapshot via native pset.Clone()
    original = ps.clone()
    print(f"  Snapshot: FindString='{original.FindString}'")

    # Temporary change
    ps.FindString = "temporary"
    ps.IgnoreCase = 1
    print(f"  Changed: FindString='{ps.FindString}', IgnoreCase={ps.IgnoreCase}")

    # Check they differ
    print(f"  is_equivalent(original): {ps.is_equivalent(original)}")

    # Restore
    ps.update_from(original)
    print(f"  Restored: FindString='{ps.FindString}'")


def example_5_introspection(app):
    """Introspect find/replace action without creating it."""
    print("\n--- Example 5: Introspection ---")

    for name in ['FindReplace', 'AllReplace', 'ForwardFind', 'BackwardFind']:
        if name in app.actions:
            cls = app.actions.get_pset_class(name)
            desc = app.actions.get_description(name)
            print(f"  {name} → {cls.__name__ if cls else 'None'}: {desc}")


def main():
    app = App(is_visible=True)
    example_1_simple(app)
    example_2_detailed_pset(app)
    example_3_nested_formatting(app)
    example_4_snapshot_and_restore(app)
    example_5_introspection(app)
    print("\n✅ All examples completed!")


if __name__ == "__main__":
    main()
