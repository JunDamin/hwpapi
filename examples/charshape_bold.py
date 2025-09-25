#!/usr/bin/env python3
"""
Example: Character Shape Bold Toggle using HParameterSet

This example demonstrates how to use HParameterSet support to:
1. Toggle HCharShape.Bold and other character formatting
2. Use temp_edit_hparam for safe parameter manipulation
3. Apply multiple character formatting changes at once

Requirements:
- HWP application running
- hwpapi package with HParameterSet support
- A document with some text selected (optional)
"""

import sys
import os

# Add hwpapi to path if running as standalone script
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

try:
    from hwpapi.core import App
    from hwpapi.parametersets import (
        make_backend, resolve_action_args, temp_edit_hparam,
        snapshot_hparam, to_dict_hparam
    )
except ImportError as e:
    print(f"Error importing hwpapi: {e}")
    print("Make sure hwpapi is installed and HWP is running")
    sys.exit(1)


def main():
    """Main example function."""
    try:
        # Initialize HWP application
        print("Connecting to HWP...")
        app = App()
        
        # Get HParameterSet
        hparam_root = app.api.HParameterSet
        backend = make_backend(hparam_root)
        print(f"Backend type: {type(backend).__name__}")
        
        # Example 1: Inspect current character shape
        print("\n=== Example 1: Current Character Shape ===")
        
        # Get current character shape as nested dict (primitives only)
        charshape_dict = to_dict_hparam(hparam_root, "HCharShape", depth=2)
        print("Current HCharShape properties:")
        for key, value in sorted(charshape_dict.items()):
            if isinstance(value, (int, float, bool, str)):
                print(f"  {key}: {value}")
        
        # Example 2: Toggle bold formatting
        print("\n=== Example 2: Bold Toggle ===")
        
        # Get current bold state
        try:
            current_bold = backend.get("HCharShape.Bold")
            print(f"Current bold state: {current_bold}")
            
            # Toggle bold
            new_bold = not bool(current_bold) if current_bold is not None else True
            backend.set("HCharShape.Bold", new_bold)
            
            # Verify the change
            updated_bold = backend.get("HCharShape.Bold")
            print(f"Updated bold state: {updated_bold}")
            
        except Exception as e:
            print(f"Bold toggle failed: {e}")
        
        # Example 3: Multiple character formatting changes with temp_edit_hparam
        print("\n=== Example 3: Temporary Formatting Changes ===")
        
        # Take snapshot of current character shape
        original_snapshot = snapshot_hparam(hparam_root, "HCharShape")
        print(f"Original character shape has {len(original_snapshot)} properties")
        
        # Define multiple formatting changes
        formatting_changes = {
            "Bold": True,
            "Italic": True,
            # "FontSize": 14,  # Uncomment if FontSize property exists
            # "Underline": 1,  # Uncomment if Underline property exists
        }
        
        print(f"Applying temporary changes: {formatting_changes}")
        
        with temp_edit_hparam(hparam_root, "HCharShape", formatting_changes):
            # Inside context: formatting is temporarily applied
            print("Inside context - temporary formatting applied:")
            for key in formatting_changes:
                try:
                    value = backend.get(f"HCharShape.{key}")
                    print(f"  {key}: {value}")
                except Exception as e:
                    print(f"  {key}: Error reading - {e}")
            
            # Here you could apply the formatting to selected text
            # try:
            #     action_name, arg = resolve_action_args(app, "CharShape", hparam_root.HCharShape)
            #     app.HAction.Execute(action_name, arg)
            #     print("Character formatting applied to selection")
            # except Exception as e:
            #     print(f"Failed to apply formatting: {e}")
        
        # Outside context: formatting is restored
        print("After context - formatting restored:")
        for key in formatting_changes:
            try:
                value = backend.get(f"HCharShape.{key}")
                print(f"  {key}: {value}")
            except Exception as e:
                print(f"  {key}: Error reading - {e}")
        
        # Example 4: Safe property exploration
        print("\n=== Example 4: Safe Property Exploration ===")
        
        # Use delete() method to demonstrate neutralization
        print("Testing neutralization (delete) behavior:")
        try:
            # Try to neutralize some properties
            test_props = ["Bold", "Italic"]
            for prop in test_props:
                full_key = f"HCharShape.{prop}"
                try:
                    original_value = backend.get(full_key)
                    success = backend.delete(full_key)
                    neutral_value = backend.get(full_key)
                    print(f"  {prop}: {original_value} -> neutralized to {neutral_value} (success: {success})")
                    
                    # Restore original value
                    backend.set(full_key, original_value)
                except Exception as e:
                    print(f"  {prop}: Error during neutralization test - {e}")
                    
        except Exception as e:
            print(f"Neutralization test failed: {e}")
        
        # Example 5: Error handling demonstration
        print("\n=== Example 5: Error Handling ===")
        
        # Try to access non-existent property
        try:
            backend.get("HCharShape.NonExistentProperty")
        except KeyError as e:
            print(f"Expected KeyError for non-existent property: {e}")
        
        # Try to set invalid value type (should show type coercion)
        try:
            backend.set("HCharShape.Bold", "invalid")  # String instead of bool
            print("Type coercion succeeded (or no validation)")
        except TypeError as e:
            print(f"Expected TypeError for invalid type: {e}")
        
        print("\n=== Character Shape Example Complete ===")
        
    except Exception as e:
        print(f"Error in example: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
