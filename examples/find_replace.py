#!/usr/bin/env python3
"""
Example: Find and Replace using HParameterSet

This example demonstrates how to use the new HParameterSet support to:
1. Set HFindReplace parameters using dotted keys
2. Execute the FindReplace action using resolve_action_args
3. Use temp_edit_hparam for non-destructive parameter changes

Requirements:
- HWP application running
- hwpapi package with HParameterSet support
"""

import sys
import os

# Add hwpapi to path if running as standalone script
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

try:
    from hwpapi.core import App
    from hwpapi.parametersets import (
        make_backend, resolve_action_args, temp_edit_hparam,
        snapshot_hparam, apply_dict_hparam
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
        
        # Get HParameterSet - this should auto-detect as HParamBackend
        hparam_root = app.api.HParameterSet
        backend = make_backend(hparam_root)
        print(f"Backend type: {type(backend).__name__}")
        
        # Example 1: Direct backend usage with dotted keys
        print("\n=== Example 1: Direct Backend Usage ===")
        
        # Set find/replace parameters using dotted keys
        backend.set("HFindReplace.FindString", "example")
        backend.set("HFindReplace.ReplaceString", "sample")
        backend.set("HFindReplace.IgnoreCase", True)
        backend.set("HFindReplace.WholeWordOnly", False)
        
        # Read back the values
        find_str = backend.get("HFindReplace.FindString")
        replace_str = backend.get("HFindReplace.ReplaceString")
        ignore_case = backend.get("HFindReplace.IgnoreCase")
        
        print(f"Find: '{find_str}' -> Replace: '{replace_str}'")
        print(f"Ignore case: {ignore_case}")
        
        # Example 2: Using resolve_action_args for HAction calls
        print("\n=== Example 2: Action Execution ===")
        
        # Resolve action arguments (tries .HSet first, falls back to node)
        action_name, arg = resolve_action_args(
            app, "FindReplace", hparam_root.HFindReplace
        )
        print(f"Action: {action_name}, Arg type: {type(arg).__name__}")
        
        # Get default parameters (optional)
        try:
            app.HAction.GetDefault(action_name, arg)
            print("GetDefault succeeded")
        except Exception as e:
            print(f"GetDefault failed: {e}")
        
        # Execute the action (would actually perform find/replace)
        # Commented out to avoid modifying documents during example
        # try:
        #     result = app.HAction.Execute(action_name, arg)
        #     print(f"Execute result: {result}")
        # except Exception as e:
        #     print(f"Execute failed: {e}")
        
        # Example 3: Non-destructive editing with temp_edit_hparam
        print("\n=== Example 3: Non-destructive Editing ===")
        
        # Take a snapshot of current state
        original_snapshot = snapshot_hparam(hparam_root, "HFindReplace")
        print(f"Original state: {original_snapshot}")
        
        # Use context manager for temporary changes
        temp_changes = {
            "FindString": "temporary",
            "ReplaceString": "changed",
            "IgnoreCase": False
        }
        
        with temp_edit_hparam(hparam_root, "HFindReplace", temp_changes):
            # Inside context: values are temporarily changed
            current_find = backend.get("HFindReplace.FindString")
            print(f"Inside context - Find: '{current_find}'")
            
            # Here you could execute actions with the temporary parameters
            # action_name, arg = resolve_action_args(app, "FindReplace", hparam_root.HFindReplace)
            # app.HAction.Execute(action_name, arg)
        
        # Outside context: values are restored
        restored_find = backend.get("HFindReplace.FindString")
        print(f"After context - Find: '{restored_find}' (should be restored)")
        
        # Verify restoration
        final_snapshot = snapshot_hparam(hparam_root, "HFindReplace")
        print(f"Final state matches original: {original_snapshot == final_snapshot}")
        
        print("\n=== Example Complete ===")
        
    except Exception as e:
        print(f"Error in example: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
