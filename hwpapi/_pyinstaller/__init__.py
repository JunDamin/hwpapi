"""PyInstaller hook support for hwpapi."""
import os

def get_hook_dirs():
    """Return hook directory for PyInstaller to discover."""
    return [os.path.dirname(__file__)]
