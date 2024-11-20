# src/utils/utils.py

import os
import sys

def resource_path(relative_path):
    """Get the absolute path to a resource, works for development and PyInstaller."""
    try:
        # When using PyInstaller
        base_path = sys._MEIPASS
    except AttributeError:
        # During development
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)