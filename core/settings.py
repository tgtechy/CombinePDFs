# core/settings.py

import os
print("CWD:", os.getcwd())
print("Loaded settings.py from:", __file__)

import json
from pathlib import Path

def load_settings(config_path: Path) -> dict:
    """Load saved settings from a JSON config file."""
    try:
        if config_path.exists():
            with open(config_path, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_settings(config_path: Path, settings: dict) -> None:
    """Save settings dictionary to a JSON config file."""
    try:
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception:
        # Silently ignore save errors (same behavior as load)
        pass