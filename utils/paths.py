from pathlib import Path
import os
import sys

APP_NAME = "Suzuki Parts Manager"

def appdata_root() -> Path:
    local = os.getenv("LOCALAPPDATA")
    if local:
        root = Path(local) / APP_NAME
    else:
        root = Path.home() / f".{APP_NAME}"
    root.mkdir(parents=True, exist_ok=True)
    return root

def output_root() -> Path:
    p = appdata_root() / "output"
    p.mkdir(parents=True, exist_ok=True)
    return p

def get_lock_file() -> Path:
    return appdata_root() / "app.lock"

def resource_path(relative: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return Path(__file__).resolve().parent.parent / relative