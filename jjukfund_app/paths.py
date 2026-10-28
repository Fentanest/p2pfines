from __future__ import annotations

import os
import sys
from pathlib import Path


APP_NAME = "Jjukfund"
PROJECT_ROOT = Path(__file__).resolve().parents[1]
IS_FROZEN = bool(getattr(sys, "frozen", False))


def _resource_root() -> Path:
    if IS_FROZEN:
        bundled_root = getattr(sys, "_MEIPASS", None)
        if bundled_root:
            return Path(bundled_root).resolve()
        return Path(sys.executable).resolve().parent
    return PROJECT_ROOT


def _data_root() -> Path:
    override = os.environ.get("JJUKFUND_DATA_DIR", "").strip()
    if override:
        return Path(override).expanduser().resolve()
    if not IS_FROZEN:
        return PROJECT_ROOT
    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if local_app_data:
        return (Path(local_app_data) / APP_NAME).resolve()
    return (Path(sys.executable).resolve().parent / "data").resolve()


RESOURCE_ROOT = _resource_root()
DATA_ROOT = _data_root()
