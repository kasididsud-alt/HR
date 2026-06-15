from __future__ import annotations

import os
import sys


def app_root_dir() -> str:
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        parts = exe_dir.split(os.sep)
        for i, part in enumerate(parts):
            if part.endswith(".app"):
                return os.sep.join(parts[:i]) or os.sep
        return exe_dir
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def default_master_dir() -> str:
    return os.path.join(app_root_dir(), "master")


def asset_path(*parts: str) -> str:
    return os.path.join(app_root_dir(), "assets", *parts)
