from __future__ import annotations

import os
import sys


def app_root_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def default_master_dir() -> str:
    return os.path.join(app_root_dir(), "master")


def asset_path(*parts: str) -> str:
    return os.path.join(app_root_dir(), "assets", *parts)
