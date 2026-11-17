"""Пути программы (как в passport_macro): рядом с exe, внутри сборки, данные пользователя."""
from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = 'KK'


def app_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def bundle_dir() -> Path:
    return Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent.parent))


def resource(name) -> Path:
    return bundle_dir() / 'kk' / 'resources' / name


def user_data_dir() -> Path:
    base = os.environ.get('LOCALAPPDATA') or os.environ.get('XDG_STATE_HOME') or str(Path.home() / '.local' / 'state')
    path = Path(base) / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path
