r"""Журнал программы: %LOCALAPPDATA%\KK\kk.log, 3 архива по 1 МБ."""
from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler

from .paths import user_data_dir

FORMAT = '%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s'


def log_path():
    return user_data_dir() / 'kk.log'


def setup_logging(level=logging.INFO):
    root = logging.getLogger()
    if root.handlers:
        return
    root.setLevel(level)
    try:
        handler = RotatingFileHandler(log_path(), maxBytes=1_000_000, backupCount=3, encoding='utf-8')
    except OSError:
        handler = logging.StreamHandler() if sys.stderr else logging.NullHandler()
    handler.setFormatter(logging.Formatter(FORMAT))
    root.addHandler(handler)

    def hook(kind, value, tb):
        logging.getLogger('kk').critical('Необработанная ошибка', exc_info=(kind, value, tb))
    sys.excepthook = hook
