r"""Настройки запуска: kk_config.json рядом с exe (общие) и в %LOCALAPPDATA%\KK (личные правки).

Ключи:
  macro1      путь к .swp макроса 1 (баллоны и карта). Рядом с ним должен лежать Настройки_ОК.xlsx.
  macro2      путь к .swp макроса 2 (листы ОК).
  module1/2   модуль с процедурой Main; пусто — поиск автоматически (как в kc_start.py).
  procedure1/2  имя процедуры, по умолчанию Main.
"""
from __future__ import annotations

import json
import logging

from .paths import app_dir, user_data_dir

log = logging.getLogger(__name__)
NAME = 'kk_config.json'
DEFAULTS = {'macro1': '', 'macro2': '', 'module1': '', 'module2': '', 'procedure1': 'Main',
            'procedure2': 'Main', 'recent': []}


def _read(path):
    try:
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as exc:
        log.warning('%s не прочитан: %s', path, exc)
        return {}


def load():
    data = dict(DEFAULTS)
    for source in (app_dir() / NAME, user_data_dir() / NAME):
        for key, value in _read(source).items():
            if key in DEFAULTS and isinstance(value, type(DEFAULTS[key])):
                data[key] = value
    return data


def save_user(data):
    path = user_data_dir() / NAME
    clean = {k: data[k] for k in DEFAULTS if k in data}
    tmp = path.with_suffix('.tmp')
    tmp.write_text(json.dumps(clean, ensure_ascii=False, indent=2), encoding='utf-8')
    tmp.replace(path)
