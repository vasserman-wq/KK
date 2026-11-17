"""Настройки_ОК.xlsx: встроенные значения по умолчанию + общий файл рядом с макросом 1.

Макросы читают файл рядом с .swp (KC28_Load: GetCurrentMacroPathName). Этот файл — общее переопределение,
править его может любой пользователь из программы. Встроенная копия (kk/resources/Настройки_ОК.xlsx)
используется, если общего файла нет, и для команды «Вернуть стандартное значение».
Запись: блокировка (файл .lock), резервная копия в «Резерв настроек», замена через временный файл.
Структура (схема OK_SETTINGS_V2): лист «Версия» B1; на листах параметров A — название,
B — значение, F — служебный ключ (как читает KC28_Load). Ключи и заголовки программа не меняет.
"""
from __future__ import annotations

import datetime as dt
import getpass
import os
import shutil
import socket
import tempfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .paths import resource

FILE_NAME = 'Настройки_ОК.xlsx'
SCHEMAS = ('OK_SETTINGS_V1', 'OK_SETTINGS_V2')
PARAM_SHEETS = ('Файлы и папки', 'Баллоны', 'Нормы контроля', 'Шифры', 'Подписи', 'Сервис и логи',
                'Таблицы ОК', 'Таблицы КЭ', 'Рамки ОК', 'Рамки КЭ', 'Свои параметры')
TABLE_SHEETS = ('Выбор инструментов', 'Добавочные средства', 'Шаблоны размеров', 'Ссылки на пункты')
LOCK_MAX_AGE = dt.timedelta(minutes=30)


class SettingsError(Exception):
    pass


def builtin_path():
    return resource(FILE_NAME)


def shared_path(macro1):
    return Path(macro1).parent / FILE_NAME if macro1 else None


def ensure_shared(macro1):
    """Если рядом с макросом нет файла настроек — создать из встроенной копии. Возвращает (путь, создан)."""
    target = shared_path(macro1)
    if target is None:
        raise SettingsError('Не указан путь к макросу 1.')
    if target.exists():
        return target, False
    if not builtin_path().exists():
        raise SettingsError('Встроенный файл настроек отсутствует в сборке (kk/resources).')
    shutil.copy2(builtin_path(), target)
    return target, True


@dataclass
class Param:
    sheet: str
    row: int
    title: str
    value: object
    unit: str
    hint: str
    allowed: str
    key: str


def read_params(path):
    wb = openpyxl.load_workbook(str(path), data_only=False)
    try:
        schema = str(wb['Версия']['B1'].value or '') if 'Версия' in wb.sheetnames else ''
        if schema not in SCHEMAS:
            raise SettingsError(f'Неподдерживаемая версия файла настроек: «{schema}».')
        out = {}
        for name in PARAM_SHEETS:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            rows = []
            for r in range(2, min(ws.max_row, 5000) + 1):
                key = ws.cell(r, 6).value
                title = ws.cell(r, 1).value
                if key in (None, '') and title in (None, ''):
                    continue
                rows.append(Param(name, r, str(title or key), ws.cell(r, 2).value, str(ws.cell(r, 3).value or ''),
                                  str(ws.cell(r, 4).value or ''), str(ws.cell(r, 5).value or ''), str(key or '')))
            out[name] = rows
        tables = {}
        for name in TABLE_SHEETS:
            if name in wb.sheetnames:
                ws = wb[name]
                tables[name] = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                                for r in range(1, min(ws.max_row, 1000) + 1)]
        return schema, out, tables
    finally:
        wb.close()


class Lock:
    """Блокировка общего файла на время записи: файл <имя>.lock с именем пользователя."""

    def __init__(self, path):
        self.path = Path(str(path) + '.lock')
        self.fd = None

    def __enter__(self):
        if self.path.exists():
            age = dt.datetime.now() - dt.datetime.fromtimestamp(self.path.stat().st_mtime)
            if age < LOCK_MAX_AGE:
                who = self.path.read_text(encoding='utf-8', errors='replace').strip()
                raise SettingsError(f'Настройки сейчас сохраняет другой пользователь ({who}). Повторите позже.')
            self.path.unlink()
        self.fd = os.open(str(self.path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(self.fd, f'{getpass.getuser()}@{socket.gethostname()} {dt.datetime.now():%Y-%m-%d %H:%M}'.encode('utf-8'))
        return self

    def __exit__(self, *exc):
        if self.fd is not None:
            os.close(self.fd)
        try:
            self.path.unlink()
        except OSError:
            pass


def write_values(path, changes, expected_mtime=None):
    """changes: {(лист, строка): значение} для колонки B или {(лист, строка, колонка): значение}."""
    path = Path(path)
    with Lock(path):
        if expected_mtime is not None and abs(path.stat().st_mtime - expected_mtime) > 1:
            raise SettingsError('Файл настроек изменил другой пользователь, пока окно было открыто. '
                                'Откройте настройки заново — ваши правки не записаны.')
        backup_dir = path.parent / 'Резерв настроек'
        backup_dir.mkdir(exist_ok=True)
        backup = backup_dir / f'{path.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{path.suffix}'
        shutil.copy2(path, backup)
        wb = openpyxl.load_workbook(str(path))
        for key, value in changes.items():
            sheet, row = key[0], key[1]
            col = key[2] if len(key) > 2 else 2
            wb[sheet].cell(row, col).value = value
        fd, tmp = tempfile.mkstemp(prefix='.kk_', suffix='.xlsx', dir=str(path.parent))
        os.close(fd)
        try:
            wb.save(tmp)
            os.replace(tmp, path)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
        return backup
