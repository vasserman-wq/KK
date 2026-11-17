"""Окна выбора файла для эмулятора Excel (вместо GetSaveAsFilename, GetOpenFilename, FileDialog).

Вызываются из потока COM-объекта, поэтому используют win32gui (не tkinter). Отмена -> ''.
"""
from __future__ import annotations

import logging
import os

log = logging.getLogger(__name__)


def _filter(text):
    """'Excel (*.xlsm),*.xlsm' -> фильтр Windows 'Excel (*.xlsm)\0*.xlsm\0'."""
    parts = [p.strip() for p in str(text or '').split(',') if p.strip()]
    if len(parts) < 2:
        return 'Все файлы\0*.*\0'
    pairs = [(parts[i], parts[i + 1]) for i in range(0, len(parts) - 1, 2)]
    return ''.join(f'{a}\0{b}\0' for a, b in pairs)


class WinDialogs:
    def __init__(self, owner_hwnd=0):
        self.owner = owner_hwnd

    def open_file(self, title, initial, filters):
        import win32gui
        flt = ''.join(_filter(f.replace('|', ',')) for f in filters if f) or 'Все файлы\0*.*\0'
        folder = initial if os.path.isdir(initial or '') else os.path.dirname(initial or '')
        try:
            name, _, _ = win32gui.GetOpenFileNameW(hwndOwner=self.owner, Title=title or 'Выбор файла',
                                                   Filter=flt, InitialDir=folder or None, Flags=0x1000 | 0x800)
            return name
        except Exception as exc:  # отмена пользователем тоже приходит исключением
            log.info('Выбор файла отменён: %s', exc)
            return ''

    def save_file(self, title, initial, file_filter):
        import win32gui
        folder, name = os.path.split(initial or '')
        try:
            path, _, _ = win32gui.GetSaveFileNameW(hwndOwner=self.owner, Title=title or 'Сохранить',
                                                   Filter=_filter(file_filter), File=name or None,
                                                   InitialDir=folder or None, Flags=0x2 | 0x800)
            return path
        except Exception as exc:
            log.info('Сохранение отменено: %s', exc)
            return ''
