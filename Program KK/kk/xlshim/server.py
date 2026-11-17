"""Запуск COM-объекта KK.ExcelShim в процессе программы (только Windows, pywin32).

Реестр: только HKCU (права администратора не нужны): ProgID -> CLSID, чтобы VBA мог
GetObject(, "KK.ExcelShim"). Сам объект регистрируется в Running Object Table на время работы программы.
НЕ ПРОВЕРЕНО В WINDOWS: см. docs/TESTING.md, шаг «Проверка эмулятора».
"""
from __future__ import annotations

import logging
import threading

from .com import CLSID, PROGID, NoOp, dispatch
from .app import App, Factory
from .core import Book, Sheet
from .ranges import ShimError

_CACHED = (Sheet, Book, App, Factory)

log = logging.getLogger(__name__)
_SCALARS = (str, int, float, bool)


def register_progid():
    import winreg
    for key, value in ((rf'Software\Classes\{PROGID}\CLSID', CLSID),
                       (rf'Software\Classes\CLSID\{CLSID}\ProgID', PROGID),
                       (rf'Software\Classes\CLSID\{CLSID}', 'KK Excel Shim')):
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key) as handle:
            winreg.SetValueEx(handle, '', 0, winreg.REG_SZ, value)


class ComObj:
    """Обёртка для DynamicPolicy: любое имя свойства/метода приходит в _dynamic_."""
    _public_methods_ = []

    def __init__(self, target, factory):
        self.target, self.factory = target, factory

    def _dynamic_(self, name, lcid, wFlags, args):
        import pythoncom
        from win32com.server.exception import COMException
        try:
            result = dispatch(self.target, name, wFlags, args, self.factory)
        except ShimError as exc:
            self.factory.log.append(f'Ошибка: {type(self.target).__name__}.{name}: {exc}')
            raise COMException(description=str(exc), scode=-2147352567)
        except AttributeError:
            self.factory.log.append(f'Не поддержано: {type(self.target).__name__}.{name}')
            log.warning('Не поддержано: %s.%s', type(self.target).__name__, name)
            raise COMException(description=f'Эмулятор Excel не поддерживает {name}',
                               scode=-2147352573)
        except Exception as exc:  # неожиданная ошибка программы не должна ронять SOLIDWORKS
            log.exception('Сбой эмулятора: %s', name)
            raise COMException(description=f'Сбой эмулятора Excel: {exc}', scode=-2147352567)
        if name.lower() == '_newenum':
            from win32com.server.util import NewEnum
            return NewEnum([to_com(x, self.factory) for x in result], iid=pythoncom.IID_IEnumVARIANT)
        return to_com(result, self.factory)


def to_com(value, factory):
    import pythoncom
    from win32com.server import policy, util
    if value is None:
        return pythoncom.Empty
    if isinstance(value, _SCALARS):
        return value
    if isinstance(value, list):
        if value and isinstance(value[0], list):
            # Массив значений диапазона. pywin32 создаёт SAFEARRAY с нижней границей 0, а макросы
            # читают values(1..UBound). Добавляем пустую строку и колонку с индексом 0.
            width = len(value[0])
            pad = tuple([pythoncom.Empty] * (width + 1))
            rows = [pad] + [tuple([pythoncom.Empty] + [pythoncom.Empty if v is None else v for v in row])
                            for row in value]
            return tuple(rows)
        return tuple(to_com(v, factory) for v in value)
    if isinstance(value, _CACHED):
        # VBA сравнивает объекты через Is (area.Parent Is ws): один Python-объект = один COM-объект.
        cache = factory.__dict__.setdefault('_com_cache', {})
        hit = cache.get(id(value))
        if hit is None or hit[0] is not value:
            hit = cache[id(value)] = (value, util.wrap(ComObj(value, factory), usePolicy=policy.DynamicPolicy))
        return hit[1]
    return util.wrap(ComObj(value, factory), usePolicy=policy.DynamicPolicy)


class ShimServer:
    """Поток STA с обработкой сообщений: здесь живёт объект, сюда приходят вызовы VBA."""

    def __init__(self, factory):
        self.factory = factory
        self.thread = None
        self.thread_id = None
        self.ready = threading.Event()
        self.error = None

    def start(self, timeout=10):
        if self.thread and self.thread.is_alive():
            return
        self.ready.clear()
        self.thread = threading.Thread(target=self._run, name='KK-ExcelShim', daemon=True)
        self.thread.start()
        if not self.ready.wait(timeout):
            raise RuntimeError('Эмулятор Excel не запустился за отведённое время.')
        if self.error:
            raise RuntimeError(f'Эмулятор Excel не запущен: {self.error}')

    def _run(self):
        import pythoncom
        import pywintypes
        import win32api
        handle = None
        pythoncom.CoInitialize()
        try:
            self.thread_id = win32api.GetCurrentThreadId()
            register_progid()
            obj = to_com(self.factory, self.factory)
            handle = pythoncom.RegisterActiveObject(obj, pywintypes.IID(CLSID), 0)
            log.info('KK.ExcelShim зарегистрирован в ROT')
            self.ready.set()
            pythoncom.PumpMessages()
        except Exception as exc:
            log.exception('Эмулятор Excel не запущен')
            self.error = str(exc)
            self.ready.set()
        finally:
            if handle is not None:
                try:
                    pythoncom.RevokeActiveObject(handle)
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def stop(self):
        if self.thread_id:
            import win32api
            import win32con
            win32api.PostThreadMessage(self.thread_id, win32con.WM_QUIT, 0, 0)
