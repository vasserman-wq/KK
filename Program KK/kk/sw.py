"""SOLIDWORKS через COM (pywin32): подключение, запуск макросов, список объектов, подсветка.

Подключение и RunMacro2 перенесены из kc_start.py без изменения проверок.
Методы обхода видов и аннотаций те же, что в EnumSheetAnns макроса 1 (GetFirstView, GetName2,
GetFirstDisplayDimension5, GetNext5, GetFirstNote, GetNext, GetAnnotation, GetName, GetPosition).
НЕ ПРОВЕРЕНО из Python: позднее связывание этих методов и выделение IAnnotation.Select3.
"""
from __future__ import annotations

import logging
import os
import re
import secrets
import tempfile
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

log = logging.getLogger(__name__)
SW_DOC_DRAWING = 3
SW_METHODS_WITHOUT_ARGUMENTS = 1
PROCEDURE_RE = re.compile(r'[A-Za-z_][A-Za-z_0-9]*')
PROGRESS_STATUSES = ('menu', 'running', 'done', 'stopped', 'error')
PROGRESS_TOTAL = 8


class SwError(Exception):
    pass


def _m(obj, name, *args):
    """Вызов метода SOLIDWORKS через позднее связывание (как _FlagAsMethod в kc_start.py)."""
    try:
        obj._FlagAsMethod(name)
    except AttributeError:
        pass
    return getattr(obj, name)(*args)


@contextmanager
def session():
    """Подключение к запущенному SOLIDWORKS в текущем потоке."""
    import pythoncom
    import pywintypes
    import win32com.client.dynamic
    pythoncom.CoInitialize()
    try:
        try:
            running = pythoncom.GetActiveObject(pywintypes.IID('SldWorks.Application'))
        except pywintypes.com_error:
            raise SwError('SOLIDWORKS не запущен или недоступен. Запустите SOLIDWORKS и откройте чертёж.') from None
        sw = win32com.client.dynamic.Dispatch(running.QueryInterface(pythoncom.IID_IDispatch))
        sw._FlagAsMethod('GetMacroMethods', 'RunMacro2')
        yield sw
    finally:
        pythoncom.CoUninitialize()


def active_drawing(sw):
    doc = sw.ActiveDoc
    if doc is None:
        raise SwError('В SOLIDWORKS нет активного документа. Откройте чертёж.')
    doc._FlagAsMethod('GetType', 'GetTitle', 'GetPathName')
    if int(doc.GetType()) != SW_DOC_DRAWING:
        raise SwError('Активный документ — не чертёж.')
    path = str(doc.GetPathName() or '')
    if not path:
        raise SwError('Сначала сохраните чертёж.')
    return doc, path, str(doc.GetTitle() or Path(path).name)


def select_entry(methods, procedure):
    """Как в kc_start.py: строки «Module.Proc», ровно одно совпадение по процедуре."""
    found = []
    for item in methods or ():
        module, _, proc = str(item).rpartition('.')
        if module and proc.lower() == procedure.lower():
            found.append(module)
    if not found:
        raise SwError(f'Не найдена процедура {procedure}. Доступные методы: {", ".join(map(str, methods or ())) or "нет"}')
    if len(found) > 1:
        raise SwError(f'В SWP несколько подходящих процедур {procedure}: {", ".join(found)}. Укажите модуль в настройках.')
    return found[0]


def check_macro_file(path):
    p = Path(str(path or ''))
    if not str(path or '').strip():
        raise SwError('Не указан путь к макросу (.swp). Укажите его в меню «Сервис → Пути к макросам».')
    if p.suffix.lower() != '.swp':
        raise SwError(f'Макрос должен быть файлом .swp, а не {p.suffix or "без расширения"}: {p}')
    if not p.is_file():
        raise SwError(f'Файл макроса не найден: {p}')
    return p


def run_macro(sw, macro, module='', procedure='Main'):
    """RunMacro2 (синхронно). True от RunMacro2 означает только успешный вызов, а не успех VBA. Повторов нет."""
    import pythoncom
    from win32com.client import VARIANT
    macro = str(check_macro_file(macro))
    if not PROCEDURE_RE.fullmatch(procedure or ''):
        raise SwError(f'Недопустимое имя процедуры: {procedure!r}')
    if any(ch in (module or '') for ch in '\r\n\t'):
        raise SwError('Недопустимое имя модуля.')
    if not module:
        module = select_entry(sw.GetMacroMethods(macro, SW_METHODS_WITHOUT_ARGUMENTS), procedure)
    error = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    result = sw.RunMacro2(macro, module, procedure, 1, error)
    if isinstance(result, tuple):
        if len(result) != 2:
            raise SwError('Неизвестный формат ответа RunMacro2. Автоматический повтор запрещён.')
        ok, code = result
    else:
        ok, code = result, error.value
    if not isinstance(ok, (bool, int)) or not ok or int(code or 0) != 0:
        raise SwError(f'SOLIDWORKS не подтвердил запуск макроса. RunMacro2: {ok}; код: {code}. Проверьте макрос и сообщения SOLIDWORKS.')
    return module


class Progress:
    """Канал KC_PROGRESS_V1 (тот же, что читает KC_ProgressConnect в макросе 1)."""

    def __init__(self):
        self.folder = Path(tempfile.mkdtemp(prefix='KC_progress_'))
        self.state = self.folder / 'state.txt'
        self.nonce = secrets.token_hex(16)
        self.descriptor = Path(tempfile.gettempdir()) / 'KC_Start_progress_session.txt'

    def __enter__(self):
        self.descriptor.write_text(f'KC_PROGRESS_V1\n{self.state}\n{self.nonce}\n', encoding='utf-16')
        return self

    def __exit__(self, *exc):
        for path in (self.descriptor, self.state):
            try:
                path.unlink()
            except OSError:
                pass
        try:
            self.folder.rmdir()
        except OSError:
            pass

    def read(self):
        """(status, completed, total, label) или None, если файла нет или он чужой."""
        try:
            lines = self.state.read_text(encoding='utf-16').splitlines()
        except (OSError, UnicodeError):
            return None
        if len(lines) < 5 or lines[0].strip() != self.nonce or lines[1].strip() not in PROGRESS_STATUSES:
            return None
        try:
            completed, total = int(lines[2]), int(lines[3])
        except ValueError:
            return None
        if total != PROGRESS_TOTAL or not 0 <= completed <= total:
            return None
        return lines[1].strip(), completed, total, lines[4].strip()


@dataclass
class Candidate:
    sheet: str
    view: str
    name: str
    kind: str        # DIM | NOTE
    text: str
    x: float
    y: float
    marker: bool     # в чертеже уже стоит # (для размера) или @ (для заметки)

    @property
    def default_include(self):
        return (not self.marker) if self.kind == 'DIM' else self.marker


def _dim_text(dd):
    parts = []
    for index in (1, 0, 2):  # префикс, основной, суффикс (порядок для показа; не проверено)
        try:
            text = str(_m(dd, 'GetText', index) or '')
        except Exception:
            text = ''
        if text:
            parts.append(text)
    try:
        dim = _m(dd, 'GetDimension2', 0)
        value = float(_m(dim, 'GetSystemValue3', 1, None)[0]) * 1000 if dim is not None else None
    except Exception:
        value = None
    if value is not None and not any(ch.isdigit() for ch in ''.join(parts)):
        parts.insert(0, f'{value:g}')
    return ' '.join(parts).replace('<MOD-DIAM>', 'Ø').replace('<MOD-PM>', '±').replace('<MOD-DEG>', '°')


def candidates(sw, doc):
    """Все размеры и заметки всех листов. Активный лист возвращается на место."""
    result = []
    original = str(_m(_m(doc, 'GetCurrentSheet'), 'GetName'))
    try:
        for sheet in _m(doc, 'GetSheetNames') or ():
            sheet = str(sheet)
            _m(doc, 'ActivateSheet', sheet)
            view = _m(doc, 'GetFirstView')
            while view is not None:
                view_name = str(_m(view, 'GetName2'))
                dd = _m(view, 'GetFirstDisplayDimension5')
                while dd is not None:
                    ann = _m(dd, 'GetAnnotation')
                    if ann is not None:
                        pos = _m(ann, 'GetPosition')
                        text = _dim_text(dd)
                        result.append(Candidate(sheet, view_name, str(_m(ann, 'GetName')), 'DIM', text,
                                                float(pos[0]), float(pos[1]), '#' in text))
                    dd = _m(dd, 'GetNext5')
                note = _m(view, 'GetFirstNote')
                while note is not None:
                    text = str(_m(note, 'GetText') or '')
                    ann = _m(note, 'GetAnnotation')
                    if ann is not None and text.strip() and not text.strip().isdigit():  # цифра — баллон
                        pos = _m(ann, 'GetPosition')
                        result.append(Candidate(sheet, view_name, str(_m(ann, 'GetName')), 'NOTE',
                                                ' '.join(text.split()), float(pos[0]), float(pos[1]), '@' in text))
                    note = _m(note, 'GetNext')
                view = _m(view, 'GetNextView')
    finally:
        _m(doc, 'ActivateSheet', original)
    return result


def highlight(sw, doc, cand):
    """Показать объект: активировать лист и выделить аннотацию. НЕ ПРОВЕРЕНО: Select3 через позднее связывание."""
    _m(doc, 'ActivateSheet', cand.sheet)
    _m(doc, 'ClearSelection2', True)
    view = _m(doc, 'GetFirstView')
    while view is not None:
        items = []
        dd = _m(view, 'GetFirstDisplayDimension5')
        while dd is not None:
            items.append(dd)
            dd = _m(dd, 'GetNext5')
        note = _m(view, 'GetFirstNote')
        while note is not None:
            items.append(note)
            note = _m(note, 'GetNext')
        for obj in items:
            ann = _m(obj, 'GetAnnotation')
            if ann is None or str(_m(ann, 'GetName')) != cand.name:
                continue
            pos = _m(ann, 'GetPosition')
            if abs(float(pos[0]) - cand.x) <= 5e-5 and abs(float(pos[1]) - cand.y) <= 5e-5:
                ok = _m(ann, 'Select3', False, None)
                _m(doc, 'GraphicsRedraw2')
                return bool(ok)
        view = _m(view, 'GetNextView')
    return False
