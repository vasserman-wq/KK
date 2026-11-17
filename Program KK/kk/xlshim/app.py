"""«Приложение Excel» эмулятора и фабрика, которую видит VBA (GetObject(, "KK.ExcelShim"))."""
from __future__ import annotations

import threading
from pathlib import Path

import openpyxl

from .core import EDITOR_PING, Book, load_workbook
from .ranges import ShimError

EDITOR_COMMANDS = {'ok_editorping': EDITOR_PING, 'ok_refreshbuttons': None,
                   'ok_updateeditor': None, 'ok_variantschanged': None}


class FileDialog:
    """Application.FileDialog(3): выбор файла. Окно показывает программа (dialogs.open_file)."""

    def __init__(self, factory):
        self.factory = factory
        self.title, self.initial, self.multi = '', '', False
        self.filters = []
        self.selected = []

    def show(self):
        path = self.factory.dialogs.open_file(self.title, self.initial, self.filters)
        self.selected = [path] if path else []
        return -1 if path else 0


class App:
    """Один «экземпляр Excel»."""

    def __init__(self, factory):
        self.factory = factory
        self.books = []
        self.props = {'visible': False, 'enableevents': True, 'screenupdating': True,
                      'displayalerts': True, 'calculation': -4105, 'automationsecurity': 1,
                      'interactive': True, 'statusbar': False, 'cutcopymode': False}
        self.quit_called = False
        self._new_count = 0

    def new_book_name(self, book):
        if not hasattr(book, '_temp_name'):
            self._new_count += 1
            book._temp_name = f'Книга{self._new_count}'
        return book._temp_name

    def find_open(self, path):
        key = str(Path(path)).lower()
        for app in self.factory.apps:
            for book in app.books:
                if book.path and str(Path(book.path)).lower() == key:
                    return app, book
        return None, None

    def open(self, path, update_links=None, read_only=False):
        path = str(path)
        owner, book = self.find_open(path)
        if owner is self:
            return book
        ro = read_only is True or read_only == -1
        if owner is not None and not book.saved:
            ro = True  # как Excel: вторая копия занятой книги открывается только для чтения
        with self.factory.io_lock:
            owb = load_workbook(path)
        book = Book(self, owb, path, ro)
        self.books.append(book)
        return book

    def add(self):
        owb = openpyxl.Workbook()
        owb.active.title = 'Лист1'
        book = Book(self, owb)
        book.saved = False
        self.books.append(book)
        return book

    def book(self, key):
        if isinstance(key, (int, float)) and not isinstance(key, bool):
            if 1 <= int(key) <= len(self.books):
                return self.books[int(key) - 1]
            raise ShimError('Нет книги с таким номером.')
        for book in self.books:
            if book.name.lower() == str(key).lower() or book.full_name.lower() == str(key).lower():
                return book
        raise ShimError(f'Книга «{key}» не открыта.')

    @property
    def active_book(self):
        return self.books[-1] if self.books else None

    def run(self, qualified, *args):
        command = str(qualified).rsplit('!', 1)[-1].strip().lower()
        if command in EDITOR_COMMANDS:
            return EDITOR_COMMANDS[command]
        raise ShimError(f'Команда книги не поддерживается программой: {qualified}')

    def get_save_as_filename(self, initial='', file_filter='', filter_index=None, title=''):
        path = self.factory.dialogs.save_file(str(title or ''), str(initial or ''), str(file_filter or ''))
        return path if path else False

    def get_open_filename(self, file_filter='', filter_index=None, title='', button=None, multi=False):
        path = self.factory.dialogs.open_file(str(title or ''), '', [str(file_filter or '')])
        return path if path else False

    def quit(self):
        for book in list(self.books):
            book.close(False)
        self.quit_called = True
        if self in self.factory.apps:
            self.factory.apps.remove(self)


class NoDialogs:
    """Диалоги по умолчанию (тесты): всегда «Отмена»."""

    def open_file(self, title, initial, filters):
        return ''

    def save_file(self, title, initial, file_filter):
        return ''


class Factory:
    """Общая точка для VBA: новые экземпляры, команды программы и решения по разметке."""

    def __init__(self, dialogs=None):
        self.apps = []
        self.dialogs = dialogs or NoDialogs()
        self.io_lock = threading.RLock()
        self.commands = {}
        self.decisions = {}   # (путь чертежа, лист, вид, имя) -> 0/1
        self.decision_drawing = ''
        self.log = []

    def new_app(self):
        app = App(self)
        self.apps.append(app)
        return app

    def active_app(self):
        for app in reversed(self.apps):
            if not app.quit_called:
                return app
        raise ShimError('Нет запущенного экземпляра.')

    # --- команды программы для макроса (читаются один раз)
    def set_command(self, name, value):
        self.commands[str(name).lower()] = str(value)

    def take_command(self, name):
        return self.commands.pop(str(name).lower(), '')

    # --- решения «включить в карту». Ключ: лист + имя аннотации + положение на листе.
    POSITION_TOLERANCE_M = 0.00005  # 0,05 мм, как в GetViewNameForAnnotation макроса 1

    def set_decisions(self, drawing, rows):
        """rows: (лист, имя аннотации, x м, y м, включить: bool)."""
        table = {}
        for sheet, name, x, y, include in rows:
            table.setdefault((str(sheet).strip(), str(name).strip()), []).append((float(x), float(y), 1 if include else 0))
        self.decisions = table
        self.decision_drawing = str(drawing).strip().lower()

    def clear_decisions(self):
        self.decisions, self.decision_drawing = {}, ''

    def has_decisions(self, drawing):
        return bool(self.decisions) and str(drawing).strip().lower() == self.decision_drawing

    def decision(self, drawing, sheet, name, x, y):
        """-1: решения нет (макрос проверяет маркеры по-старому), 0: не включать, 1: включить.

        Если объект сдвинули после выбора в программе, решения нет: действует прежнее правило.
        """
        if not self.has_decisions(drawing):
            return -1
        try:
            x, y = float(x), float(y)
        except (TypeError, ValueError):
            return -1
        hits = [d for (px, py, d) in self.decisions.get((str(sheet).strip(), str(name).strip()), [])
                if abs(px - x) <= self.POSITION_TOLERANCE_M and abs(py - y) <= self.POSITION_TOLERANCE_M]
        return hits[0] if len(hits) == 1 else -1
