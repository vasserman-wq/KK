"""Книги, листы, имена и «приложение» эмулятора Excel. Только Python, без COM.

Каждый App ведёт себя как отдельный экземпляр Excel (CreateObject создаёт новый,
GetObject возвращает последний живой). Обмен данными между экземплярами и с программой
идёт через файлы на диске, как с настоящим Excel.
"""
from __future__ import annotations

import os
import tempfile
import threading
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.workbook.defined_name import DefinedName

from .ranges import MAX_COLS, MAX_ROWS, Range, ShimError, parse_address

EDITOR_PING = 'OK_INTRO_EDITOR_V19'
_STATES = {'visible': -1, 'hidden': 0, 'veryHidden': 2}
_STATES_BACK = {-1: 'visible', 1: 'visible', True: 'visible', 0: 'hidden', False: 'hidden', 2: 'veryHidden'}


class Sheet:
    def __init__(self, book, ows):
        self.book, self.ows = book, ows

    @property
    def name(self):
        return self.ows.title

    def set_name(self, value):
        value = str(value)
        for other in self.book.owb.worksheets:
            if other is not self.ows and other.title.lower() == value.lower():
                raise ShimError(f'Лист с именем «{value}» уже есть.')
        self.ows.title = value
        self.book.touch()

    def filled(self, r, c):
        cell = self.ows._cells.get((r, c))
        return cell is not None and cell.value not in (None, '')

    def filled_cells(self):
        return [k for k, cell in self.ows._cells.items() if cell.value not in (None, '')]

    def max_row(self):
        rows = [r for r, _ in self.filled_cells()]
        return max(rows) if rows else 0

    def range(self, a, b=None):
        if isinstance(a, Range):
            first = a
        else:
            sheet, r1, c1, r2, c2 = parse_address(a)
            target = self if not sheet or sheet == self.name else self.book.sheet(sheet)
            first = Range(target, r1, c1, r2, c2)
        if b is None:
            return first
        second = b if isinstance(b, Range) else self.range(b)
        return Range(first.sheet, min(first.r1, second.r1), min(first.c1, second.c1),
                     max(first.r2, second.r2), max(first.c2, second.c2))

    def cells(self, r=None, c=None):
        if r is None:
            return Range(self, 1, 1, MAX_ROWS, MAX_COLS)
        if c is None:
            n = int(r) - 1
            return Range(self, n // MAX_COLS + 1, n % MAX_COLS + 1)
        if isinstance(c, str):
            c = column_index_from_string(c.upper())
        return Range(self, int(r), int(c))

    def rows(self, index=None):
        if index is None:
            return Range(self, 1, 1, MAX_ROWS, MAX_COLS)
        if isinstance(index, str):
            _, r1, _, r2, _ = parse_address(index if ':' in index else f'{index}:{index}')
            return Range(self, r1, 1, r2, MAX_COLS)
        return Range(self, int(index), 1, int(index), MAX_COLS)

    def columns(self, index=None):
        if index is None:
            return Range(self, 1, 1, MAX_ROWS, MAX_COLS)
        if isinstance(index, str):
            _, _, c1, _, c2 = parse_address(index if ':' in index else f'{index}:{index}')
            return Range(self, 1, c1, MAX_ROWS, c2)
        return Range(self, 1, int(index), MAX_ROWS, int(index))

    def used_range(self):
        cells = self.filled_cells()
        if not cells:
            return Range(self, 1, 1)
        return Range(self, min(r for r, _ in cells), min(c for _, c in cells),
                     max(r for r, _ in cells), max(c for _, c in cells))

    @property
    def visible(self):
        return _STATES.get(self.ows.sheet_state, -1)

    def set_visible(self, value):
        try:
            value = int(value)
        except (TypeError, ValueError):
            value = -1
        self.ows.sheet_state = _STATES_BACK.get(value, 'visible')
        self.book.touch()

    def delete(self):
        self.book.check_writable()
        if len(self.book.owb.worksheets) == 1:
            raise ShimError('Нельзя удалить единственный лист книги.')
        self.book.owb.remove(self.ows)
        self.book.touch()

    def shift_rows(self, start, count, c1, c2, down):
        """Insert/Delete со сдвигом только в колонках c1..c2 (как Range("A7:E10").Insert)."""
        self.book.check_writable()
        ws = self.ows
        if not down:
            for key in [k for k in ws._cells if start <= k[0] < start + count and c1 <= k[1] <= c2]:
                del ws._cells[key]
        moving = [k for k in ws._cells if k[0] >= (start if down else start + count) and c1 <= k[1] <= c2]
        moving.sort(key=lambda k: k[0], reverse=down)
        delta = count if down else -count
        for (r, c) in moving:
            cell = ws._cells.pop((r, c))
            cell.row = r + delta
            ws._cells[(r + delta, c)] = cell
        for mr in list(ws.merged_cells.ranges):
            inside = mr.min_col >= c1 and mr.max_col <= c2
            if not inside or mr.max_row < start:
                continue
            if not down and mr.min_row >= start and mr.max_row < start + count:
                ws.merged_cells.ranges.remove(mr)
            elif mr.min_row >= (start if down else start + count):
                mr.shift(row_shift=delta)
        self.book.touch()


class NameObj:
    def __init__(self, book, name):
        self.book, self.name = book, name

    @property
    def _dn(self):
        dn = self.book.owb.defined_names.get(self.name)
        if dn is None:
            raise ShimError(f'Имя «{self.name}» не найдено.')
        return dn

    @property
    def refers_to(self):
        return '=' + (self._dn.attr_text or '')

    def refers_to_range(self):
        sheet, r1, c1, r2, c2 = parse_address(self._dn.attr_text)
        if not sheet:
            raise ShimError(f'Имя «{self.name}» не ссылается на диапазон.')
        return Range(self.book.sheet(sheet), r1, c1, r2, c2)

    def delete(self):
        self.book.owb.defined_names.pop(self.name, None)
        self.book.touch()


class Names:
    def __init__(self, book):
        self.book = book

    def keys(self):
        return list(self.book.owb.defined_names.keys())

    def item(self, key):
        keys = self.keys()
        if isinstance(key, (int, float)):
            index = int(key)
            if not 1 <= index <= len(keys):
                raise ShimError('Нет имени с таким номером.')
            return NameObj(self.book, keys[index - 1])
        for k in keys:
            if k.lower() == str(key).lower():
                return NameObj(self.book, k)
        raise ShimError(f'Имя «{key}» не найдено.')

    def add(self, name, refers_to):
        text = str(refers_to)
        text = text[1:] if text.startswith('=') else text
        self.book.owb.defined_names[str(name)] = DefinedName(str(name), attr_text=text)
        self.book.touch()
        return NameObj(self.book, str(name))


class Book:
    def __init__(self, app, owb, path=None, read_only=False):
        self.app, self.owb, self.read_only = app, owb, bool(read_only)
        self.path = str(path) if path else None
        self.saved = True
        self._sheets = {}
        self.names = Names(self)

    def touch(self):
        self.saved = False

    def check_writable(self):
        return True  # Excel позволяет править книгу «только для чтения»; запрещено только Save

    @property
    def name(self):
        return Path(self.path).name if self.path else self.app.new_book_name(self)

    @property
    def full_name(self):
        return self.path or self.name

    @property
    def folder(self):
        return str(Path(self.path).parent) if self.path else ''

    def wrap(self, ows):
        sheet = self._sheets.get(id(ows))
        if sheet is None or sheet.ows is not ows:
            sheet = self._sheets[id(ows)] = Sheet(self, ows)
        return sheet

    def sheets(self):
        return [self.wrap(ws) for ws in self.owb.worksheets]

    def sheet(self, key):
        sheets = self.sheets()
        if isinstance(key, Sheet):
            return key
        if isinstance(key, (int, float)) and not isinstance(key, bool):
            index = int(key)
            if 1 <= index <= len(sheets):
                return sheets[index - 1]
            raise ShimError(f'Нет листа с номером {index}.')
        for sheet in sheets:
            if sheet.name.lower() == str(key).lower():
                return sheet
        raise ShimError(f'В книге «{self.name}» нет листа «{key}».')

    def has_sheet(self, name):
        return any(s.name.lower() == str(name).lower() for s in self.sheets())

    def add_sheet(self, before=None, after=None):
        sheets = self.owb.worksheets
        if isinstance(after, Sheet):
            index = sheets.index(after.ows) + 1
        elif isinstance(before, Sheet):
            index = sheets.index(before.ows)
        else:
            index = sheets.index(self.owb.active) if self.owb.active in sheets else 0
        n = 1
        while self.has_sheet(f'Лист{n}'):
            n += 1
        ows = self.owb.create_sheet(f'Лист{n}', index)
        self.touch()
        return self.wrap(ows)

    def _write(self, target):
        target = Path(target)
        if target.suffix.lower() == '.xls':
            raise ShimError('Сохранение в старый формат .xls не поддерживается.')
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix='.kk_', suffix=target.suffix, dir=str(target.parent))
        os.close(fd)
        try:
            self.owb.save(tmp)
            os.replace(tmp, target)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    def save(self):
        if not self.path:
            raise ShimError('У новой книги нет пути: используйте SaveAs.')
        if self.read_only:
            raise ShimError(f'Книга открыта только для чтения: {self.path}')
        self._write(self.path)
        self.saved = True

    def save_as(self, path):
        self._write(path)
        self.path, self.read_only, self.saved = str(path), False, True

    def save_copy_as(self, path):
        self._write(path)

    def change_access(self, mode):
        self.read_only = (int(mode) == 3) if str(mode).lstrip('-').isdigit() else self.read_only

    def close(self, save_changes=False):
        if save_changes is True or save_changes == -1:
            self.save()
        if self in self.app.books:
            self.app.books.remove(self)


def load_workbook(path):
    path = Path(path)
    ext = path.suffix.lower()
    if ext == '.xls':
        raise ShimError(f'Формат .xls не поддерживается: сохраните файл как .xlsm или .xlsx ({path.name}).')
    if not path.exists():
        raise ShimError(f'Файл не найден: {path}')
    return openpyxl.load_workbook(str(path), keep_vba=ext in ('.xlsm', '.xltm'), data_only=False)
