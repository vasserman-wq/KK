"""Адреса и диапазоны ячеек: поведение Excel Range в пределах того, что используют макросы."""
from __future__ import annotations

import datetime as _dt
import re

from openpyxl.utils import column_index_from_string, get_column_letter

MAX_ROWS = 1048576
MAX_COLS = 16384
XL_UP, XL_DOWN, XL_TO_LEFT, XL_TO_RIGHT = -4162, -4121, -4159, -4161
_EPOCH = _dt.datetime(1899, 12, 30)
_PART = re.compile(r'^([A-Z]{0,3})(\d*)$')


class ShimError(Exception):
    """Ошибка, которую VBA получит как ошибку вызова объекта (Err.Description)."""


def _part(text):
    m = _PART.match(text)
    if not m or (not m.group(1) and not m.group(2)):
        raise ShimError(f'Неподдерживаемый адрес: {text}')
    col = column_index_from_string(m.group(1)) if m.group(1) else None
    row = int(m.group(2)) if m.group(2) else None
    return row, col


def parse_address(text):
    """'A1', 'A1:E10', 'A:A', '7:9', "='Лист'!$A$1:$A$5" -> (лист|None, r1, c1, r2, c2)."""
    text = str(text).strip()
    sheet = None
    if '!' in text:
        sheet, text = text.rsplit('!', 1)
        sheet = sheet.lstrip('=')
        if sheet.startswith("'") and sheet.endswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
    text = text.lstrip('=').replace('$', '').upper()
    a, _, b = text.partition(':')
    r1, c1 = _part(a)
    r2, c2 = _part(b) if b else (r1, c1)
    if r1 is None or r2 is None:
        r1, r2 = 1, MAX_ROWS
    if c1 is None or c2 is None:
        c1, c2 = 1, MAX_COLS
    return sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def to_value2(value):
    """Значение openpyxl -> как Excel Value2: числа Double, даты числом, пусто None."""
    if value is None or isinstance(value, (bool, str)):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, _dt.datetime):
        return (value - _EPOCH).total_seconds() / 86400.0
    if isinstance(value, _dt.date):
        return float((_dt.datetime(value.year, value.month, value.day) - _EPOCH).days)
    if isinstance(value, _dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second) / 86400.0
    return str(value)


def from_vba(value):
    """Значение из VBA -> в ячейку. Целые Double хранятся как int (так же пишет Excel)."""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer() and abs(value) < 1e15:
        return int(value)
    return value


class Range:
    """Прямоугольный диапазон листа. Индексы с 1, как в Excel."""

    def __init__(self, sheet, r1, c1, r2=None, c2=None):
        r2 = r1 if r2 is None else r2
        c2 = c1 if c2 is None else c2
        if not (1 <= r1 <= r2 <= MAX_ROWS and 1 <= c1 <= c2 <= MAX_COLS):
            raise ShimError(f'Диапазон вне листа: {r1},{c1}:{r2},{c2}')
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2

    @property
    def nrows(self):
        return self.r2 - self.r1 + 1

    @property
    def ncols(self):
        return self.c2 - self.c1 + 1

    def address(self):
        a = f'${get_column_letter(self.c1)}${self.r1}'
        if self.nrows == 1 and self.ncols == 1:
            return a
        return a + f':${get_column_letter(self.c2)}${self.r2}'

    def cells(self, row=None, col=None):
        if row is None:
            return self
        if col is None:
            n = int(row) - 1
            row, col = n // self.ncols + 1, n % self.ncols + 1
        if isinstance(col, str):
            col = column_index_from_string(col.upper())
        return Range(self.sheet, self.r1 + int(row) - 1, self.c1 + int(col) - 1)

    def resize(self, rows=None, cols=None):
        rows = self.nrows if rows is None else int(rows)
        cols = self.ncols if cols is None else int(cols)
        return Range(self.sheet, self.r1, self.c1, self.r1 + rows - 1, self.c1 + cols - 1)

    def offset(self, rows=None, cols=None):
        rows, cols = int(rows or 0), int(cols or 0)
        return Range(self.sheet, self.r1 + rows, self.c1 + cols, self.r2 + rows, self.c2 + cols)

    def row_range(self, index):
        r = self.r1 + int(index) - 1
        return Range(self.sheet, r, self.c1, r, self.c2)

    def column_range(self, index):
        if isinstance(index, str):
            _, _, c1, _, c2 = parse_address(index if ':' in index else f'{index}:{index}')
            return Range(self.sheet, self.r1, c1, self.r2, c2)
        c = self.c1 + int(index) - 1
        return Range(self.sheet, self.r1, c, self.r2, c)

    def _raw(self, r, c):
        cell = self.sheet.ows._cells.get((r, c))
        return None if cell is None else cell.value

    def _existing(self):
        return [(r, c) for (r, c) in list(self.sheet.ows._cells)
                if self.r1 <= r <= self.r2 and self.c1 <= c <= self.c2]

    def get_value2(self):
        if self.nrows == 1 and self.ncols == 1:
            return to_value2(self._raw(self.r1, self.c1))
        if self.nrows * self.ncols > 2_000_000:
            raise ShimError('Слишком большой диапазон для чтения.')
        return [[to_value2(self._raw(r, c)) for c in range(self.c1, self.c2 + 1)]
                for r in range(self.r1, self.r2 + 1)]

    def set_value2(self, value):
        ws = self.sheet.ows
        self.sheet.book.check_writable()
        if isinstance(value, (list, tuple)):
            rows = [list(v) if isinstance(v, (list, tuple)) else [v] for v in value]
            if len(rows) == 1 and self.nrows > 1 and self.ncols == 1:
                rows = [[v] for v in rows[0]]
            for i, row in enumerate(rows[:self.nrows]):
                for j, item in enumerate(row[:self.ncols]):
                    ws.cell(self.r1 + i, self.c1 + j).value = from_vba(item)
        else:
            if self.nrows * self.ncols > 200_000:
                raise ShimError('Слишком большой диапазон для записи.')
            for r in range(self.r1, self.r2 + 1):
                for c in range(self.c1, self.c2 + 1):
                    if value is None and (r, c) not in ws._cells:
                        continue
                    ws.cell(r, c).value = from_vba(value)
        self.sheet.book.touch()

    def has_formula(self):
        return any(isinstance(self._raw(r, c), str) and self._raw(r, c).startswith('=')
                   for r, c in self._existing())

    def get_formula(self):
        v = self._raw(self.r1, self.c1)
        if v is None:
            return ''
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    def clear(self, contents_only=False):
        self.sheet.book.check_writable()
        if not contents_only:
            self.unmerge()
        for r, c in self._existing():
            cell = self.sheet.ows._cells[(r, c)]
            cell.value = None
            if not contents_only and hasattr(cell, 'number_format'):
                cell.number_format = 'General'
        self.sheet.book.touch()

    def count_a(self):
        return sum(1 for r, c in self._existing() if self._raw(r, c) not in (None, ''))

    def end(self, direction):
        """Range.End: как Ctrl+стрелка в Excel."""
        direction = int(direction)
        r, c = self.r1, self.c1
        filled = self.sheet.filled
        if direction in (XL_UP, XL_DOWN):
            step = -1 if direction == XL_UP else 1
            nxt = r + step
            if filled(r, c) and 1 <= nxt <= MAX_ROWS and filled(nxt, c):
                while 1 <= nxt + step <= MAX_ROWS and filled(nxt + step, c):
                    nxt += step
                return Range(self.sheet, nxt, c)
            rows = [rr for (rr, cc) in self.sheet.filled_cells() if cc == c and (rr - r) * step > 0]
            if not rows:
                return Range(self.sheet, 1 if step < 0 else MAX_ROWS, c)
            return Range(self.sheet, max(rows) if step < 0 else min(rows), c)
        step = -1 if direction == XL_TO_LEFT else 1
        nxt = c + step
        if filled(r, c) and 1 <= nxt <= MAX_COLS and filled(r, nxt):
            while 1 <= nxt + step <= MAX_COLS and filled(r, nxt + step):
                nxt += step
            return Range(self.sheet, r, nxt)
        cols = [cc for (rr, cc) in self.sheet.filled_cells() if rr == r and (cc - c) * step > 0]
        if not cols:
            return Range(self.sheet, r, 1 if step < 0 else MAX_COLS)
        return Range(self.sheet, r, max(cols) if step < 0 else min(cols))

    def insert(self, shift=None):
        self.sheet.shift_rows(self.r1, self.nrows, self.c1, self.c2, down=True)

    def delete(self, shift=None):
        self.sheet.shift_rows(self.r1, self.nrows, self.c1, self.c2, down=False)

    def copy_to(self, dest):
        values = self.get_value2()
        if not isinstance(values, list):
            values = [[values]]
        Range(dest.sheet, dest.r1, dest.c1, dest.r1 + len(values) - 1,
              dest.c1 + len(values[0]) - 1).set_value2(values)

    def _merged_hits(self):
        return [mr for mr in list(self.sheet.ows.merged_cells.ranges)
                if not (mr.max_row < self.r1 or mr.min_row > self.r2
                        or mr.max_col < self.c1 or mr.min_col > self.c2)]

    def merge(self):
        if self.nrows * self.ncols > 1:
            self.unmerge()
            self.sheet.ows.merge_cells(start_row=self.r1, start_column=self.c1,
                                       end_row=self.r2, end_column=self.c2)
            self.sheet.book.touch()

    def unmerge(self):
        for mr in self._merged_hits():
            self.sheet.ows.unmerge_cells(str(mr))
            self.sheet.book.touch()

    def merge_cells(self):
        return bool(self._merged_hits())

    def set_number_format(self, fmt):
        last = min(self.r2, max(self.sheet.max_row(), self.r1) + 50)
        for r in range(self.r1, last + 1):
            for c in range(self.c1, min(self.c2, self.c1 + 63) + 1):
                cell = self.sheet.ows.cell(r, c)
                if hasattr(cell, 'number_format'):
                    cell.number_format = str(fmt)

    def get_number_format(self):
        cell = self.sheet.ows._cells.get((self.r1, self.c1))
        return getattr(cell, 'number_format', 'General') or 'General'
