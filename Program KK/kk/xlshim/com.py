"""COM-слой эмулятора: VBA в SOLIDWORKS получает объекты через GetObject(, "KK.ExcelShim").

НЕ ПРОВЕРЕНО В WINDOWS. Места, требующие проверки при первом запуске (см. docs/TESTING.md):
- имя свойства по умолчанию, которое pywin32 DynamicPolicy передаёт в _dynamic_ (ожидается '_value_');
- вид, в котором приходят пропущенные аргументы (VT_ERROR DISP_E_PARAMNOTFOUND);
- RegisterActiveObject и доставка вызовов из другого процесса (SLDWORKS.exe).
Именованные аргументы (After:=) не используются: скрипт tools/patch_macros.py заменяет их позиционными.
"""
from __future__ import annotations

import logging
import threading

from .app import App, Factory, FileDialog
from .core import Book, NameObj, Names, Sheet
from .ranges import MAX_COLS, MAX_ROWS, Range, ShimError

log = logging.getLogger(__name__)
PROGID = 'KK.ExcelShim'
CLSID = '{6C1B9E52-3F0B-4F5E-9C2A-8D4E1A7B2C31}'
PARAM_NOT_FOUND = -2147352572
METHOD, GET, PUT, PUTREF = 1, 2, 4, 8


class Missing:
    pass


MISSING = Missing()


class NoOp:
    """Оформление (Font, Borders, Interior, Validation...): без Excel не нужно, вызовы принимаются и игнорируются."""


class Coll:
    """Коллекция для COM: items() -> список, item(key) -> элемент."""

    def __init__(self, kind, owner):
        self.kind, self.owner = kind, owner

    def items(self):
        o = self.owner
        if self.kind == 'sheets':
            return o.sheets()
        if self.kind == 'books':
            return list(o.books)
        if self.kind == 'names':
            return [NameObj(o.book, k) for k in o.keys()]
        if self.kind == 'selected':
            return list(o.selected)
        if self.kind == 'filters':
            return list(o.filters)
        return []

    def item(self, key):
        o = self.owner
        if self.kind == 'sheets':
            return o.sheet(key)
        if self.kind == 'books':
            return o.book(key)
        if self.kind == 'names':
            return o.item(key)
        if self.kind == 'selected':
            return o.selected[int(key) - 1]
        if self.kind in ('rows', 'columns'):
            rng = o
            return rng.row_range(key) if self.kind == 'rows' else rng.column_range(key)
        if self.kind == 'areas':
            return self.owner
        return NoOp()

    def count(self):
        if self.kind == 'areas':
            return 1
        if self.kind == 'rows':
            return self.owner.nrows
        if self.kind == 'columns':
            return self.owner.ncols
        return len(self.items())


def _clean(args):
    out = []
    for a in args:
        if a is None or (isinstance(a, int) and a == PARAM_NOT_FOUND):
            out.append(MISSING)
        else:
            out.append(_unwrap(a))
    while out and out[-1] is MISSING:
        out.pop()
    return out


def _arg(args, i, default=None):
    return args[i] if len(args) > i and args[i] is not MISSING else default


def _unwrap(value):
    try:
        import pythoncom
        from win32com.server import util
        if isinstance(value, pythoncom.TypeIIDs.get(pythoncom.IID_IDispatch, type(None))):
            return util.unwrap(value).target
    except Exception:  # не наш объект или не Windows
        pass
    if isinstance(value, (list, tuple)):
        return [_unwrap(v) for v in value]
    return value


def dispatch(target, name, flags, args, factory):
    """Чистая функция диспетчеризации (тестируется без COM). Возвращает Python-значение."""
    n = name.lower()
    put = bool(flags & (PUT | PUTREF))
    a = _clean(args)
    value = a.pop() if put and a else None
    if isinstance(target, NoOp):
        return False if n == '_value_' else (None if put else NoOp())
    if isinstance(target, Coll):
        if n in ('_value_', 'item'):
            return target.item(a[0]) if a else target
        if n == 'count':
            return target.count()
        if n == '_newenum':
            return target.items()
        if n == 'add' and target.kind == 'sheets':
            return target.owner.add_sheet(_arg(a, 0), _arg(a, 1))
        if n == 'add' and target.kind == 'names':
            return target.owner.add(_arg(a, 0), _arg(a, 1))
        if n == 'add' and target.kind == 'books':
            return target.owner.add()
        if n == 'open' and target.kind == 'books':
            return target.owner.open(_arg(a, 0), _arg(a, 1), _arg(a, 2, False))
        if n == 'add' and target.kind == 'filters':
            target.owner.filters.append(f'{_arg(a, 0, "")}|{_arg(a, 1, "")}')
            return None
        if n == 'clear' and target.kind == 'filters':
            target.owner.filters.clear()
            return None
        if n in ('delete', 'autofit', 'select', 'insert') and target.kind in ('rows', 'columns'):
            rng = target.owner
            if n == 'delete':
                rng.delete()
            elif n == 'insert':
                rng.insert()
            return None
        if target.kind in ('rows', 'columns'):
            return None if put else NoOp()
        raise AttributeError(name)
    if isinstance(target, Range):
        return _range(target, n, put, a, value)
    if isinstance(target, Sheet):
        return _sheet(target, n, put, a, value)
    if isinstance(target, Book):
        return _book(target, n, put, a, value)
    if isinstance(target, App):
        return _app(target, n, put, a, value)
    if isinstance(target, NameObj):
        if n in ('name', '_value_') and not put:
            return target.name if n == 'name' else target.refers_to
        if n in ('refersto', 'value'):
            return target.refers_to
        if n == 'referstorange':
            return target.refers_to_range()
        if n == 'delete':
            return target.delete()
        raise AttributeError(name)
    if isinstance(target, FileDialog):
        if put:
            if n == 'title':
                target.title = str(value)
            elif n == 'initialfilename':
                target.initial = str(value)
            elif n == 'allowmultiselect':
                target.multi = bool(value)
            return None
        if n == 'show':
            return target.show()
        if n == 'filters':
            return Coll('filters', target)
        if n == 'selecteditems':
            coll = Coll('selected', target)
            return coll.item(a[0]) if a else coll
        if n in ('title', 'initialfilename', 'allowmultiselect'):
            return {'title': target.title, 'initialfilename': target.initial,
                    'allowmultiselect': target.multi}[n]
        raise AttributeError(name)
    if isinstance(target, WorksheetFunction):
        if n == 'counta':
            return float(sum(x.count_a() if isinstance(x, Range) else (0 if x in (None, '') else 1) for x in a))
        raise AttributeError(name)
    if isinstance(target, Factory):
        return _factory(target, n, a)
    raise AttributeError(name)


class WorksheetFunction:
    pass


def _range(r, n, put, a, value):
    if n in ('value2', 'value', '_value_', 'text'):
        if put:
            return r.set_value2(value)
        v = r.get_value2()
        if n == 'text':
            return '' if v is None else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
        return v
    if n == 'formula':
        return r.set_value2(value) if put else r.get_formula()
    if n == 'numberformat':
        return r.set_number_format(value) if put else r.get_number_format()
    if put:
        return None  # оформление: ColumnWidth, RowHeight, WrapText, HorizontalAlignment...
    if n == 'cells':
        return r.cells(_arg(a, 0), _arg(a, 1))
    if n == 'range':
        return r.sheet.range(_arg(a, 0), _arg(a, 1))
    if n == 'row':
        return r.r1
    if n == 'column':
        return r.c1
    if n == 'count':
        return r.nrows * r.ncols
    if n == 'rows':
        return Coll('rows', r).item(a[0]) if a else Coll('rows', r)
    if n == 'columns':
        return Coll('columns', r).item(a[0]) if a else Coll('columns', r)
    if n == 'entirerow':
        return Range(r.sheet, r.r1, 1, r.r2, MAX_COLS)
    if n == 'entirecolumn':
        return Range(r.sheet, 1, r.c1, MAX_ROWS, r.c2)
    if n == 'address':
        return r.address()
    if n == 'areas':
        return Coll('areas', r)
    if n == 'end':
        return r.end(_arg(a, 0, -4162))
    if n == 'resize':
        return r.resize(_arg(a, 0), _arg(a, 1))
    if n == 'offset':
        return r.offset(_arg(a, 0), _arg(a, 1))
    if n == 'hasformula':
        return r.has_formula()
    if n == 'mergecells':
        return r.merge_cells()
    if n in ('clear', 'clearcontents'):
        return r.clear(n == 'clearcontents')
    if n == 'delete':
        return r.delete()
    if n == 'insert':
        return r.insert()
    if n == 'copy':
        dest = _arg(a, 0)
        if not isinstance(dest, Range):
            raise ShimError('Copy без целевого диапазона не поддерживается.')
        return r.copy_to(dest)
    if n == 'merge':
        return r.merge()
    if n == 'unmerge':
        return r.unmerge()
    if n in ('parent', 'worksheet'):
        return r.sheet
    if n == 'columnwidth':
        return 8.43
    if n == 'rowheight':
        return 15.0
    return NoOp()  # Font, Interior, Borders, Validation, AutoFit, Select, AutoFilter, Activate...


def _sheet(s, n, put, a, value):
    if n == 'name':
        return s.set_name(value) if put else s.name
    if n == 'visible':
        return s.set_visible(value) if put else s.visible
    if put:
        return None
    if n == 'cells':
        return s.cells(_arg(a, 0), _arg(a, 1))
    if n == 'range':
        return s.range(_arg(a, 0), _arg(a, 1))
    if n == 'rows':
        return s.rows(a[0]) if a else Coll('rows', Range(s, 1, 1, MAX_ROWS, MAX_COLS))
    if n == 'columns':
        return s.columns(a[0]) if a else Coll('columns', Range(s, 1, 1, MAX_ROWS, MAX_COLS))
    if n == 'usedrange':
        return s.used_range()
    if n == 'delete':
        return s.delete()
    if n in ('parent',):
        return s.book
    if n == 'index':
        return s.book.sheets().index(s) + 1
    if n == 'names':
        return Coll('names', s.book.names)
    if n in ('checkboxes', 'buttons', 'shapes', 'oleobjects'):
        return NoOp() if a else Coll('empty', s)
    return NoOp()  # Activate, Select, Protect, Unprotect, PageSetup...


def _book(b, n, put, a, value):
    if n == 'saved':
        if put:
            b.saved = bool(value)
            return None
        return b.saved
    if put:
        return None
    if n == 'name':
        return b.name
    if n == 'fullname':
        return b.full_name
    if n == 'path':
        return b.folder
    if n == 'readonly':
        return b.read_only
    if n in ('worksheets', 'sheets'):
        coll = Coll('sheets', b)
        return coll.item(a[0]) if a else coll
    if n == 'names':
        coll = Coll('names', b.names)
        return coll.item(a[0]) if a else coll
    if n == 'save':
        return b.save()
    if n == 'saveas':
        return b.save_as(_arg(a, 0))
    if n == 'savecopyas':
        return b.save_copy_as(_arg(a, 0))
    if n == 'close':
        return b.close(_arg(a, 0, False))
    if n == 'changefileaccess':
        return b.change_access(_arg(a, 0, 2))
    if n == 'activesheet':
        return b.sheets()[0]
    if n in ('parent', 'application'):
        return b.app
    return NoOp()  # Activate, RefreshAll...


def _app(app, n, put, a, value):
    if n in app.props:
        if put:
            app.props[n] = value
            return None
        return app.props[n]
    if put:
        return None
    if n == 'workbooks':
        coll = Coll('books', app)
        return coll.item(a[0]) if a else coll
    if n == 'activeworkbook':
        return app.active_book
    if n == 'activesheet':
        return app.active_book.sheets()[0] if app.active_book else None
    if n == 'run':
        return app.run(_arg(a, 0, ''))
    if n == 'getsaveasfilename':
        return app.get_save_as_filename(_arg(a, 0, ''), _arg(a, 1, ''), _arg(a, 2), _arg(a, 3, ''))
    if n == 'getopenfilename':
        return app.get_open_filename(_arg(a, 0, ''), _arg(a, 1), _arg(a, 2, ''))
    if n == 'filedialog':
        return FileDialog(app.factory)
    if n == 'worksheetfunction':
        return WorksheetFunction()
    if n == 'quit':
        return app.quit()
    if n == 'name':
        return 'KK Excel Shim'
    if n == 'version':
        return '16.0'
    return NoOp()  # Calculate, Wait...


def _factory(f, n, a):
    if n == 'newapp':
        return f.new_app()
    if n == 'activeapp':
        return f.active_app()
    if n == 'command':
        return f.take_command(_arg(a, 0, ''))
    if n == 'hasdecisions':
        return f.has_decisions(_arg(a, 0, ''))
    if n == 'decision':
        return f.decision(_arg(a, 0, ''), _arg(a, 1, ''), _arg(a, 2, ''), _arg(a, 3), _arg(a, 4))
    if n == 'ping':
        return 'KK'
    raise AttributeError(n)
