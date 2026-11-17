"""Эмулятор подмножества объектной модели Excel поверх openpyxl.

Макросы получают объекты через COM (kk.xlshim.com), программа использует их напрямую.
"""
from .ranges import Range, ShimError, parse_address  # noqa: F401
from .core import Book, Sheet, load_workbook  # noqa: F401
from .app import App, Factory, FileDialog  # noqa: F401
