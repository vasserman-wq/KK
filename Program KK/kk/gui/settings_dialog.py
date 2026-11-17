"""Окна «Пути к макросам» и «Настройки ОК» (общий файл рядом с макросом 1)."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .. import config
from ..settings_file import SettingsError, builtin_path, ensure_shared, read_params, write_values
from . import base


class PathsDialog:
    def __init__(self, app):
        self.app = app
        win = self.win = tk.Toplevel(app.root)
        win.title('Пути к макросам')
        win.geometry('860x300')
        win.transient(app.root)
        win.grab_set()
        frm = ttk.Frame(win, padding=16)
        frm.pack(fill='both', expand=True)
        frm.columnconfigure(1, weight=1)
        self.vars = {k: tk.StringVar(value=app.cfg.get(k, '')) for k in ('macro1', 'macro2', 'module1', 'module2')}
        rows = (('macro1', 'Макрос 1 (.swp)'), ('macro2', 'Макрос 2 (.swp)'))
        for r, (key, label) in enumerate(rows):
            base.labeled_entry(frm, r, label, self.vars[key], width=80)
            ttk.Button(frm, text='Обзор…', command=lambda k=key: self.browse(k)).grid(row=r, column=2, padx=(8, 0))
        base.labeled_entry(frm, 2, 'Модуль макроса 1', self.vars['module1'], width=30, hint='пусто — найти Main автоматически')
        base.labeled_entry(frm, 3, 'Модуль макроса 2', self.vars['module2'], width=30, hint='пусто — найти Main автоматически')
        ttk.Label(frm, text='Рядом с макросом 1 должен лежать Настройки_ОК.xlsx (если его нет, программа создаст его из встроенных значений).\n'
                  'Общие пути для всех можно задать файлом kk_config.json рядом с KK.exe; здесь сохраняются личные пути.',
                  style='Hint.TLabel', justify='left').grid(row=4, column=0, columnspan=3, sticky='w', pady=12)
        bar = ttk.Frame(frm)
        bar.grid(row=5, column=0, columnspan=3, sticky='e')
        ttk.Button(bar, text='Отмена', command=win.destroy).pack(side='right', padx=(8, 0))
        base.primary_button(bar, 'Сохранить', self.save, base.GREEN, base.GREEN_ACTIVE).pack(side='right')

    def browse(self, key):
        path = filedialog.askopenfilename(parent=self.win, title='Макрос SOLIDWORKS', filetypes=[('Макрос SOLIDWORKS', '*.swp')])
        if path:
            self.vars[key].set(str(Path(path)))

    def save(self):
        for key, var in self.vars.items():
            self.app.cfg[key] = var.get().strip()
        config.save_user(self.app.cfg)
        self.win.destroy()
        self.app.pages['drawing'].update_texts()


class SettingsDialog:
    """Правка значений (колонка B) на листах параметров и ячеек таблиц правил. Служебные ключи не меняются."""

    def __init__(self, app):
        self.app = app
        try:
            self.path, created = ensure_shared(app.cfg['macro1'])
        except (SettingsError, OSError) as exc:
            messagebox.showerror('Настройки ОК', f'{exc}\nУкажите путь к макросу 1 (Сервис → Пути к макросам).')
            return
        try:
            self.schema, self.params, self.tables = read_params(self.path)
            self.defaults = read_params(builtin_path())[1] if builtin_path().exists() else {}
        except (SettingsError, OSError, KeyError, ValueError) as exc:
            messagebox.showerror('Настройки ОК', f'Файл настроек не прочитан: {exc}')
            return
        self.mtime = self.path.stat().st_mtime
        self.changes = {}
        win = self.win = tk.Toplevel(app.root)
        win.title(f'Настройки ОК — {self.path}')
        win.geometry('1180x680')
        win.transient(app.root)
        top = ttk.Frame(win, padding=(12, 10))
        top.pack(fill='x')
        ttk.Label(top, text=('Общий файл настроек для всех пользователей. Перед записью делается резервная копия '
                             'в папке «Резерв настроек». Макросы читают настройки при запуске.')
                  + ('\nФайл только что создан из встроенных значений.' if created else ''),
                  style='Hint.TLabel', justify='left').pack(side='left')
        base.primary_button(top, 'Сохранить', self.save, base.GREEN, base.GREEN_ACTIVE).pack(side='right')
        ttk.Button(top, text='Вернуть стандартное', command=self.reset_selected).pack(side='right', padx=8)
        paned = ttk.Panedwindow(win, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=12, pady=(0, 12))
        self.sheets = tk.Listbox(paned, font=('Arial', 11), activestyle='none', exportselection=False, width=24)
        names = list(self.params) + list(self.tables)
        for name in names:
            self.sheets.insert('end', name)
        self.sheets.bind('<<ListboxSelect>>', lambda e: self.show_sheet())
        paned.add(self.sheets, weight=0)
        right = ttk.Frame(paned)
        paned.add(right, weight=1)
        holder, self.tree = base.scrolled_tree(right, ('title', 'value', 'unit', 'hint'),
                                               ('Параметр', 'Значение', 'Ед.', 'Для чего / допустимое значение'),
                                               (300, 260, 60, -420), style='Small.Treeview')
        holder.pack(fill='both', expand=True)
        self.tree.tag_configure('changed', background='#FEF3C7')
        self.tree.bind('<Double-1>', self.edit_cell)
        self.current = None
        if names:
            self.sheets.selection_set(0)
            self.show_sheet()

    def show_sheet(self):
        sel = self.sheets.curselection()
        if not sel:
            return
        name = self.sheets.get(sel[0])
        self.current = name
        self.tree.delete(*self.tree.get_children())
        if name in self.params:
            self.tree.configure(columns=('title', 'value', 'unit', 'hint'), displaycolumns=('title', 'value', 'unit', 'hint'))
            for p in self.params[name]:
                value = self.changes.get((name, p.row), p.value)
                self.tree.insert('', 'end', iid=f'{p.row}', tags=('changed',) if (name, p.row) in self.changes else (),
                                 values=(p.title, '' if value is None else value, p.unit, f'{p.hint} {p.allowed}'.strip()))
        else:
            rows = self.tables[name]
            width = max((len(r) for r in rows), default=1)
            cols = tuple(f'c{i}' for i in range(1, width + 1))
            self.tree.configure(columns=cols, displaycolumns=cols)
            for i, c in enumerate(cols):
                head = rows[0][i] if rows and i < len(rows[0]) else ''
                self.tree.heading(c, text=str(head or f'Колонка {i + 1}')[:30])
                self.tree.column(c, width=160, stretch=True)
            for r, row in enumerate(rows[1:], 2):
                vals = [self.changes.get((name, r, c + 1), row[c] if c < len(row) else None) for c in range(width)]
                self.tree.insert('', 'end', iid=f'{r}', values=['' if v is None else v for v in vals])

    def edit_cell(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item or not column:
            return
        row, col = int(item), int(column[1:])
        name = self.current
        if name in self.params:
            if col != 2:
                return
            key = (name, row)
        else:
            key = (name, row, col)
        current = self.tree.set(item, self.tree['columns'][col - 1])
        value = AskValue(self.win, 'Значение', current).result
        if value is None:
            return
        if name in self.params:
            param = next(p for p in self.params[name] if p.row == row)
            value = self.coerce(param, value)
            if value is _BAD:
                return
        self.changes[key] = value
        self.show_sheet()

    def coerce(self, param, text):
        text = text.strip()
        if isinstance(param.value, (int, float)) and not isinstance(param.value, bool):
            try:
                num = float(text.replace(',', '.'))
            except ValueError:
                messagebox.showerror('Значение', f'Нужно число. Допустимо: {param.allowed}', parent=self.win)
                return _BAD
            return int(num) if num.is_integer() else num
        return text

    def reset_selected(self):
        name = self.current
        sel = self.tree.selection()
        if name not in self.params or not sel:
            return
        defaults = {p.key: p.value for p in self.defaults.get(name, []) if p.key}
        for item in sel:
            param = next(p for p in self.params[name] if p.row == int(item))
            if param.key in defaults:
                self.changes[(name, param.row)] = defaults[param.key]
        self.show_sheet()

    def save(self):
        if not self.changes:
            self.win.destroy()
            return
        try:
            backup = write_values(self.path, self.changes, self.mtime)
        except (SettingsError, OSError) as exc:
            messagebox.showerror('Настройки ОК', str(exc), parent=self.win)
            return
        messagebox.showinfo('Настройки ОК', f'Сохранено. Резервная копия:\n{backup}', parent=self.win)
        self.win.destroy()


_BAD = object()


class AskValue:
    def __init__(self, parent, title, value):
        self.result = None
        win = self.win = tk.Toplevel(parent)
        win.title(title)
        win.transient(parent)
        win.grab_set()
        self.text = tk.Text(win, width=70, height=6, wrap='word', font=('Arial', 11))
        self.text.insert('1.0', value)
        self.text.pack(padx=12, pady=12)
        bar = ttk.Frame(win)
        bar.pack(fill='x', padx=12, pady=(0, 12))
        ttk.Button(bar, text='Отмена', command=win.destroy).pack(side='right', padx=(8, 0))
        ttk.Button(bar, text='Готово', command=self.ok).pack(side='right')
        self.text.focus_set()
        win.wait_window()

    def ok(self):
        self.result = self.text.get('1.0', 'end').rstrip('\n')
        self.win.destroy()
