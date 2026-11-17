"""Оформление и общие виджеты. Цвета, шрифты и каркас взяты из passport_macro/gui.py."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

HEADER_BG = '#1F2937'
SIDEBAR_BG = '#F3F4F6'
SIDEBAR_ACTIVE = '#DBEAFE'
GREEN, GREEN_ACTIVE = '#C8E6C9', '#AED6B0'
BLUE, BLUE_ACTIVE = '#DBEAFE', '#BFDBFE'
STEP_COLORS = {'ok': '#166534', 'warn': '#92400E', 'idle': '#1F2937', 'error': '#B91C1C'}
REVIEW_BG = '#fff0d6'
LOCKED_FG = '#6B7280'


def configure_styles(root):
    style = ttk.Style(root)
    style.configure('Treeview', rowheight=32, font=('Arial', 11))
    style.configure('Treeview.Heading', font=('Arial', 11, 'bold'))
    style.configure('Small.Treeview', rowheight=26, font=('Arial', 10))
    style.configure('Big.TLabel', font=('Arial', 22, 'bold'), foreground='#111827')
    style.configure('Title.TLabel', font=('Arial', 18, 'bold'), foreground='#111827')
    style.configure('Step.TLabel', font=('Arial', 10), foreground='#6B7280')
    style.configure('Hint.TLabel', font=('Arial', 10), foreground='#4B5563')
    style.configure('Strong.TLabel', font=('Arial', 12, 'bold'), foreground='#111827')
    style.configure('Ok.TLabel', font=('Arial', 11, 'bold'), foreground='#15803D')
    style.configure('Warn.TLabel', font=('Arial', 11, 'bold'), foreground='#B45309')
    style.configure('Error.TLabel', font=('Arial', 10, 'bold'), foreground='#B91C1C')
    style.configure('Section.TLabelframe.Label', font=('Arial', 11, 'bold'), foreground='#1D4ED8')
    style.configure('TButton', font=('Arial', 10))
    style.configure('TLabel', font=('Arial', 10))
    style.configure('TCheckbutton', font=('Arial', 10))
    return style


class Tooltip:
    """Всплывающая подсказка (как ui_widgets.Tooltip в passport_macro)."""

    def __init__(self, widget, text, delay=550):
        self.widget, self.text, self.delay = widget, text, delay
        self.window = self.after_id = None
        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self._hide, add='+')
        widget.bind('<ButtonPress>', self._hide, add='+')

    def _schedule(self, _event=None):
        self._cancel()
        self.after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self.after_id is not None:
            try:
                self.widget.after_cancel(self.after_id)
            except tk.TclError:
                pass
            self.after_id = None

    def _show(self):
        self.after_id = None
        if self.window is not None or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
            self.window = tk.Toplevel(self.widget)
            self.window.wm_overrideredirect(True)
            self.window.wm_geometry(f'+{x}+{y}')
            tk.Label(self.window, text=self.text, justify='left', wraplength=380, background='#FFFBEA',
                     foreground='#111827', relief='solid', borderwidth=1, font=('Arial', 10),
                     padx=8, pady=5).pack()
        except tk.TclError:
            self.window = None

    def _hide(self, _event=None):
        self._cancel()
        if self.window is not None:
            try:
                self.window.destroy()
            except tk.TclError:
                pass
            self.window = None


def tip(widget, text):
    Tooltip(widget, text)
    return widget


def primary_button(parent, text, command, color=BLUE, active=BLUE_ACTIVE):
    return tk.Button(parent, text=text, command=command, background=color, activebackground=active,
                     foreground='#111827', font=('Arial', 11, 'bold'), relief='solid', borderwidth=1,
                     padx=16, pady=6, cursor='hand2')


def page_frame(parent, number, total, title, hint, on_help=None):
    page = ttk.Frame(parent)
    head = ttk.Frame(page, padding=(24, 16, 24, 8))
    head.pack(side='top', fill='x')
    if number:
        ttk.Label(head, text=f'Шаг {number} из {total}', style='Step.TLabel').pack(anchor='w')
    row = ttk.Frame(head)
    row.pack(fill='x')
    ttk.Label(row, text=title, style='Title.TLabel').pack(side='left')
    if on_help:
        ttk.Button(row, text='? Подсказка', command=on_help).pack(side='right')
    ttk.Label(head, text=hint, style='Hint.TLabel', wraplength=880, justify='left').pack(anchor='w', pady=(4, 0))
    return page


def page_nav(page, back=None, forward=None):
    """Кнопки «Назад» / «Далее» внизу шага: (текст, команда)."""
    nav = ttk.Frame(page, padding=(24, 8, 24, 12))
    nav.pack(side='bottom', fill='x')
    ttk.Separator(page, orient='horizontal').pack(side='bottom', fill='x')
    if back:
        ttk.Button(nav, text=back[0], command=back[1]).pack(side='left')
    if forward:
        primary_button(nav, forward[0], forward[1]).pack(side='right')
    return nav


def scrolled_tree(parent, columns, headings, widths, style=None, height=12):
    holder = ttk.Frame(parent)
    tree = ttk.Treeview(holder, columns=columns, show='headings', selectmode='extended', height=height,
                        **({'style': style} if style else {}))
    for col, head, width in zip(columns, headings, widths):
        tree.heading(col, text=head)
        stretch = width < 0
        tree.column(col, width=abs(width), minwidth=40, stretch=stretch, anchor='w')
    bar = ttk.Scrollbar(holder, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=bar.set)
    tree.pack(side='left', fill='both', expand=True)
    bar.pack(side='right', fill='y')
    return holder, tree


def labeled_entry(parent, row, label, var, width=48, hint=None):
    ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', padx=(0, 10), pady=4)
    entry = ttk.Entry(parent, textvariable=var, width=width, font=('Arial', 11))
    entry.grid(row=row, column=1, sticky='we', pady=4)
    if hint:
        ttk.Label(parent, text=hint, style='Hint.TLabel').grid(row=row, column=2, sticky='w', padx=(10, 0))
    return entry
