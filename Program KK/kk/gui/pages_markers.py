"""Шаг 2: разметка. Пользователь решает, что включать в карту; # и @ вручную не вводятся.

По умолчанию: размер включён, если у него нет #; заметка включена, если у неё есть @.
Допуски и сварные швы макрос включает всегда (как раньше), в списке их нет.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from .. import decisions, sw
from . import base
from .page import Page

KINDS = {'DIM': 'Размер', 'NOTE': 'Заметка'}


class MarkersPage(Page):
    key = 'markers'
    title = 'Разметка'
    hint = ('Отметьте, что включать в карту контроля. Щелчок по строке подсвечивает объект в SOLIDWORKS, '
            'двойной щелчок или Пробел — «Включить в карту / Не ставить баллон». Знаки # и @ в чертёж вводить не нужно.')

    def build(self):
        base.page_nav(self.frame, back=('← Назад', self.go('drawing')),
                      forward=('Далее: баллоны и карта →', self.next_step))
        body = ttk.Frame(self.frame, padding=(24, 4))
        body.pack(fill='both', expand=True)
        bar = ttk.Frame(body)
        bar.pack(fill='x', pady=(0, 6))
        base.primary_button(bar, 'Прочитать чертёж', self.load_from_drawing).pack(side='left')
        ttk.Button(bar, text='☑ Включить выбранные', command=lambda: self.set_selected(True)).pack(side='left', padx=(12, 4))
        ttk.Button(bar, text='☐ Не включать выбранные', command=lambda: self.set_selected(False)).pack(side='left', padx=4)
        ttk.Button(bar, text='Показать в SOLIDWORKS', command=self.highlight).pack(side='left', padx=4)
        self.save_button = base.primary_button(bar, 'Сохранить разметку', self.save, base.GREEN, base.GREEN_ACTIVE)
        self.save_button.pack(side='right')
        flt = ttk.Frame(body)
        flt.pack(fill='x', pady=(0, 6))
        ttk.Label(flt, text='Показать:').pack(side='left')
        self.show_dims = tk.BooleanVar(value=True)
        self.show_notes = tk.BooleanVar(value=True)
        self.only_off = tk.BooleanVar(value=False)
        for text, var in (('размеры', self.show_dims), ('заметки', self.show_notes), ('только не включённые', self.only_off)):
            ttk.Checkbutton(flt, text=text, variable=var, command=self.fill).pack(side='left', padx=6)
        ttk.Label(flt, text='Поиск:').pack(side='left', padx=(18, 4))
        self.search = tk.StringVar()
        self.search.trace_add('write', lambda *a: self.fill())
        ttk.Entry(flt, textvariable=self.search, width=30).pack(side='left')
        holder, self.tree = base.scrolled_tree(body, ('inc', 'kind', 'sheet', 'view', 'name', 'text'),
                                               ('В карту', 'Тип', 'Лист', 'Вид', 'Имя', 'Текст'),
                                               (80, 90, 90, 160, 90, -520))
        holder.pack(fill='both', expand=True)
        self.tree.tag_configure('off', foreground=base.LOCKED_FG)
        self.tree.bind('<Double-1>', lambda e: self.toggle())
        self.tree.bind('<space>', lambda e: (self.toggle(), 'break')[1])
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.summary = tk.StringVar(value='')
        ttk.Label(body, textvariable=self.summary, style='Hint.TLabel').pack(anchor='w', pady=(6, 0))
        self.rows = []
        self.dirty = False
        self.auto_highlight = tk.BooleanVar(value=True)
        ttk.Checkbutton(body, text='Подсвечивать выбранную строку в SOLIDWORKS', variable=self.auto_highlight).pack(anchor='w')

    # --- данные
    def on_show(self):
        if not self.rows and self.app.drawing:
            self.rows = decisions.load(self.app.card, self.app.drawing)
            self.fill()

    def load_from_drawing(self):
        if not self.app.drawing:
            messagebox.showwarning('Чертёж', 'Сначала подключитесь к SOLIDWORKS на шаге 1.')
            return
        saved = {decisions.row_key(r): r['include'] for r in decisions.load(self.app.card, self.app.drawing)}

        def work():
            with sw.session() as swapp:
                doc, path, _ = sw.active_drawing(swapp)
                if path.lower() != self.app.drawing.lower():
                    raise sw.SwError('Активный чертёж не совпадает с выбранным на шаге 1.')
                return sw.candidates(swapp, doc)

        def done(cands):
            rows = []
            for c in cands:
                row = {'sheet': c.sheet, 'view': c.view, 'name': c.name, 'kind': c.kind, 'text': c.text,
                       'x': c.x, 'y': c.y, 'include': c.default_include, 'marker': c.marker}
                row['include'] = saved.get(decisions.row_key(row), row['include'])
                rows.append(row)
            self.rows = rows
            self.dirty = True
            self.fill()
            self.app.status.set(f'Прочитано объектов: {len(rows)}. Не забудьте сохранить разметку.')

        self.app.run_bg('Чтение размеров и заметок чертежа…', work, done)

    def visible_rows(self):
        q = self.search.get().strip().lower()
        for i, r in enumerate(self.rows):
            if r['kind'] == 'DIM' and not self.show_dims.get() or r['kind'] == 'NOTE' and not self.show_notes.get():
                continue
            if self.only_off.get() and r['include']:
                continue
            if q and q not in f"{r['sheet']} {r['view']} {r['name']} {r['text']}".lower():
                continue
            yield i, r

    def fill(self):
        self.tree.delete(*self.tree.get_children())
        for i, r in self.visible_rows():
            self.tree.insert('', 'end', iid=str(i), tags=() if r['include'] else ('off',),
                             values=('☑ да' if r['include'] else '☐ нет', KINDS.get(r['kind'], r['kind']), r['sheet'],
                                     r['view'], r['name'], r['text']))
        on = sum(1 for r in self.rows if r['include'])
        self.summary.set(f'Всего объектов: {len(self.rows)} · в карту: {on} · не включать: {len(self.rows) - on}'
                         + (' · ● есть несохранённые изменения' if self.dirty else ''))

    def _set(self, indexes, value):
        for i in indexes:
            self.rows[i]['include'] = value if value is not None else not self.rows[i]['include']
        self.dirty = True
        selection = self.tree.selection()
        self.fill()
        keep = [s for s in selection if self.tree.exists(s)]
        if keep:
            self.tree.selection_set(keep)

    def toggle(self):
        self._set([int(s) for s in self.tree.selection()], None)

    def set_selected(self, value):
        self._set([int(s) for s in self.tree.selection()], value)

    def on_select(self, _event=None):
        if self.auto_highlight.get() and len(self.tree.selection()) == 1 and not self.app.busy:
            self.highlight(quiet=True)

    def highlight(self, quiet=False):
        sel = self.tree.selection()
        if not sel:
            return
        r = self.rows[int(sel[0])]
        cand = sw.Candidate(r['sheet'], r['view'], r['name'], r['kind'], r['text'], r['x'], r['y'], False)

        def work():
            with sw.session() as swapp:
                doc, _, _ = sw.active_drawing(swapp)
                return sw.highlight(swapp, doc, cand)

        def done(ok):
            self.app.status.set('Объект выделен в SOLIDWORKS.' if ok else 'Объект не найден на листе (возможно, сдвинут). Прочитайте чертёж заново.')

        def failed(exc):
            self.app.status.set(f'Подсветка не выполнена: {exc}')
            if not quiet:
                messagebox.showwarning('Подсветка', str(exc))

        self.app.run_bg('Подсветка в SOLIDWORKS…', work, done, failed)

    def save(self):
        if not self.app.drawing:
            return False
        where = decisions.save(self.app.card, self.app.drawing, [{k: v for k, v in r.items() if k != 'marker'} for r in self.rows])
        if where == 'card' and not self.app.save_card():
            return False
        self.dirty = False
        self.fill()
        self.app.mark_step('markers', 'ok')
        self.app.status.set('Разметка сохранена ' + ('в рабочую карту.' if where == 'card' else 'временно; попадёт в карту после шага 3.'))
        return True

    def commit(self):
        if self.dirty and self.rows and self.app.card is not None:
            decisions.save(self.app.card, self.app.drawing, [{k: v for k, v in r.items() if k != 'marker'} for r in self.rows])
            self.dirty = False
        return True

    def next_step(self):
        if self.dirty:
            answer = messagebox.askyesnocancel('Разметка', 'Сохранить разметку перед переходом?')
            if answer is None or (answer and not self.save()):
                return
        self.app.show_step('balloons')
