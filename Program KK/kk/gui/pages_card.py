"""Шаги 4 и 5: реквизиты и подписи; карта контроля одной таблицей."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from ..card import strip_number, to_minutes
from ..xlshim import ShimError
from . import base
from .page import Page

KIND_NAMES = {'intro': 'вводный', 'param': 'из чертежа', 'stamp': 'клеймо ОТК'}


def _no_card(frame):
    label = ttk.Label(frame, text='Карты пока нет. Выполните шаг 3 «Баллоны и карта» или откройте существующую карту.',
                      style='Title.TLabel', wraplength=800)
    return label


class HeaderPage(Page):
    key = 'header'
    title = 'Реквизиты и подписи'
    hint = ('Реквизиты изделия, масса, оборудование ОК1, должности и фамилии для шапки ОК1/КЭ1, таблица исполнений. '
            'Номер КД, наименование, масса и фамилия разработчика обязательны для макроса 2.')
    FIELDS = (('kd_number', 'Номер КД *'), ('product_name', 'Наименование изделия *'),
              ('primary_application', 'Первичное применение'), ('mass', 'Масса детали, кг *'),
              ('equipment', 'Оборудование ОК1'))

    def build(self):
        base.page_nav(self.frame, back=('← Назад', self.go('balloons')), forward=('Далее: карта контроля →', self.next_step))
        self.body = ttk.Frame(self.frame, padding=(24, 8))
        self.body.pack(fill='both', expand=True)
        self.empty = _no_card(self.body)
        self.form = ttk.Frame(self.body)
        self.vars = {key: tk.StringVar() for key, _ in self.FIELDS}
        box = ttk.LabelFrame(self.form, text='Реквизиты изделия', padding=12, style='Section.TLabelframe')
        box.pack(fill='x')
        box.columnconfigure(1, weight=1)
        for row, (key, label) in enumerate(self.FIELDS):
            base.labeled_entry(box, row, label, self.vars[key])
        signs = ttk.LabelFrame(self.form, text='Шапка ОК1 / КЭ1 — должность и фамилия', padding=12, style='Section.TLabelframe')
        signs.pack(fill='x', pady=10)
        self.roles = [tk.StringVar() for _ in range(5)]
        self.names = [tk.StringVar() for _ in range(5)]
        ttk.Label(signs, text='Должность', style='Hint.TLabel').grid(row=0, column=0, sticky='w')
        ttk.Label(signs, text='Фамилия', style='Hint.TLabel').grid(row=0, column=1, sticky='w', padx=8)
        for i in range(5):
            ttk.Entry(signs, textvariable=self.roles[i], width=24, font=('Arial', 11)).grid(row=i + 1, column=0, pady=2)
            ttk.Entry(signs, textvariable=self.names[i], width=34, font=('Arial', 11)).grid(row=i + 1, column=1, padx=8, pady=2)
        ttk.Label(signs, text='Первая строка — «Разраб.», фамилия обязательна.', style='Hint.TLabel').grid(
            row=6, column=0, columnspan=2, sticky='w', pady=(4, 0))
        var = ttk.LabelFrame(self.form, text='Исполнения', padding=12, style='Section.TLabelframe')
        var.pack(fill='x')
        self.variant_on = tk.BooleanVar()
        ttk.Checkbutton(var, text='Добавить таблицу исполнений в рамки ОК / КЭ', variable=self.variant_on).pack(anchor='w')
        self.variant_text = tk.Text(var, height=3, wrap='word', font=('Arial', 11), relief='solid', borderwidth=1)
        self.variant_text.pack(fill='x', pady=(6, 0))
        ttk.Label(var, text='Флажок выключен — собственная таблица исполнений удаляется при заполнении.',
                  style='Hint.TLabel').pack(anchor='w')
        self.loaded = None

    def on_show(self):
        card = self.app.card
        if card is None:
            self.form.pack_forget()
            self.empty.pack(anchor='w')
            self.loaded = None
            return
        self.empty.pack_forget()
        self.form.pack(fill='both', expand=True)
        h = card.header()
        for key, _ in self.FIELDS:
            self.vars[key].set(h.get(key, ''))
        for i, (role, name) in enumerate(h['signatures'][:5]):
            self.roles[i].set(role)
            self.names[i].set(name)
        self.variant_on.set(h['variant_enabled'].strip().upper() in ('TRUE', 'ИСТИНА', '1', 'ДА'))
        self.variant_text.delete('1.0', 'end')
        self.variant_text.insert('1.0', h['variant_text'])
        self.loaded = card

    def commit(self):
        card = self.app.card
        if card is None or self.loaded is not card:
            return True
        text = self.variant_text.get('1.0', 'end').strip()
        if self.variant_on.get() and not text:
            messagebox.showerror('Исполнения', 'Включена таблица исполнений: заполните текст.')
            return False
        mass = self.vars['mass'].get().strip()
        if mass and to_minutes(mass) is None:
            messagebox.showerror('Масса', 'Масса должна быть числом, например 12,5.')
            return False
        before = card.header()
        values = {key: self.vars[key].get().strip() for key, _ in self.FIELDS}
        values['signatures'] = [(r.get().strip(), n.get().strip()) for r, n in zip(self.roles, self.names)]
        values['variant_enabled'] = 'TRUE' if self.variant_on.get() else 'FALSE'
        values['variant_text'] = text
        changed = {k: v for k, v in values.items() if before.get(k) != v}
        if 'variant_enabled' in changed and before['variant_enabled'] == '' and not self.variant_on.get():
            changed.pop('variant_enabled')
        if changed:
            try:
                card.set_header(**changed)
            except ShimError as exc:
                messagebox.showerror('Реквизиты', str(exc))
                return False
            self.app.refresh_title()
        return True

    def next_step(self):
        if self.commit():
            self.app.show_step('card')


class BlockDialog:
    """Правка пункта: текст, до 4 средств ТО из списка, объём, время. Номер не редактируется."""

    def __init__(self, page, block):
        self.page, self.block = page, block
        card = page.app.card
        win = self.win = tk.Toplevel(page.app.root)
        win.title(f'Пункт {block.key}' if block.key != 'новый' else 'Новый вводный пункт')
        win.geometry('760x500')
        win.transient(page.app.root)
        win.grab_set()
        frm = ttk.Frame(win, padding=16)
        frm.pack(fill='both', expand=True)
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text='Номер').grid(row=0, column=0, sticky='w')
        ttk.Label(frm, text=f'{block.key}  ({KIND_NAMES[block.kind]}; номер защищён от ручного изменения)',
                  foreground=base.LOCKED_FG).grid(row=0, column=1, sticky='w')
        ttk.Label(frm, text='Контролируемый параметр').grid(row=1, column=0, sticky='nw', pady=(8, 0))
        self.text = tk.Text(frm, height=5, wrap='word', font=('Arial', 11), relief='solid', borderwidth=1)
        self.text.grid(row=1, column=1, sticky='we', pady=(8, 0))
        self.text.insert('1.0', strip_number(block.text))
        ttk.Label(frm, text='Номер «N.» в начале добавляется автоматически.', style='Hint.TLabel').grid(row=2, column=1, sticky='w')
        tools = [''] + card.tools_list()
        self.tool_vars = []
        for i in range(block.rows):
            ttk.Label(frm, text='Средство ТО' if i == 0 else f'Доп. средство {i}').grid(row=3 + i, column=0, sticky='w', pady=3)
            var = tk.StringVar(value=block.tools[i] if i < len(block.tools) else '')
            values = tools if var.get() in tools else tools + [var.get()]
            ttk.Combobox(frm, textvariable=var, values=values, state='readonly', font=('Arial', 11)).grid(
                row=3 + i, column=1, sticky='we', pady=3)
            self.tool_vars.append(var)
        r = 3 + block.rows
        self.volume = tk.StringVar(value=block.volume or '100%')
        self.minutes = tk.StringVar(value='' if block.minutes is None else f'{block.minutes:g}')
        base.labeled_entry(frm, r, 'Объём и ПК', self.volume, width=12, hint='например 100% или 50%')
        base.labeled_entry(frm, r + 1, 'Время То/Тв, мин', self.minutes, width=12, hint='число, например 1,5')
        bar = ttk.Frame(frm)
        bar.grid(row=r + 2, column=0, columnspan=3, sticky='e', pady=(16, 0))
        ttk.Button(bar, text='Отмена', command=win.destroy).pack(side='right', padx=(8, 0))
        base.primary_button(bar, 'Готово', self.apply, base.GREEN, base.GREEN_ACTIVE).pack(side='right')
        win.bind('<Escape>', lambda e: win.destroy())
        self.text.focus_set()

    def apply(self):
        text = ' '.join(self.text.get('1.0', 'end').split())
        if not text:
            messagebox.showerror('Пункт', 'Текст пункта не может быть пустым.', parent=self.win)
            return
        try:
            self.page.app.card.update_block(self.block, text=text, tools=[v.get() for v in self.tool_vars],
                                            volume=self.volume.get(), minutes=self.minutes.get().strip())
        except ShimError as exc:
            messagebox.showerror('Пункт', str(exc), parent=self.win)
            return
        self.win.destroy()
        self.page.fill(select_row=self.block.row)


class CardPage(Page):
    key = 'card'
    title = 'Карта контроля'
    hint = ('Одна таблица, как в Excel: вводные пункты, параметры чертежа, клеймо ОТК. Двойной щелчок — правка. '
            'Номера баллонов защищены. Добавление или удаление вводного пункта меняет нумерацию: '
            'после него снова выполните шаг 3.')

    def build(self):
        base.page_nav(self.frame, back=('← Назад', self.go('header')), forward=('Далее: листы ОК →', self.next_step))
        self.body = ttk.Frame(self.frame, padding=(24, 4))
        self.body.pack(fill='both', expand=True)
        self.empty = _no_card(self.body)
        self.view = ttk.Frame(self.body)
        fixed = ttk.LabelFrame(self.view, text='Первые три строки ОК1', padding=8, style='Section.TLabelframe')
        fixed.pack(fill='x')
        fixed.columnconfigure(1, weight=1)
        self.fixed = [tk.StringVar() for _ in range(3)]
        for i in range(3):
            base.labeled_entry(fixed, i, f'{i + 1:02d}', self.fixed[i], width=100)
        bar = ttk.Frame(self.view)
        bar.pack(fill='x', pady=6)
        ttk.Button(bar, text='✎ Изменить пункт', command=self.edit).pack(side='left')
        ttk.Button(bar, text='＋ Вводный пункт', command=self.add_intro).pack(side='left', padx=6)
        ttk.Button(bar, text='− Удалить вводный', command=self.delete_intro).pack(side='left')
        self.stamp_button = ttk.Button(bar, text='＋ Клеймо ОТК', command=self.add_stamp)
        self.stamp_button.pack(side='left', padx=6)
        base.tip(self.stamp_button, 'Запись без номера: номер назначается автоматически как последний свободный.')
        base.primary_button(bar, 'Сохранить карту', self.app.save_card, base.GREEN, base.GREEN_ACTIVE).pack(side='right')
        holder, self.tree = base.scrolled_tree(self.view, ('key', 'kind', 'text', 'tools', 'volume', 'minutes'),
                                               ('Р', 'Тип', 'Контролируемые параметры', 'Наименование средств ТО', 'Объём и ПК', 'То/Тв'),
                                               (60, 110, -520, 330, 90, 70))
        holder.pack(fill='both', expand=True)
        self.tree.configure(selectmode='browse')
        self.tree.tag_configure('intro', background='#F0F9FF')
        self.tree.tag_configure('stamp', background='#FEF3C7')
        self.tree.tag_configure('review', background=base.REVIEW_BG)
        self.tree.bind('<Double-1>', lambda e: self.edit())
        self.tree.bind('<Return>', lambda e: self.edit())
        self.totals = tk.StringVar(value='')
        ttk.Label(self.view, textvariable=self.totals, style='Strong.TLabel').pack(anchor='w', pady=(6, 0))
        self.blocks = []
        self.loaded = None

    def on_show(self):
        card = self.app.card
        if card is None:
            self.view.pack_forget()
            self.empty.pack(anchor='w')
            self.loaded = None
            return
        self.empty.pack_forget()
        self.view.pack(fill='both', expand=True)
        for var, text in zip(self.fixed, card.header()['fixed_rows']):
            var.set(text)
        self.loaded = card
        self.fill()

    def fill(self, select_row=None):
        card = self.app.card
        self.blocks = card.blocks()
        self.tree.delete(*self.tree.get_children())
        for i, b in enumerate(self.blocks):
            warn = not b.tools or b.minutes is None
            tag = 'review' if warn and strip_number(b.text) else (b.kind if b.kind != 'param' else '')
            self.tree.insert('', 'end', iid=str(i), tags=(tag,) if tag else (),
                             values=(b.key, KIND_NAMES[b.kind], b.text, '; '.join(b.tools), b.volume,
                                     '' if b.minutes is None else f'{b.minutes:g}'))
            if b.row == select_row:
                self.tree.selection_set(str(i))
                self.tree.see(str(i))
        t = card.totals(self.blocks)
        issues = card.problems()
        self.totals.set(f'Пунктов: {t["items"]} (вводных {t["intro"]}, из чертежа {t["param"]}, клеймо {t["stamp"]}) · '
                        f'Итого время: {t["minutes"]:g} мин'
                        + (f' · без времени: {t["no_time"]}' if t['no_time'] else '')
                        + (f' · замечаний: {len(issues)}' if issues else '')
                        + f' · состояние {card.state}')
        self.stamp_button.state(['disabled'] if t['stamp'] else ['!disabled'])
        self.app.refresh_title()

    def selected(self):
        sel = self.tree.selection()
        return self.blocks[int(sel[0])] if sel else None

    def edit(self):
        block = self.selected()
        if block is not None:
            BlockDialog(self, block)

    def add_intro(self):
        block = self.selected()
        try:
            row = self.app.card.add_intro(after=block if block is not None and block.kind == 'intro' else None)
        except ShimError as exc:
            messagebox.showerror('Вводный пункт', str(exc))
            return
        self.fill(select_row=row)
        new = self.selected()
        if new is not None:
            BlockDialog(self, new)

    def delete_intro(self):
        block = self.selected()
        if block is None:
            return
        if block.kind != 'intro':
            messagebox.showinfo('Удаление', 'Удалять можно только вводный пункт. Параметры чертежа убираются на шаге 2.')
            return
        if not messagebox.askyesno('Удаление', 'Удалить выбранный пункт со всеми продолжениями и инструментами?'):
            return
        self.app.card.delete_intro(block)
        self.fill()

    def add_stamp(self):
        try:
            number = self.app.card.add_stamp()
        except ShimError as exc:
            messagebox.showinfo('Клеймо ОТК', str(exc))
            return
        self.fill()
        self.app.status.set(f'Добавлено клеймо ОТК, пункт {number}. Укажите время.')

    def commit(self):
        card = self.app.card
        if card is None or self.loaded is not card:
            return True
        current = card.header()['fixed_rows']
        values = [v.get().strip() for v in self.fixed]
        if values != current:
            card.set_header(fixed_rows=values)
        return True

    def next_step(self):
        if self.commit():
            self.app.show_step('sheets')
