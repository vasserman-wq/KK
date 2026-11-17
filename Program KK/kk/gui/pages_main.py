"""Шаги 1, 3 и 6: чертёж, запуск режима 1 макроса 1, запуск макроса 2."""
from __future__ import annotations

import logging
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from .. import decisions, sw
from ..card import default_card_path
from ..settings_file import SettingsError, ensure_shared, shared_path
from . import base
from .page import Page

log = logging.getLogger(__name__)


def run_macro_session(app, macro, module, procedure, progress=None, mode=None, rows=None):
    """В фоновом потоке: запустить макрос в SOLIDWORKS, когда эмулятор Excel уже запущен."""
    factory = app.factory
    with sw.session() as swapp:
        _doc, path, _title = sw.active_drawing(swapp)
        if app.drawing and Path(path).resolve() != Path(app.drawing).resolve():
            raise sw.SwError(f'Активный чертёж сменился:\n{path}\nОжидался:\n{app.drawing}\nОткройте нужный чертёж и нажмите «Обновить» на шаге 1.')
        factory.commands.clear()
        factory.clear_decisions()
        if rows:
            factory.set_decisions(path, decisions.to_factory_rows(rows))
        if mode:
            factory.set_command('mode', mode)
        start = len(factory.log)
        try:
            if progress:
                with progress:
                    used = sw.run_macro(swapp, macro, module, procedure)
                    state = progress.read()
            else:
                used, state = sw.run_macro(swapp, macro, module, procedure), None
        finally:
            factory.commands.clear()
            factory.clear_decisions()
        return {'module': used, 'state': state, 'shim_log': factory.log[start:]}


class DrawingPage(Page):
    key = 'drawing'
    title = 'Чертёж'
    hint = ('Откройте чертёж в SOLIDWORKS и нажмите «Обновить». Рабочая карта хранится одним файлом рядом с чертежом; '
            'старые карты .xlsm открываются без переноса.')

    def build(self):
        base.page_nav(self.frame, forward=('Далее: разметка →', self.go('markers')))
        body = ttk.Frame(self.frame, padding=(24, 8))
        body.pack(fill='both', expand=True)
        box = ttk.LabelFrame(body, text='SOLIDWORKS', padding=12, style='Section.TLabelframe')
        box.pack(fill='x')
        self.sw_text = tk.StringVar(value='Не подключено')
        ttk.Label(box, textvariable=self.sw_text, style='Strong.TLabel', wraplength=900).pack(anchor='w')
        row = ttk.Frame(box)
        row.pack(anchor='w', pady=(8, 0))
        base.primary_button(row, 'Обновить из SOLIDWORKS', self.refresh).pack(side='left')
        ttk.Button(row, text='Открыть папку чертежа', command=self.open_folder).pack(side='left', padx=8)
        card = ttk.LabelFrame(body, text='Рабочая карта', padding=12, style='Section.TLabelframe')
        card.pack(fill='x', pady=12)
        self.card_text = tk.StringVar(value='')
        ttk.Label(card, textvariable=self.card_text, wraplength=900, justify='left').pack(anchor='w')
        ttk.Button(card, text='Открыть другую карту…', command=self.app.open_card_dialog).pack(anchor='w', pady=(8, 0))
        setup = ttk.LabelFrame(body, text='Макросы и настройки', padding=12, style='Section.TLabelframe')
        setup.pack(fill='x')
        self.setup_text = tk.StringVar(value='')
        ttk.Label(setup, textvariable=self.setup_text, wraplength=900, justify='left').pack(anchor='w')
        row = ttk.Frame(setup)
        row.pack(anchor='w', pady=(8, 0))
        ttk.Button(row, text='Пути к макросам…', command=self.app.edit_paths).pack(side='left')
        ttk.Button(row, text='Настройки ОК…', command=self.app.edit_settings).pack(side='left', padx=8)

    def on_show(self):
        self.update_texts()

    def update_texts(self):
        a = self.app
        if a.drawing:
            card_path = default_card_path(a.drawing)
            if a.card:
                t = a.card.totals()
                self.card_text.set(f'Файл: {a.card.path}\nСостояние: {a.card.state or "—"} · пунктов: {t["items"]} '
                                   f'(вводных {t["intro"]}) · время {t["minutes"]:g} мин')
                a.mark_step('drawing', 'ok')
            else:
                self.card_text.set(f'Карты пока нет. Она будет создана на шаге 3:\n{card_path}')
                a.mark_step('drawing', 'warn')
        else:
            self.card_text.set('Сначала подключитесь к SOLIDWORKS.')
        cfg = a.cfg
        lines = [f'Макрос 1: {cfg["macro1"] or "не указан"}', f'Макрос 2: {cfg["macro2"] or "не указан"}']
        settings = shared_path(cfg['macro1'])
        if settings:
            lines.append(f'Настройки: {settings}' + ('' if settings.exists() else ' — файла нет, будет создан из встроенных'))
        self.setup_text.set('\n'.join(lines))

    def refresh(self):
        def work():
            with sw.session() as swapp:
                _doc, path, title = sw.active_drawing(swapp)
                return path, title, ''

        def done(result):
            path, title, rev = result
            if self.app.drawing != path and not self.app.close_card():
                return
            self.app.set_drawing(path)
            self.sw_text.set(f'Активный чертёж: {title}\n{path}')
            self.app.status.set('Подключено к SOLIDWORKS.')
            self.update_texts()

        self.app.run_bg('Подключение к SOLIDWORKS…', work, done)

    def open_folder(self):
        if self.app.drawing:
            self.app.open_path(Path(self.app.drawing).parent)


class MacroPage(Page):
    """Общая часть шагов с запуском макроса."""
    which = 1

    def build_run_area(self, body, button_text):
        self.info = tk.StringVar(value='')
        ttk.Label(body, textvariable=self.info, wraplength=900, justify='left').pack(anchor='w')
        self.run_button = base.primary_button(body, button_text, self.start, base.GREEN, base.GREEN_ACTIVE)
        self.run_button.pack(anchor='w', pady=12)
        self.stage = tk.StringVar(value='')
        ttk.Label(body, textvariable=self.stage, style='Strong.TLabel', wraplength=900).pack(anchor='w')
        self.result = tk.Text(body, height=12, wrap='word', font=('Arial', 10), relief='solid', borderwidth=1)
        self.result.pack(fill='both', expand=True, pady=(8, 0))
        self.progress = None
        self.started = 0.0

    def macro_config(self):
        cfg = self.app.cfg
        n = self.which
        return cfg[f'macro{n}'], cfg[f'module{n}'], cfg[f'procedure{n}'] or 'Main'

    def check_common(self):
        macro, _, _ = self.macro_config()
        try:
            sw.check_macro_file(macro)
        except sw.SwError as exc:
            messagebox.showerror('Макрос', str(exc))
            return False
        try:
            path, created = ensure_shared(self.app.cfg['macro1'])
            if created:
                messagebox.showinfo('Настройки', f'Рядом с макросом не было файла настроек. Создан из встроенных значений:\n{path}')
        except (SettingsError, OSError) as exc:
            messagebox.showerror('Настройки', f'Файл настроек недоступен: {exc}')
            return False
        if not self.app.drawing:
            messagebox.showwarning('Чертёж', 'Сначала подключитесь к SOLIDWORKS на шаге 1.')
            return False
        try:
            self.app.ensure_server()
        except Exception as exc:  # не Windows, нет pywin32, ошибка регистрации
            messagebox.showerror('Эмулятор Excel', f'Не удалось запустить эмулятор Excel для макросов:\n{exc}')
            return False
        return True

    def tick(self):
        if self.progress is not None and self.app.busy:
            state = self.progress.read()
            elapsed = int(time.monotonic() - self.started)
            if state:
                status, done, total, label = state
                self.app.progress.configure(value=done)
                self.stage.set(f'{label} · этап {done} из {total} · {elapsed // 60}:{elapsed % 60:02d}')
            else:
                self.stage.set(f'Макрос работает · {elapsed // 60}:{elapsed % 60:02d}. Отвечайте на вопросы в окне SOLIDWORKS.')

    def show_result(self, text):
        self.result.configure(state='normal')
        self.result.delete('1.0', 'end')
        self.result.insert('1.0', text)
        self.result.configure(state='disabled')

    def launch(self, mode=None, rows=None, after=None):
        macro, module, procedure = self.macro_config()
        self.progress = sw.Progress() if self.which == 1 else None
        self.started = time.monotonic()
        self.app.progress.configure(value=0)
        self.run_button.configure(state='disabled')
        self.show_result('')

        def done(result):
            self.run_button.configure(state='normal')
            state = result.get('state')
            lines = []
            if state:
                status, done_, total, label = state
                lines.append(f'Итог макроса: {label} (состояние {status}, этап {done_} из {total}).')
            else:
                lines.append('Макрос завершил работу. Проверьте сообщения SOLIDWORKS.')
            if result.get('shim_log'):
                lines.append('\nЗамечания эмулятора Excel (передайте разработчику, если что-то не так):')
                lines.extend(result['shim_log'][-60:])
            self.stage.set('')
            self.progress = None
            self.app.reload_card()
            if after:
                after(result, lines)
            self.show_result('\n'.join(lines))
            self.app.status.set('Макрос завершён.')

        def failed(exc):
            self.run_button.configure(state='normal')
            self.progress = None
            self.stage.set('')
            self.app.reload_card()
            self.show_result(f'Ошибка: {exc}\n\n' + '\n'.join(self.app.factory.log[-40:]))
            messagebox.showerror('Макрос', str(exc))

        self.app.run_bg('Запуск макроса в SOLIDWORKS…',
                        lambda: run_macro_session(self.app, macro, module, procedure, self.progress, mode, rows),
                        done, failed)


class BalloonsPage(MacroPage):
    key = 'balloons'
    title = 'Баллоны и карта'
    hint = ('Запуск режима 1 макроса 1: недостающие баллоны, перенумерация и рабочая карта. '
            'Разметка из шага 2 передаётся макросу напрямую, знаки # и @ в чертёж не записываются. '
            'На вопросы макроса отвечайте в окне SOLIDWORKS.')
    which = 1

    def build(self):
        base.page_nav(self.frame, back=('← Назад', self.go('markers')), forward=('Далее: реквизиты →', self.go('header')))
        body = ttk.Frame(self.frame, padding=(24, 8))
        body.pack(fill='both', expand=True)
        self.build_run_area(body, '▶ Запустить режим 1')

    def on_show(self):
        a = self.app
        rows = decisions.load(a.card, a.drawing) if a.drawing else []
        off = sum(1 for r in rows if not r['include'])
        on = len(rows) - off
        self.info.set((f'Разметка: {on} объектов включить, {off} не включать.' if rows else
                       'Разметка в программе не сохранена: макрос возьмёт маркеры # и @ из чертежа, как раньше.')
                      + '\nПеред запуском карта сохраняется и закрывается в программе, после — открывается снова.')

    def start(self):
        if not self.check_common():
            return
        a = self.app
        rows = decisions.load(a.card, a.drawing)
        if not a.close_card():
            return

        def after(result, lines):
            state = result.get('state')
            if a.card is not None and rows and not a.card.book.has_sheet(decisions.SHEET):
                decisions.save(a.card, a.drawing, rows)
                a.card.save()
                lines.append('Разметка сохранена в рабочую карту.')
            ok = bool(state and state[0] == 'done')
            a.mark_step('balloons', 'ok' if ok else 'warn')
            if a.card is not None:
                lines.append(f'Карта: {a.card.path.name}, состояние {a.card.state}.')

        self.launch(mode='1', rows=rows, after=after)


class SheetsPage(MacroPage):
    key = 'sheets'
    title = 'Листы ОК'
    hint = ('Запуск макроса 2: листы ОК и таблицы в чертеже по рабочей карте. '
            'Перед запуском программа проверяет реквизиты и состояние карты (READY).')
    which = 2

    def build(self):
        base.page_nav(self.frame, back=('← Назад', self.go('card')))
        body = ttk.Frame(self.frame, padding=(24, 8))
        body.pack(fill='both', expand=True)
        self.build_run_area(body, '▶ Запустить макрос 2')

    def on_show(self):
        card = self.app.card
        if card is None:
            self.info.set('Карты пока нет: выполните шаг 3.')
            return
        issues = card.problems(for_macro2=True)
        errors = [m for lvl, m in issues if lvl == 'error']
        warns = [m for lvl, m in issues if lvl == 'warning']
        text = ['✔ Карта готова к построению листов ОК.'] if not errors else ['✖ Перед запуском исправьте:'] + errors[:12]
        if warns:
            text.append(f'⚠ Проверьте: {len(warns)} замечаний (первое: {warns[0]})')
        self.info.set('\n'.join(text))
        self.app.mark_step('sheets', 'idle' if not errors else 'warn')

    def start(self):
        a = self.app
        if a.card is None:
            messagebox.showwarning('Карта', 'Карты пока нет: выполните шаг 3.')
            return
        if not a.save_card():
            return
        errors = [m for lvl, m in a.card.problems(for_macro2=True) if lvl == 'error']
        if errors:
            messagebox.showerror('Карта не готова', '\n'.join(errors[:15]))
            return
        if not self.check_common() or not a.close_card():
            return
        self.launch(after=lambda result, lines: a.mark_step('sheets', 'ok'))
