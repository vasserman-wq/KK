"""Главное окно: шапка, шаги слева, строка состояния, фоновые задачи, эмулятор Excel."""
from __future__ import annotations

import functools
import logging
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .. import __version__, config
from ..card import Card, default_card_path
from ..logs import log_path
from ..paths import resource
from ..xlshim import Factory, ShimError
from . import base

log = logging.getLogger(__name__)
STEPS = (
    ('drawing', 'Чертёж'),
    ('markers', 'Разметка'),
    ('balloons', 'Баллоны и карта'),
    ('header', 'Реквизиты и подписи'),
    ('card', 'Карта контроля'),
    ('sheets', 'Листы ОК'),
)
HELP = (
    'Работа идёт по шагам слева:\n'
    ' 1. Чертёж — подключение к SOLIDWORKS и рабочая карта рядом с чертежом.\n'
    ' 2. Разметка — какие размеры и заметки включать в карту (вместо # и @ в чертеже).\n'
    ' 3. Баллоны и карта — запуск режима 1 макроса 1.\n'
    ' 4. Реквизиты и подписи — номер КД, наименование, масса, оборудование, подписи, исполнения.\n'
    ' 5. Карта контроля — первые три строки ОК1, вводные пункты и параметры в одной таблице.\n'
    ' 6. Листы ОК — запуск макроса 2.\n\n'
    'Все данные хранятся в одном файле рядом с чертежом: «<чертёж> — карта контроля.xlsm». '
    'Это та же рабочая книга, что раньше; Excel для работы не нужен.\n\n'
    'Клавиши: Ctrl+S — сохранить карту, Ctrl+1…Ctrl+6 — шаги, F1 — справка.'
)
SHORTCUTS = {83: 'save', 49: 1, 50: 2, 51: 3, 52: 4, 53: 5, 54: 6}  # коды физических клавиш Windows


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.cfg = config.load()
        self.factory = Factory()
        self.server = None
        self.drawing = ''          # полный путь активного чертежа
        self.card = None           # kk.card.Card или None
        self.busy = False
        self.events = queue.Queue()
        self.status = tk.StringVar(value='')
        self.title_var = tk.StringVar(value='Чертёж не выбран')
        self.save_state = tk.StringVar(value='')
        root.title('Карта контроля KK')
        root.geometry('1280x820')
        root.minsize(1000, 640)
        self._set_icon()
        base.configure_styles(root)
        self._build_menu()
        self._build_header()
        self._build_statusbar()
        body = ttk.Frame(root)
        body.pack(side='top', fill='both', expand=True)
        self._build_sidebar(body)
        self.content = ttk.Frame(body)
        self.content.pack(side='left', fill='both', expand=True)
        from .pages_main import BalloonsPage, DrawingPage, SheetsPage
        from .pages_markers import MarkersPage
        from .pages_card import CardPage, HeaderPage
        self.pages = {}
        for number, (key, cls) in enumerate((('drawing', DrawingPage), ('markers', MarkersPage),
                                             ('balloons', BalloonsPage), ('header', HeaderPage),
                                             ('card', CardPage), ('sheets', SheetsPage)), 1):
            self.pages[key] = cls(self, number)
        root.protocol('WM_DELETE_WINDOW', self.close_app)
        root.bind_all('<Key>', self._on_key, add='+')
        root.bind_all('<F1>', lambda e: self.show_help())
        self.show_step('drawing')
        root.after(150, self._poll)

    # --- каркас
    def _set_icon(self):
        icon = resource('app.ico')
        if icon.exists() and os.name == 'nt':
            try:
                self.root.iconbitmap(default=str(icon))
            except tk.TclError:
                pass

    def _build_menu(self):
        menu = tk.Menu(self.root)
        project = tk.Menu(menu, tearoff=False)
        project.add_command(label='Открыть карту…', command=self.open_card_dialog)
        project.add_command(label='Сохранить карту', command=self.save_card, accelerator='Ctrl+S')
        project.add_separator()
        project.add_command(label='Выход', command=self.close_app)
        menu.add_cascade(label='Проект', menu=project)
        steps = tk.Menu(menu, tearoff=False)
        for number, (key, title) in enumerate(STEPS, 1):
            steps.add_command(label=f'{number}. {title}', accelerator=f'Ctrl+{number}',
                              command=functools.partial(self.show_step, key))
        menu.add_cascade(label='Шаги', menu=steps)
        service = tk.Menu(menu, tearoff=False)
        service.add_command(label='Пути к макросам…', command=self.edit_paths)
        service.add_command(label='Настройки ОК…', command=self.edit_settings)
        service.add_separator()
        service.add_command(label='Журнал эмулятора Excel', command=self.show_shim_log)
        service.add_command(label='Папка журнала программы', command=lambda: self.open_path(log_path().parent))
        menu.add_cascade(label='Сервис', menu=service)
        helpm = tk.Menu(menu, tearoff=False)
        helpm.add_command(label='Как пользоваться', command=self.show_help, accelerator='F1')
        helpm.add_command(label='О программе', command=lambda: messagebox.showinfo(
            'О программе', f'Карта контроля KK {__version__}\nМакросы SOLIDWORKS + карта без Excel.'))
        menu.add_cascade(label='Справка', menu=helpm)
        self.root.configure(menu=menu)

    def _build_header(self):
        header = tk.Frame(self.root, background=base.HEADER_BG, padx=14, pady=8)
        header.pack(side='top', fill='x')
        tk.Label(header, text='Карта контроля', background=base.HEADER_BG, foreground='#FFFFFF',
                 font=('Arial', 13, 'bold')).pack(side='left')
        tk.Label(header, text='│', background=base.HEADER_BG, foreground='#6B7280', font=('Arial', 13)).pack(side='left', padx=10)
        tk.Label(header, textvariable=self.title_var, background=base.HEADER_BG, foreground='#E5E7EB',
                 font=('Arial', 12)).pack(side='left')
        tk.Label(header, textvariable=self.save_state, background=base.HEADER_BG, foreground='#FDE68A',
                 font=('Arial', 10)).pack(side='left', padx=(16, 0))
        options = dict(background='#374151', foreground='#FFFFFF', activebackground='#4B5563',
                       activeforeground='#FFFFFF', disabledforeground='#6B7280', relief='flat', borderwidth=0,
                       highlightthickness=0, font=('Arial', 10, 'bold'), padx=12, pady=5, cursor='hand2')
        self.save_button = tk.Button(header, text='Сохранить карту', command=self.save_card, **options)
        self.save_button.pack(side='right')
        base.tip(self.save_button, 'Записать рабочую карту рядом с чертежом (Ctrl+S).')

    def _build_statusbar(self):
        bar = tk.Frame(self.root, background='#E5E7EB', padx=12, pady=4)
        bar.pack(side='bottom', fill='x')
        self.progress = ttk.Progressbar(bar, mode='determinate', length=220, maximum=8)
        self.progress.pack(side='right')
        tk.Label(bar, textvariable=self.status, background='#E5E7EB', foreground='#374151',
                 font=('Arial', 10), anchor='w').pack(side='left', fill='x', expand=True)

    def _build_sidebar(self, body):
        side = tk.Frame(body, background=base.SIDEBAR_BG, width=260)
        side.pack(side='left', fill='y')
        side.pack_propagate(False)
        tk.Label(side, text='ШАГИ', background=base.SIDEBAR_BG, foreground='#6B7280',
                 font=('Arial', 9, 'bold'), anchor='w').pack(fill='x', padx=18, pady=(18, 6))
        self.step_buttons = {}
        for number, (key, title) in enumerate(STEPS, 1):
            button = tk.Button(side, text=f'{number}. {title}', command=functools.partial(self.show_step, key),
                               anchor='w', justify='left', font=('Arial', 11), relief='flat', borderwidth=0,
                               highlightthickness=0, padx=14, pady=8, background=base.SIDEBAR_BG,
                               activebackground=base.SIDEBAR_ACTIVE, foreground=base.STEP_COLORS['idle'], cursor='hand2')
            button.pack(fill='x', padx=8, pady=2)
            self.step_buttons[key] = button
        tk.Frame(side, background=base.SIDEBAR_BG).pack(fill='both', expand=True)
        tk.Button(side, text='? Как пользоваться (F1)', command=self.show_help, anchor='w', font=('Arial', 10),
                  relief='flat', borderwidth=0, padx=14, pady=5, background=base.SIDEBAR_BG,
                  activebackground='#E5E7EB', foreground='#374151', cursor='hand2').pack(fill='x', padx=8, pady=(0, 14))

    def show_step(self, key):
        if self.busy and key != getattr(self, 'current', None):
            self.status.set('Дождитесь окончания текущей операции.')
            return
        for name, page in self.pages.items():
            page.frame.pack_forget()
            self.step_buttons[name].configure(background=base.SIDEBAR_BG)
        self.current = key
        self.pages[key].frame.pack(fill='both', expand=True)
        self.step_buttons[key].configure(background=base.SIDEBAR_ACTIVE)
        self.pages[key].on_show()

    def mark_step(self, key, state):
        self.step_buttons[key].configure(foreground=base.STEP_COLORS.get(state, base.STEP_COLORS['idle']))

    def show_help(self):
        messagebox.showinfo('Как пользоваться', HELP)

    def _on_key(self, event):
        if not (event.state & 0x4):
            return None
        action = SHORTCUTS.get(event.keycode) if os.name == 'nt' else {'s': 'save'}.get(event.keysym.lower())
        if action == 'save':
            self.save_card()
            return 'break'
        if isinstance(action, int) and action <= len(STEPS):
            self.show_step(STEPS[action - 1][0])
            return 'break'
        return None

    # --- фоновые задачи (SOLIDWORKS, макросы): окно не замирает
    def run_bg(self, label, func, on_done=None, on_error=None):
        if self.busy:
            messagebox.showwarning('Занято', 'Дождитесь окончания текущей операции.')
            return
        self.busy = True
        self.status.set(label)

        def work():
            try:
                result = func()
            except Exception as exc:  # любая ошибка показывается пользователю, окно не падает
                log.exception('%s: ошибка', label)
                self.events.put(('error', exc, on_error))
            else:
                self.events.put(('done', result, on_done))
        threading.Thread(target=work, name='KK-worker', daemon=True).start()

    def _poll(self):
        try:
            while True:
                kind, payload, callback = self.events.get_nowait()
                self.busy = False
                if kind == 'error':
                    self.status.set('Ошибка: ' + str(payload).splitlines()[0][:150])
                    if callback:
                        callback(payload)
                    else:
                        messagebox.showerror('Ошибка', str(payload))
                elif callback:
                    callback(payload)
        except queue.Empty:
            pass
        for page in self.pages.values():
            page.tick()
        self.root.after(150, self._poll)

    # --- эмулятор Excel для макросов
    def ensure_server(self):
        if self.server is None:
            from ..dialogs import WinDialogs
            from ..xlshim.server import ShimServer
            self.factory.dialogs = WinDialogs(int(self.root.winfo_id()) if os.name == 'nt' else 0)
            self.server = ShimServer(self.factory)
        self.server.start()

    def show_shim_log(self):
        lines = self.factory.log[-300:]
        text = '\n'.join(lines) if lines else 'Замечаний нет: макросы не обращались к неподдержанным возможностям Excel.'
        win = tk.Toplevel(self.root)
        win.title('Журнал эмулятора Excel')
        win.geometry('820x480')
        box = tk.Text(win, wrap='word', font=('Consolas', 10))
        box.insert('1.0', text)
        box.configure(state='disabled')
        box.pack(fill='both', expand=True)

    # --- чертёж и карта
    def set_drawing(self, path):
        self.drawing = str(path or '')
        if 'markers' in getattr(self, 'pages', {}):
            self.pages['markers'].rows = []
            self.pages['markers'].dirty = False
        card_path = default_card_path(self.drawing) if self.drawing else None
        self.close_card(ask=False)
        if card_path is not None and card_path.exists():
            self.load_card(card_path, quiet=True)
        self.refresh_title()

    def load_card(self, path, quiet=False):
        try:
            card = Card(path, Factory())
        except (ShimError, OSError) as exc:
            if not quiet:
                messagebox.showerror('Карта не открыта', str(exc))
            log.warning('Карта не открыта: %s', exc)
            return None
        owner = card.drawing_path
        if self.drawing and owner and os.path.normcase(owner) != os.path.normcase(self.drawing):
            messagebox.showwarning('Другой чертёж', f'Карта привязана к другому чертежу:\n{owner}\n\n'
                                   'Макросы откажутся работать с ней для текущего чертежа (защитная проверка).')
        self.card = card
        self.refresh_title()
        return card

    def close_card(self, ask=True):
        if self.card is None:
            return True
        if ask and self.card.dirty:
            answer = messagebox.askyesnocancel('Несохранённые изменения', 'Сохранить изменения карты?')
            if answer is None:
                return False
            if answer and not self.save_card():
                return False
        self.card.close()
        self.card = None
        self.refresh_title()
        return True

    def reload_card(self):
        path = self.card.path if self.card else (default_card_path(self.drawing) if self.drawing else None)
        if self.card:
            self.card.close()
            self.card = None
        if path is not None and Path(path).exists():
            self.load_card(path, quiet=True)
        self.refresh_title()

    def save_card(self):
        if self.card is None:
            return True
        for page in self.pages.values():
            if not page.commit():
                return False
        try:
            self.card.save()
        except (ShimError, OSError) as exc:
            messagebox.showerror('Не сохранено', f'Карта не записана: {exc}\nЕсли файл открыт в Excel, закройте его.')
            return False
        self.status.set(f'Карта сохранена: {self.card.path.name}')
        self.refresh_title()
        return True

    def open_card_dialog(self):
        path = filedialog.askopenfilename(title='Рабочая карта контроля', filetypes=[('Карта контроля', '*.xlsm *.xlsx')])
        if path and self.close_card():
            self.load_card(path)
            self.pages[self.current].on_show()

    def refresh_title(self):
        name = Path(self.drawing).name if self.drawing else 'Чертёж не выбран'
        state = f' · карта: {self.card.state or "нет состояния"}' if self.card else (' · карты пока нет' if self.drawing else '')
        self.title_var.set(name + state)
        self.save_state.set('● Есть несохранённые изменения' if self.card and self.card.dirty else '')

    # --- сервис
    def edit_paths(self):
        from .settings_dialog import PathsDialog
        PathsDialog(self)

    def edit_settings(self):
        from .settings_dialog import SettingsDialog
        SettingsDialog(self)

    @staticmethod
    def open_path(path):
        if os.name == 'nt':
            os.startfile(str(path))  # noqa: S606

    def close_app(self):
        if self.busy and not messagebox.askyesno('Операция идёт',
                                                 'Макрос ещё работает. Закрытие программы прервёт его работу с картой. Закрыть?'):
            return
        if not self.close_card():
            return
        if self.server:
            self.server.stop()
        self.root.destroy()


def run_gui():
    root = tk.Tk()
    MainWindow(root)
    root.mainloop()
