"""Windows EXE entry point for a separately editable SolidWorks VBA macro.
No SWP or configuration is embedded in the EXE. No retries of RunMacro2.
"""
from __future__ import annotations
import argparse
import ctypes
import json
import os
from pathlib import Path
import re
import sys
import traceback
import uuid

TITLE = 'КК — запуск макроса'
VERSION = '1.4'
CONFIG_NAME = 'KC_Start.json'

class LauncherError(Exception):
    pass

def app_directory():
    # PyInstaller onefile extracts code into _MEIPASS. The editable config must
    # be beside the REAL executable, not in that temporary directory or CWD.
    return Path(sys.executable if getattr(sys, 'frozen', False) else __file__).absolute().parent

def read_config(path):
    try:
        cfg = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    except Exception as exc:
        raise LauncherError(f'Не удалось прочитать настройки:\n{path}\n\n{exc}') from exc
    if not isinstance(cfg, dict):
        raise LauncherError('KC_Start.json должен содержать объект JSON.')
    name = cfg.get('macro_path')
    if not isinstance(name, str) or not name.strip():
        raise LauncherError('Укажите macro_path в KC_Start.json.')
    macro = Path(name)
    if not macro.is_absolute():
        macro = Path(path).absolute().parent / macro
    if macro.suffix.lower() != '.swp':
        raise LauncherError('macro_path должен указывать на настоящий файл .swp, не на TXT.')
    if not macro.is_file():
        raise LauncherError(f'Не найден VBA-макрос:\n{macro}\n\nПроверьте имя SWP, сеть и права доступа.')
    proc = cfg.get('procedure', 'Main')
    module = cfg.get('module', '')
    if not isinstance(proc, str) or not re.fullmatch(r'[A-Za-z_][A-Za-z_0-9]*', proc):
        raise LauncherError('Некорректное имя procedure в KC_Start.json.')
    if not isinstance(module, str) or any(c in module for c in '\r\n\t'):
        raise LauncherError('Некорректное имя module в KC_Start.json.')
    # Empty module is an explicit request for UNIQUE entry-point discovery.
    return dict(macro=macro, module=module.strip(), procedure=proc)

def select_entry(methods, configured_module, procedure):
    if methods is None:
        methods = []
    if isinstance(methods, str):
        methods = [methods]
    available = []
    for method in methods:
        text = str(method).strip()
        if '.' not in text:
            continue
        module, name = text.rsplit('.', 1)
        if module and name and (module, name) not in available:
            available.append((module, name))
    matches = [(m, p) for m, p in available
               if p.casefold() == procedure.casefold()
               and (not configured_module or m.casefold() == configured_module.casefold())]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        listed = ', '.join(f'{m}.{p}' for m, p in available) or '(пусто)'
        raise LauncherError(f'Не найдена процедура {configured_module + "." if configured_module else ""}{procedure} без аргументов.\n'
                            f'Доступные методы: {listed}\n\nСкомпилируйте и сохраните SWP. Проверьте module/procedure в KC_Start.json.')
    raise LauncherError('В SWP несколько подходящих процедур Main.\n'
                        'Укажите точное имя модуля в поле module файла KC_Start.json.\n'
                        + '\n'.join(f'{m}.{p}' for m, p in matches))

class Log:
    """Compatibility interface. No persistent launcher logs."""
    warning = ''
    def __init__(self, directory):
        self.path = None
    def write(self, text):
        pass
    def location(self):
        return ''

class ProgressChannel:
    """Per-run UTF-16 snapshot, readable while COM blocks in the worker thread.
    Publish only after acquiring SingleInstance. All files are temporary.
    """
    def __init__(self):
        import tempfile
        self.folder = tempfile.TemporaryDirectory(prefix='KC_progress_')
        self.path = Path(self.folder.name) / 'state.txt'
        self.nonce = uuid.uuid4().hex
        self.descriptor = Path(tempfile.gettempdir()) / 'KC_Start_progress_session.txt'
    def __enter__(self):
        try:
            self.descriptor.write_text('KC_PROGRESS_V1\n' + str(self.path) + '\n' + self.nonce + '\n', encoding='utf-16')
        except Exception:
            self.folder.cleanup()
            raise
        return self
    def read(self):
        return read_progress(self.path, self.nonce)
    def __exit__(self, *unused):
        try:
            if self.nonce in self.descriptor.read_text(encoding='utf-16'):
                self.descriptor.unlink(missing_ok=True)
        except OSError:
            pass
        self.folder.cleanup()

def read_progress(path, nonce):
    try:
        lines = Path(path).read_text(encoding='utf-16').splitlines()
        if len(lines) != 5 or lines[0] != nonce:
            return None
        status, completed, total, label = lines[1:]
        completed, total = int(completed), int(total)
        if status not in ('menu', 'running', 'done', 'stopped', 'error'):
            return None
        if total != 8 or not 0 <= completed <= total or not label:
            return None
        if status == 'done' and completed != total:
            return None
        return dict(status=status, completed=completed, total=total, label=label)
    except (OSError, UnicodeError, ValueError):
        # VBA may currently be rewriting the snapshot; retain the last good one.
        return None


def notify(message, error=False):
    if sys.stdout is not None:
        print(message)
    if os.name == 'nt':
        # A windowed build has no terminal: errors and check results stay visible.
        ctypes.windll.user32.MessageBoxW(None, str(message), TITLE, (0x10 if error else 0x40) | 0x10000)

class SingleInstance:
    def __init__(self):
        self.handle = None
    def __enter__(self):
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p]
        kernel.CreateMutexW.restype = ctypes.c_void_p
        self.kernel = kernel
        ctypes.set_last_error(0)
        self.handle = kernel.CreateMutexW(None, False, r'Local\KC_Start_SolidWorks_Macro')
        if not self.handle:
            raise LauncherError(f'Не создана защита повторного запуска: {ctypes.get_last_error()}')
        if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
            self.__exit__(None, None, None)
            raise LauncherError('Этот запускатель уже работает.\nЗавершите текущий макрос и закройте его диалоги.')
        return self
    def __exit__(self, *unused):
        if self.handle:
            self.kernel.CloseHandle.argtypes = [ctypes.c_void_p]
            self.kernel.CloseHandle.restype = ctypes.c_int
            self.kernel.CloseHandle(self.handle)
            self.handle = None

def use_solidworks(cfg, log, check_only=False, progress=None):
    progress = progress or (lambda text: None)
    # Imports intentionally inside this function: --check does not need COM.
    import pythoncom
    import pywintypes
    from win32com.client import VARIANT, dynamic
    pythoncom.CoInitialize()
    try:
        try:
            # Same ProgID -> IID -> ROT lookup used by pywin32 GetActiveObject.
            # pythoncom.CLSIDFromProgID does NOT exist.
            progress('Подключение к открытому SolidWorks…')
            log.write('CONNECT: pywintypes.IID + pythoncom.GetActiveObject')
            running = pythoncom.GetActiveObject(pywintypes.IID('SldWorks.Application'))
            running = running.QueryInterface(pythoncom.IID_IDispatch)
            sw = dynamic.Dispatch(running)  # No makepy/gen_py generation on the share.
        except Exception as exc:
            raise LauncherError('Не удалось подключиться к открытому SolidWorks.\n'
                                'Откройте SolidWorks и чертёж. Запускайте его и лаунчер под одним пользователем '
                                'и с одинаковым уровнем прав.\n\n' + str(exc)) from exc
        log.write('CONNECT OK; resolving application methods')
        # Explicitly distinguish methods from properties for late-bound COM.
        # Without this, GetType may be executed on attribute access and return
        # integer 3, so appending () then calls an integer a second time.
        sw._FlagAsMethod('GetMacroMethods', 'RunMacro2')
        progress('Проверка активного чертежа…')
        log.write('READ ActiveDoc')
        active = sw.ActiveDoc
        if active is None:
            raise LauncherError('В SolidWorks нет активного документа. Откройте нужный чертёж.')
        doc = dynamic.Dispatch(active._oleobj_)
        doc._FlagAsMethod('GetType', 'GetTitle')
        log.write('CALL GetType (explicit method)')
        doc_type = int(doc.GetType())
        log.write(f'Document type: {doc_type}')
        if doc_type != 3:  # swDocDRAWING
            raise LauncherError('Активный документ — не чертёж. Переключитесь на нужный SLDDRW.')
        log.write('CALL GetTitle (explicit method)')
        title = str(doc.GetTitle())
        log.write('Active drawing: ' + title)
        progress('Определение стартовой процедуры VBA-макроса…')
        module, procedure = cfg['module'], cfg['procedure']
        if not module:
            try:
                methods = sw.GetMacroMethods(str(cfg['macro']), 1)  # swMethodsWithoutArguments
            except Exception as exc:
                raise LauncherError('Не удалось определить модуль SWP автоматически.\n'
                                    'Укажите его имя в поле module файла KC_Start.json.\n\n' + str(exc)) from exc
            module, procedure = select_entry(methods, '', procedure)
        log.write(f'Entry: {module}.{procedure}; file: {cfg["macro"]}')
        if check_only:
            return f'Подключение к SolidWorks выполнено.\nЧертёж: {title}\nМакрос: {cfg["macro"]}\nПроцедура: {module}.{procedure}\n\nRunMacro2 НЕ вызывался: макрос не запускался.'
        error = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        progress('Выполняется макрос в SolidWorks.\nОжидание его завершения или вашего ответа в диалоге SolidWorks…')
        log.write('RunMacro2 BEGIN; unload_after_run=1; no automatic retry')
        raw = sw.RunMacro2(str(cfg['macro']), module, procedure, 1, error)
        ok = raw
        code = int(error.value)
        # Pywin32 invocation wrappers can expose an out argument as a tuple.
        if isinstance(raw, tuple):
            if len(raw) != 2:
                raise LauncherError(f'Неизвестный формат ответа RunMacro2: {raw!r}. Автоматический повтор запрещён.')
            ok, returned_code = raw
            code = int(returned_code)
        if not isinstance(ok, (bool, int)):
            raise LauncherError(f'Неизвестный результат RunMacro2: {raw!r}. Автоматический повтор запрещён.')
        log.write(f'RunMacro2 RETURN; result={raw!r}; code={code}')
        if not bool(ok) or code != 0:
            raise LauncherError(f'SolidWorks не подтвердил запуск макроса.\nRunMacro2: {ok}; код: {code}.\n\n'
                                'Проверьте компиляцию SWP, настройки module/procedure и разрешение запуска макросов. '
                                'Повторите запуск вручную только после проверки результата предыдущей попытки.')
        # True from RunMacro2 means invocation succeeded; it is NOT validation
        # that all operations inside VBA succeeded or that the user did not cancel.
        progress('SolidWorks вернул управление. Завершение запуска…')
        log.write('Invocation finished. Check the VBA macro report for placement/export results.')
        return None
    finally:
        pythoncom.CoUninitialize()

def execute_request(config_path, directory, check_only, events):
    """Worker entry: all COM objects are created and used in THIS thread.
    Only queue messages cross into the GUI thread; no Tk objects or COM objects.
    """
    log = None
    def progress(text):
        events.put(('stage', str(text)))
    try:
        progress('Подготовка файлов и журнала запуска…')
        log = Log(directory)
        progress('Чтение настроек и поиск VBA-макроса…')
        cfg = read_config(config_path)
        log.write(f'Config: {config_path}; macro: {cfg["macro"]}')
        events.put(('log', log.location()))
        with SingleInstance():
            with ProgressChannel() as channel:
                events.put(('channel', (channel.path, channel.nonce)))
                try:
                    message = use_solidworks(cfg, log, check_only, progress=progress)
                finally:
                    final = channel.read()
                    if final:
                        events.put(('macro', final))
        events.put(('done', dict(code=0, message=message or '', warning=log.warning,
                                location=log.location())))
    except Exception as exc:
        if log is not None:
            log.write(traceback.format_exc())
        location = log.location() if log is not None else 'Лог не создан.'
        events.put(('done', dict(code=2, message=str(exc), warning='', location=location)))

def format_elapsed(seconds):
    minutes, seconds = divmod(max(0, int(seconds)), 60)
    hours, minutes = divmod(minutes, 60)
    return f'{hours:02}:{minutes:02}:{seconds:02}'

def run_with_progress(config_path, directory, check_only=False):
    import queue
    import threading
    import time
    try:
        import tkinter as tk
        from tkinter import ttk
        root = tk.Tk()
    except Exception as exc:
        raise LauncherError('Не удалось открыть окно прогресса (Tkinter).\n' + str(exc)) from exc
    root.title(f'{TITLE} — {VERSION}')
    width, height = 420, 190
    left, top, right, bottom = 0, 0, root.winfo_screenwidth(), root.winfo_screenheight()
    if os.name == 'nt':
        from ctypes import wintypes
        area = wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(area), 0):
            left, top, right, bottom = area.left, area.top, area.right, area.bottom
    root.geometry(f'{width}x{height}+{max(left, right-width-12)}+{top+10}')
    root.resizable(False, False)
    root.attributes('-topmost', True)
    frame = ttk.Frame(root, padding=12)
    frame.pack(fill='both', expand=True)
    stage = tk.StringVar(value='Подготовка запуска…')
    ttk.Label(frame, textvariable=stage, wraplength=390, justify='left',
              font=('Segoe UI', 10)).pack(anchor='w')
    canvas = tk.Canvas(frame, height=20, highlightthickness=0, background='#f0f0f0')
    canvas.pack(fill='x', pady=(10, 6))
    detail = tk.StringVar(value='Ожидание сведений от макроса')
    ttk.Label(frame, textvariable=detail, wraplength=390).pack(anchor='w')
    elapsed = tk.StringVar(value='Время: 00:00:00')
    ttk.Label(frame, textvariable=elapsed).pack(anchor='w', pady=(4, 0))
    bottom_frame = ttk.Frame(frame)
    bottom_frame.pack(fill='x', pady=(6, 0))
    ttk.Label(bottom_frame, text='Ответы на вопросы — в окнах SW/Excel.').pack(side='left')
    ttk.Button(bottom_frame, text='Свернуть', command=root.iconify).pack(side='right')
    root.protocol('WM_DELETE_WINDOW', root.iconify)
    events = queue.Queue()
    result = dict(done=False, code=2, message='', warning='', location='')
    state = dict(channel=None, macro=None)
    started = time.monotonic()

    def draw_bar(event=None):
        canvas.delete('all')
        m = state['macro']
        completed = m['completed'] if m else 0
        active = completed if m and m['status'] == 'running' else -1
        w = max(100, canvas.winfo_width())
        gap = 4
        unit = (w - 7*gap) / 8
        for i in range(8):
            fill = '#4c8ead' if i < completed else '#e2e6e9'
            outline = '#4c8ead' if i == active else fill
            canvas.create_rectangle(i*(unit+gap), 2, i*(unit+gap)+unit, 18,
                                    fill=fill, outline=outline, width=2)

    def apply_macro(m):
        state['macro'] = m
        stage.set(m['label'])
        if m['status'] == 'running':
            detail.set(f"Этап {min(8, m['completed']+1)} из 8 · завершено {m['completed']}/8")
        elif m['status'] == 'done':
            detail.set('Завершено 8/8 этапов · проверьте итоговый отчёт')
        elif m['status'] == 'menu':
            detail.set('Ожидание выбора · шкала относится к режиму 1')
        else:
            detail.set(f"Подтверждено {m['completed']}/8 этапов · смотрите сообщение макроса")
        draw_bar()

    canvas.bind('<Configure>', draw_bar)
    worker = threading.Thread(target=execute_request,
                              args=(config_path, directory, check_only, events),
                              name='SolidWorks-COM-worker', daemon=False)
    def poll():
        elapsed.set('Время: ' + format_elapsed(time.monotonic() - started))
        try:
            while True:
                kind, value = events.get_nowait()
                if kind == 'channel':
                    state['channel'] = value
                elif kind == 'macro':
                    apply_macro(value)
                elif kind == 'stage' and state['macro'] is None:
                    stage.set(value)
                elif kind == 'done':
                    result.update(value)
                    result['done'] = True
                    if value['code'] != 0:
                        stage.set('Выполнение остановлено. Смотрите описание ошибки.')
                    elif check_only:
                        stage.set('Проверка подключения завершена. Макрос не запускался.')
                    elif state['macro'] is None:
                        stage.set('Вызов завершён. Результат — в сообщении макроса.')
                        detail.set('Макрос не передал поэтапный прогресс')
                    # Never infer 100% or VBA success from RunMacro2=True.
                    root.after(1100, root.destroy)
                    return
        except queue.Empty:
            pass
        if state['channel']:
            snapshot = read_progress(*state['channel'])
            if snapshot and snapshot != state['macro']:
                apply_macro(snapshot)
        root.after(150, poll)
    root.after(150, worker.start)
    root.after(100, poll)
    root.mainloop()
    if not result['done']:
        raise LauncherError('Окно закрыто извне. Макрос мог продолжить работу. Проверьте SolidWorks перед повторным запуском.')
    if result['code'] != 0:
        notify(result['message'], error=True)
    elif result['message']:
        notify(result['message'])
    elif result['warning']:
        notify(result['warning'])
    return result['code']

def main(argv=None):
    parser = argparse.ArgumentParser(description='Запуск редактируемого SWP через открытый SolidWorks')
    parser.add_argument('--config', help='Путь к внешнему KC_Start.json')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('--check', action='store_true', help='Проверить JSON и наличие SWP, без подключения к SW')
    group.add_argument('--check-sw', action='store_true', help='Проверить подключение/точку входа, не выполнять макрос')
    args = parser.parse_args(argv)
    directory = app_directory()
    path = Path(args.config).absolute() if args.config else directory / CONFIG_NAME
    log = None
    try:
        if os.name != 'nt':
            raise LauncherError('KC_Start предназначен для Windows с установленным SolidWorks.')
        if not args.check:
            return run_with_progress(path, directory, args.check_sw)
        log = Log(directory)
        cfg = read_config(path)
        log.write(f'Config: {path}; macro: {cfg["macro"]}')
        notify(f'KC_Start {VERSION}\nНастройки прочитаны, SWP найден.\n{cfg["macro"]}\n\n'
               'Подключение к SolidWorks и выполнение VBA ещё не проверены.\n\n' + log.location())
        return 0
    except Exception as exc:
        if log is not None:
            log.write(traceback.format_exc())
        notify(str(exc) + ('\n\n' + log.location() if log is not None else ''), error=True)
        return 2

if __name__ == '__main__':
    raise SystemExit(main())
