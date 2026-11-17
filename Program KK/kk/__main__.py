"""Точка входа: py -m kk или KK.exe."""
import logging
import sys


def main():
    from .logs import setup_logging
    setup_logging()
    log = logging.getLogger('kk')
    log.info('Запуск KK (Python %s)', sys.version.split()[0])
    if '--check' in sys.argv:
        from . import __version__
        print(f'KK {__version__}: импорт модулей в порядке')
        return 0
    try:
        import tkinter  # noqa: F401
    except ImportError:
        print('Нет tkinter: установите Python с компонентом tcl/tk.')
        return 2
    from .gui.app import run_gui
    run_gui()
    return 0


if __name__ == '__main__':
    sys.exit(main())
