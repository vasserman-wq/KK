#!/usr/bin/env python3
"""PDF -> editable checklist -> Word passport. Python 3.10+.

Точка запуска. Код программы — в папке passport_macro рядом с этим файлом.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from passport_macro import *  # noqa: E402,F401,F403  — совместимость со старыми импортами
from passport_macro.gui import main, run_gui  # noqa: E402,F401

if __name__ == "__main__":
    main()
