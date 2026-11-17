#!/usr/bin/env python3
"""Запуск программы из исходников: py KK.py (Python 3.10+, Windows)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from kk.__main__ import main  # noqa: E402

if __name__ == '__main__':
    sys.exit(main())
