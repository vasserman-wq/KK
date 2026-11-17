# PyInstaller: сборка KK.exe (запускается из build.bat).
# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

root = Path(SPECPATH)
datas = [(str(p), 'kk/resources') for p in (root / 'kk' / 'resources').glob('*') if p.is_file()]

a = Analysis(
    [str(root / 'KK.py')],
    pathex=[str(root)],
    datas=datas,
    hiddenimports=['win32com.server.policy', 'win32com.server.util', 'win32com.server.exception',
                   'win32com.client.dynamic', 'pythoncom', 'pywintypes', 'win32gui', 'win32api', 'win32con',
                   'openpyxl.cell._writer'],
    excludes=['cv2', 'numpy', 'fitz', 'pymupdf', 'matplotlib'],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='KK',
    console=False,
    icon=str(root / 'kk' / 'resources' / 'app.ico') if (root / 'kk' / 'resources' / 'app.ico').exists() else None,
)
