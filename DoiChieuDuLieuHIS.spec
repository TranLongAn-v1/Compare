# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = [('C:\\Users\\Long An\\AppData\\Local\\Python\\pythoncore-3.14-64\\tcl\\tcl8.6', '_tcl_data'), ('C:\\Users\\Long An\\AppData\\Local\\Python\\pythoncore-3.14-64\\tcl\\tk8.6', '_tk_data')]
binaries = []
hiddenimports = ['tkinter', '_tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.font', 'tkinter.constants']
datas += collect_data_files('tkinter')
hiddenimports += collect_submodules('tkinter')
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkcalendar')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=['F:\\up level\\PYTHON\\Compare\\hooks'],
    hooksconfig={},
    runtime_hooks=['F:\\up level\\PYTHON\\Compare\\rthook_tk.py'],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='DoiChieuDuLieuHIS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
