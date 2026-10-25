# -*- mode: python ; coding: utf-8 -*-
# 精简打包：仅 openpyxl + tkinter，排除 pandas/numpy 等无关依赖

from PyInstaller.utils.hooks import collect_all

datas = [('templates', 'templates')]
binaries = []
hiddenimports = ['openpyxl.cell._writer', 'reconcile']

tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

a = Analysis(
    ['reconcile_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas',
        'numpy',
        'scipy',
        'matplotlib',
        'PIL',
        'Pillow',
        'IPython',
        'notebook',
        'pytest',
        'sqlite3',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='订单发票比对工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='订单发票比对工具',
)
