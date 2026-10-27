# -*- mode: python ; coding: utf-8 -*-
# 精简打包：openpyxl + pywebview(Edge WebView2) + 本地 ui.html

from PyInstaller.utils.hooks import collect_all

datas = [
    ('templates', 'templates'),
    ('ui.html', '.'),
]
binaries = []
hiddenimports = ['openpyxl.cell._writer', 'reconcile', 'webview', 'clr_loader']

tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

try:
    wv = collect_all('webview')
    datas += wv[0]
    binaries += wv[1]
    hiddenimports += wv[2]
except Exception:
    pass

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
    upx=False,  # UPX 压缩极易触发杀软误报，保持关闭
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
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
