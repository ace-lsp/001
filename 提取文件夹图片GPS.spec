# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\ACER-PETER-LEE\\source\\repos\\Solution1\\提取文件夹图片GPS.py'],
    pathex=['C:\\Users\\ACER-PETER-LEE\\source\\repos\\Solution1'],
    binaries=[],
    datas=[],
    hiddenimports=['PIL._imaging','pandas','openpyxl', 'pandas._libs.tslibs.np_datetime'], # 添加隐藏导入
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    # cipher=block_cipher,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='提取文件夹图片GPS',
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
