# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['file_converter.py'],
    pathex=[],
    binaries=[],
    datas=[('poppler', 'poppler')],
    hiddenimports=['pandas', 'openpyxl', 'matplotlib', 'PIL', 'pdf2image', 'img2pdf', 'pypdf'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    name='FileConverter',
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
