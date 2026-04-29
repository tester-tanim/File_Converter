# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['pandas', 'openpyxl', 'matplotlib', 'matplotlib.backends.backend_agg', 'PIL', 'PIL.Image', 'PIL.ImageOps', 'pdf2image', 'img2pdf', 'pypdf', 'docx', 'reportlab', 'reportlab.platypus', 'reportlab.lib.styles', 'reportlab.lib.pagesizes', 'reportlab.lib.units', 'reportlab.lib.colors', 'reportlab.lib.enums']
hiddenimports += collect_submodules('reportlab')
hiddenimports += collect_submodules('docx')


a = Analysis(
    ['file_converter.py'],
    pathex=[],
    binaries=[],
    datas=[('poppler', 'poppler')],
    hiddenimports=hiddenimports,
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
