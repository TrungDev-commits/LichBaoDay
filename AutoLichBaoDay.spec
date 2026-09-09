# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

docxcompose_datas = collect_data_files('docxcompose')

a = Analysis(
    ['run_desktop.py'],
    pathex=['backend'],
    binaries=[],
    datas=[
        ('frontend', 'frontend'), 
        ('app/templates', 'app/templates'), 
        ('backend', 'backend')
    ] + docxcompose_datas,
    hiddenimports=[
        'docx',
        'docxcompose',
        'docxcompose.composer',
        'docxcompose.properties',
        'docxcompose.templates',
        'lxml',
        'lxml.etree',
        'uvicorn',
        'uvicorn.logging',
        'uvicorn.loops',
        'uvicorn.loops.auto',
        'uvicorn.protocols',
        'uvicorn.protocols.http',
        'uvicorn.protocols.http.auto',
        'fastapi',
        'starlette'
    ],
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
    name='LHT_TSS',
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
