# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app_icon.ico', '.'),
        ('config.ini', '.'),
    ],
    hiddenimports=[
        'keyboard',
        'pystray',
        'pywin32',
        'win32api',
        'win32con',
        'win32gui',
        'win32process',
        'ttkbootstrap',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='app_icon.ico'
)
