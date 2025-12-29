# -*- mode: python ; coding: utf-8 -*-

# Explicit hidden imports to ensure PyInstaller bundles optional/externally-detected packages
hidden_imports = [
    'psutil',
    'PIL',
    'PIL.Image',
    'PIL.ImageDraw',
    'pystray',
    'pystray._win32',
    'win32api',
    'win32con',
    'win32gui',
    'win32process',
    'pywintypes',
    'PyQt5',
    'PyQt5.QtWidgets',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.sip',
    'configparser',
]

from PyInstaller.utils.hooks import collect_submodules
import importlib.util
import os

# include any submodules for packages that may use dynamic imports
hidden_imports += collect_submodules('psutil')

# Try to locate psutil extension binaries (.pyd/.dll) in the active environment
psutil_binaries = []
ps_spec = importlib.util.find_spec('psutil')
if ps_spec and getattr(ps_spec, 'origin', None):
    ps_dir = os.path.dirname(ps_spec.origin)
    if os.path.isdir(ps_dir):
        for fname in os.listdir(ps_dir):
            if fname.lower().endswith(('.pyd', '.dll')):
                src = os.path.join(ps_dir, fname)
                # include in root of the bundle
                psutil_binaries.append((src, '.'))


a = Analysis(
    ['dd2-keypresser.py'],
    pathex=[],
    binaries=psutil_binaries,
    datas=[('config.ini', '.'), ('app_icon.ico', '.')],
    hiddenimports=hidden_imports,
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
    name='dd2-keypresser',
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
    icon=['app_icon.ico'],
)
