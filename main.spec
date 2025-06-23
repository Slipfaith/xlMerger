# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

project_dir = os.path.abspath(os.path.dirname(__file__))

a = Analysis(
    ['main.py'],
    pathex=[project_dir],  # Рабочая папка, где лежит main.py
    binaries=[],
    datas=[],
    hiddenimports=collect_submodules('gui') + collect_submodules('core'),  # чтобы все подмодули подтянул
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=r"C:\Users\yanismik\Desktop\PythonProject1\xlM_2.0\xlM2.0.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='main',
)
