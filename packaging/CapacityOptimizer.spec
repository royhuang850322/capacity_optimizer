# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys

from PyInstaller.utils.hooks import collect_dynamic_libs, collect_submodules, copy_metadata

project_root = Path(SPECPATH).resolve().parents[0]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from build_support.packaging_manifest import (
    APP_NAME,
    DYNAMIC_LIB_PACKAGES,
    ENTRY_SCRIPT,
    HIDDEN_IMPORT_PACKAGES,
    METADATA_PACKAGES,
    iter_data_mappings,
)

datas = iter_data_mappings(project_root)
for package_name in METADATA_PACKAGES:
    datas += copy_metadata(package_name)

hiddenimports = []
for package_name in HIDDEN_IMPORT_PACKAGES:
    hiddenimports += collect_submodules(package_name)

binaries = []
for package_name in DYNAMIC_LIB_PACKAGES:
    binaries += collect_dynamic_libs(package_name)

a = Analysis(
    [str(project_root / ENTRY_SCRIPT)],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)
