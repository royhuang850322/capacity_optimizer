# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys

from PyInstaller.utils.hooks import collect_dynamic_libs, collect_submodules, copy_metadata

project_root = Path(SPECPATH).resolve().parents[0]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from build_support.packaging_manifest import get_target, iter_data_mappings

target = get_target("modeb_product_analysis")
datas = iter_data_mappings(project_root, target_id=target.target_id)
for package_name in target.metadata_packages:
    datas += copy_metadata(package_name)

hiddenimports = []
for package_name in target.hidden_import_packages:
    hiddenimports += collect_submodules(package_name)

binaries = []
for package_name in target.dynamic_lib_packages:
    binaries += collect_dynamic_libs(package_name)

a = Analysis(
    [str(project_root / target.entry_script)],
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
    name=target.app_name,
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
    name=target.app_name,
)
