"""
Shared packaging manifest for PyInstaller one-folder builds.

This module stays pure-Python so tests can validate the packaging plan without
requiring PyInstaller to be installed.
"""
from __future__ import annotations

from pathlib import Path


APP_NAME = "CapacityOptimizer"
ENTRY_SCRIPT = "CapacityOptimizerLauncher.pyw"

RESOURCE_DIR_MAPPINGS = (
    ("Data_Input", "resources/Data_Input"),
)

RESOURCE_FILE_MAPPINGS = (
    ("README.md", "resources/docs"),
    ("LICENSE", "resources/docs"),
    ("docs/CUSTOMER_LICENSE_QUICKSTART_CN.md", "resources/docs"),
    ("docs/IT_DEPLOYMENT_CHECKLIST_CN.md", "resources/docs"),
    ("docs/desktop_launcher_usage.md", "resources/docs"),
    ("docs/runtime_directory_strategy.md", "resources/docs"),
    ("docs/PYTHON_INSTALL_GUIDE_CN.md", "resources/docs"),
)

HIDDEN_IMPORT_PACKAGES = (
    "ortools",
)

DYNAMIC_LIB_PACKAGES = (
    "ortools",
)

METADATA_PACKAGES = (
    "ortools",
    "pandas",
    "openpyxl",
    "click",
    "colorama",
    "cryptography",
)


def iter_data_mappings(project_root: Path) -> list[tuple[str, str]]:
    mappings: list[tuple[str, str]] = []
    for relative_source, target_dir in RESOURCE_DIR_MAPPINGS:
        source_path = project_root / relative_source
        if not source_path.exists():
            raise FileNotFoundError(f"Required packaging resource directory not found: {source_path}")
        mappings.append((str(source_path), target_dir))
    for relative_source, target_dir in RESOURCE_FILE_MAPPINGS:
        source_path = project_root / relative_source
        if not source_path.exists():
            raise FileNotFoundError(f"Required packaging resource file not found: {source_path}")
        mappings.append((str(source_path), target_dir))
    return mappings
