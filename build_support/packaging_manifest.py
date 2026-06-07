"""
Shared packaging manifest for PyInstaller one-folder builds.

This module stays pure-Python so tests can validate the packaging plan without
requiring PyInstaller to be installed.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PackagingTarget:
    target_id: str
    app_name: str
    entry_script: str
    resource_dir_mappings: tuple[tuple[str, str], ...]
    resource_file_mappings: tuple[tuple[str, str], ...]
    required_resource_subpaths: tuple[str, ...]
    hidden_import_packages: tuple[str, ...]
    dynamic_lib_packages: tuple[str, ...]
    metadata_packages: tuple[str, ...]


_COMMON_RESOURCE_DIR_MAPPINGS: tuple[tuple[str, str], ...] = ()

_COMMON_RESOURCE_FILE_MAPPINGS = (
    ("Data_Input/planner3_load.csv", "resources/Data_Input"),
    ("Data_Input/master_capacity.csv", "resources/Data_Input"),
    ("Data_Input/master_routing.csv", "resources/Data_Input"),
    ("Data_Input/DATA_INPUT_GUIDE_CN.md", "resources/Data_Input"),
    ("README.md", "resources/docs"),
    ("LICENSE", "resources/docs"),
    ("docs/CUSTOMER_LICENSE_QUICKSTART_CN.md", "resources/docs"),
    ("docs/IT_DEPLOYMENT_CHECKLIST_CN.md", "resources/docs"),
    ("docs/desktop_launcher_usage.md", "resources/docs"),
    ("docs/runtime_directory_strategy.md", "resources/docs"),
    ("docs/PYTHON_INSTALL_GUIDE_CN.md", "resources/docs"),
)

_COMMON_HIDDEN_IMPORT_PACKAGES = (
    "ortools",
)

_COMMON_DYNAMIC_LIB_PACKAGES = (
    "ortools",
)

_COMMON_METADATA_PACKAGES = (
    "ortools",
    "pandas",
    "openpyxl",
    "click",
    "colorama",
    "cryptography",
)

DEFAULT_TARGET_ID = "capacity_optimizer"
LEGACY_PRODUCT_ANALYSIS_TARGET_ID = "modeb_product_analysis"
PRODUCT_ANALYSIS_TARGET_ID = "product_analysis"

TARGETS: dict[str, PackagingTarget] = {
    "capacity_optimizer": PackagingTarget(
        target_id="capacity_optimizer",
        app_name="CapacityOptimizer",
        entry_script="CapacityOptimizerLauncher.pyw",
        resource_dir_mappings=_COMMON_RESOURCE_DIR_MAPPINGS,
        resource_file_mappings=_COMMON_RESOURCE_FILE_MAPPINGS,
        required_resource_subpaths=("Data_Input", "docs"),
        hidden_import_packages=_COMMON_HIDDEN_IMPORT_PACKAGES,
        dynamic_lib_packages=_COMMON_DYNAMIC_LIB_PACKAGES,
        metadata_packages=_COMMON_METADATA_PACKAGES,
    ),
    PRODUCT_ANALYSIS_TARGET_ID: PackagingTarget(
        target_id=PRODUCT_ANALYSIS_TARGET_ID,
        app_name="ProductAnalysis",
        entry_script="ProductAnalysisLauncher.pyw",
        resource_dir_mappings=(),
        resource_file_mappings=(),
        required_resource_subpaths=(),
        hidden_import_packages=_COMMON_HIDDEN_IMPORT_PACKAGES,
        dynamic_lib_packages=_COMMON_DYNAMIC_LIB_PACKAGES,
        metadata_packages=_COMMON_METADATA_PACKAGES,
    ),
    "workcenter_analysis": PackagingTarget(
        target_id="workcenter_analysis",
        app_name="WorkCenterAnalysis",
        entry_script="WorkCenterAnalysisLauncher.pyw",
        resource_dir_mappings=(),
        resource_file_mappings=(),
        required_resource_subpaths=(),
        hidden_import_packages=_COMMON_HIDDEN_IMPORT_PACKAGES,
        dynamic_lib_packages=_COMMON_DYNAMIC_LIB_PACKAGES,
        metadata_packages=_COMMON_METADATA_PACKAGES,
    ),
}


def get_target(target_id: str = DEFAULT_TARGET_ID) -> PackagingTarget:
    if target_id == LEGACY_PRODUCT_ANALYSIS_TARGET_ID:
        target_id = PRODUCT_ANALYSIS_TARGET_ID
    try:
        return TARGETS[target_id]
    except KeyError as exc:
        valid = ", ".join(sorted(TARGETS))
        raise KeyError(f"Unknown packaging target '{target_id}'. Valid targets: {valid}") from exc


def iter_data_mappings(project_root: Path, *, target_id: str = DEFAULT_TARGET_ID) -> list[tuple[str, str]]:
    target = get_target(target_id)
    mappings: list[tuple[str, str]] = []
    for relative_source, target_dir in target.resource_dir_mappings:
        source_path = project_root / relative_source
        if not source_path.exists():
            raise FileNotFoundError(f"Required packaging resource directory not found: {source_path}")
        mappings.append((str(source_path), target_dir))
    for relative_source, target_dir in target.resource_file_mappings:
        source_path = project_root / relative_source
        if not source_path.exists():
            raise FileNotFoundError(f"Required packaging resource file not found: {source_path}")
        mappings.append((str(source_path), target_dir))
    return mappings


# Backward-compatible aliases for the default/main application target.
_DEFAULT_TARGET = get_target(DEFAULT_TARGET_ID)
APP_NAME = _DEFAULT_TARGET.app_name
ENTRY_SCRIPT = _DEFAULT_TARGET.entry_script
RESOURCE_DIR_MAPPINGS = _DEFAULT_TARGET.resource_dir_mappings
RESOURCE_FILE_MAPPINGS = _DEFAULT_TARGET.resource_file_mappings
REQUIRED_RESOURCE_SUBPATHS = _DEFAULT_TARGET.required_resource_subpaths
HIDDEN_IMPORT_PACKAGES = _DEFAULT_TARGET.hidden_import_packages
DYNAMIC_LIB_PACKAGES = _DEFAULT_TARGET.dynamic_lib_packages
METADATA_PACKAGES = _DEFAULT_TARGET.metadata_packages
