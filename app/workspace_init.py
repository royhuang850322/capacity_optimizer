"""
Workspace initialization helpers for the desktop launcher and packaged runs.

These helpers keep first-run behavior predictable without changing the core
optimization flow.
"""
from __future__ import annotations

import json
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.create_template import write_control_workbook
from app.runtime_paths import RuntimePaths, ensure_workspace_dirs, resolve_runtime_paths


CUSTOMER_DOC_FILES = (
    "CUSTOMER_LICENSE_QUICKSTART_CN.md",
    "IT_DEPLOYMENT_CHECKLIST_CN.md",
    "desktop_launcher_usage.md",
    "PYTHON_INSTALL_GUIDE_CN.md",
)


@dataclass(frozen=True)
class WorkspaceInitializationResult:
    paths: RuntimePaths
    workbook_created: bool
    sample_data_copied: bool


def initialize_user_workspace(paths: RuntimePaths | None = None) -> WorkspaceInitializationResult:
    runtime_paths = ensure_workspace_dirs(paths or resolve_runtime_paths())
    sample_data_copied = _ensure_workspace_sample_data(runtime_paths)
    _ensure_workspace_docs(runtime_paths)
    workbook_created = _ensure_control_workbook(runtime_paths)
    _write_workspace_manifest(runtime_paths)
    return WorkspaceInitializationResult(
        paths=runtime_paths,
        workbook_created=workbook_created,
        sample_data_copied=sample_data_copied,
    )


def _ensure_workspace_sample_data(paths: RuntimePaths) -> bool:
    if paths.workspace_input_dir.resolve() == paths.sample_data_dir.resolve():
        return False
    if paths.workspace_input_dir.exists() and any(paths.workspace_input_dir.iterdir()):
        return False
    if not paths.sample_data_dir.exists():
        return False
    shutil.copytree(paths.sample_data_dir, paths.workspace_input_dir, dirs_exist_ok=True)
    return True


def _ensure_control_workbook(paths: RuntimePaths) -> bool:
    if paths.control_workbook_path.exists():
        return False
    load_dir = paths.workspace_input_dir if paths.workspace_input_dir.exists() else paths.sample_data_dir
    write_control_workbook(str(paths.control_workbook_path), load_dir=str(load_dir))
    return True


def _ensure_workspace_docs(paths: RuntimePaths) -> bool:
    if not paths.bundled_docs_dir.exists():
        return False

    copied_any = False
    for filename in CUSTOMER_DOC_FILES:
        source = paths.bundled_docs_dir / filename
        target = paths.workspace_docs_dir / filename
        if not source.exists() or target.exists():
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        copied_any = True
    return copied_any


def _write_workspace_manifest(paths: RuntimePaths) -> None:
    existing = _read_existing_manifest(paths.workspace_manifest_path)
    payload = {
        "app_name": "CapacityOptimizer",
        "initialized_at": existing.get("initialized_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "last_checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "install_dir": str(paths.app_install_dir),
        "workspace_dir": str(paths.user_workspace_dir),
        "control_workbook_path": str(paths.control_workbook_path),
        "outputs_dir": str(paths.outputs_dir),
        "logs_dir": str(paths.logs_dir),
        "license_dir": str(paths.license_dir),
        "docs_dir": str(paths.workspace_docs_dir),
        "sample_data_dir": str(paths.workspace_input_dir),
        "is_frozen": paths.is_frozen,
    }
    with open(paths.workspace_manifest_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def _read_existing_manifest(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}
