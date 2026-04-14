"""
Centralized runtime path resolution for source mode and frozen/packaged mode.

This module is intentionally small and low-magic:
- source mode keeps the current repository-root-centric behavior
- frozen mode separates install resources from user-writable workspace
- callers can request one object and avoid scattering path rules
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path


APP_NAME = "CapacityOptimizer"
WORKSPACE_ENV_VAR = "CAPACITY_OPTIMIZER_WORKSPACE"


@dataclass(frozen=True)
class RuntimePaths:
    app_install_dir: Path
    bundled_resources_dir: Path
    bundled_docs_dir: Path
    user_workspace_dir: Path
    templates_dir: Path
    workspace_docs_dir: Path
    workspace_input_dir: Path
    outputs_dir: Path
    logs_dir: Path
    license_dir: Path
    license_active_dir: Path
    license_requests_dir: Path
    workspace_manifest_path: Path
    control_workbook_path: Path
    sample_data_dir: Path
    is_frozen: bool


def is_frozen_runtime() -> bool:
    return bool(getattr(sys, "frozen", False))


def resolve_runtime_paths() -> RuntimePaths:
    frozen = is_frozen_runtime()
    install_dir = _resolve_install_dir(frozen)
    bundled_resources_dir = _resolve_bundled_resources_dir(install_dir, frozen)
    user_workspace_dir = _resolve_user_workspace_dir(install_dir, frozen)

    bundled_docs_dir = bundled_resources_dir / "docs"
    templates_dir = user_workspace_dir / "Tooling Control Panel"
    workspace_docs_dir = user_workspace_dir / "docs"
    workspace_input_dir = user_workspace_dir / "Data_Input"
    outputs_dir = user_workspace_dir / "output"
    logs_dir = user_workspace_dir / "logs"
    license_dir = user_workspace_dir / "licenses"
    license_active_dir = license_dir / "active"
    license_requests_dir = license_dir / "requests"
    workspace_manifest_path = user_workspace_dir / "workspace_manifest.json"
    control_workbook_path = templates_dir / "Capacity_Optimizer_Control.xlsx"
    sample_data_dir = bundled_resources_dir / "Data_Input"

    return RuntimePaths(
        app_install_dir=install_dir,
        bundled_resources_dir=bundled_resources_dir,
        bundled_docs_dir=bundled_docs_dir,
        user_workspace_dir=user_workspace_dir,
        templates_dir=templates_dir,
        workspace_docs_dir=workspace_docs_dir,
        workspace_input_dir=workspace_input_dir,
        outputs_dir=outputs_dir,
        logs_dir=logs_dir,
        license_dir=license_dir,
        license_active_dir=license_active_dir,
        license_requests_dir=license_requests_dir,
        workspace_manifest_path=workspace_manifest_path,
        control_workbook_path=control_workbook_path,
        sample_data_dir=sample_data_dir,
        is_frozen=frozen,
    )


def ensure_workspace_dirs(paths: RuntimePaths | None = None) -> RuntimePaths:
    runtime_paths = paths or resolve_runtime_paths()
    for path in (
        runtime_paths.user_workspace_dir,
        runtime_paths.templates_dir,
        runtime_paths.workspace_docs_dir,
        runtime_paths.workspace_input_dir,
        runtime_paths.outputs_dir,
        runtime_paths.logs_dir,
        runtime_paths.license_dir,
        runtime_paths.license_active_dir,
        runtime_paths.license_requests_dir,
    ):
        path.mkdir(parents=True, exist_ok=True)
    return runtime_paths


def _resolve_install_dir(frozen: bool) -> Path:
    if frozen:
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _resolve_bundled_resources_dir(install_dir: Path, frozen: bool) -> Path:
    if frozen:
        resource_candidates = [
            install_dir / "resources",
            install_dir / "_internal" / "resources",
            install_dir,
        ]
        for candidate in resource_candidates:
            if candidate.exists():
                return candidate
        return resource_candidates[0]
    return install_dir


def _resolve_user_workspace_dir(install_dir: Path, frozen: bool) -> Path:
    override = os.environ.get(WORKSPACE_ENV_VAR, "").strip()
    if override:
        return Path(override).expanduser().resolve()

    if not frozen:
        return install_dir

    local_appdata = os.environ.get("LOCALAPPDATA", "").strip()
    if local_appdata:
        return Path(local_appdata) / APP_NAME

    return Path.home() / f".{APP_NAME}"
