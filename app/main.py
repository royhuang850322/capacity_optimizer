"""
Chemical Capacity Optimizer CLI entry point.

Excel-first workflow:
  python -m app.main --input-template "Tooling Control Panel/Capacity_Optimizer_Control.xlsx"
"""
from __future__ import annotations

import os
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

import click

from app.capacity_basis import CAPACITY_BASES, MAX_BASIS, PLANNED_BASIS
from app.create_template import refresh_control_workbook_license_sheet
from app.data_loader import (
    load_config,
    load_direct_mode_a_with_capacity_bases,
    load_direct_mode_b_with_capacity_bases,
)
from app.load_pressure import build_dashboard_fact_frame
from app.optimizer import run_optimization_mode_a, run_optimization_mode_b
from app.output_writer import write_capacity_basis_results, write_mode_comparison_summary, write_results
from app.run_logging import format_user_error, get_app_logger, setup_run_file_logging
from app.runtime_paths import ensure_workspace_dirs, resolve_runtime_paths
from app.validator import has_errors, print_issues, validate
from app.version import APP_VERSION
from app.workspace_init import initialize_user_workspace
from app.models import Config, ValidationIssue


_ACTIVE_LOG_PATH: Path | None = None


@click.command()
@click.option(
    "--input-template",
    "-t",
    required=True,
    help="Path to the Excel control workbook.",
)
@click.option(
    "--mode",
    type=click.Choice(["mode-a", "mode-b", "both"], case_sensitive=False),
    default=None,
    help="Override Run_Mode from the control workbook.",
)
@click.option(
    "--verbosity",
    type=click.Choice(["config", "verbose", "quiet"], case_sensitive=False),
    default="config",
    show_default=True,
    help="Override workbook verbosity, or keep the workbook setting.",
)
@click.option(
    "--validation-policy",
    type=click.Choice(["config", "skip-errors", "stop-on-errors"], case_sensitive=False),
    default="config",
    show_default=True,
    help="Override workbook validation behavior, or keep the workbook setting.",
)
@click.option(
    "--output-name",
    default=None,
    help="Override Output_FileName from the control workbook.",
)
def main(
    input_template: str,
    mode: str | None,
    verbosity: str,
    validation_policy: str,
    output_name: str | None,
) -> None:
    requested_template = os.path.abspath(input_template)
    runtime_paths = ensure_workspace_dirs(resolve_runtime_paths())
    default_workspace_template = os.path.abspath(str(runtime_paths.control_workbook_path))
    if runtime_paths.is_frozen and requested_template == default_workspace_template and not os.path.exists(requested_template):
        runtime_paths = initialize_user_workspace(runtime_paths).paths

    if not os.path.exists(input_template):
        _fatal(
            summary="Control workbook not found.",
            code="OPT-1001",
            details=f"Expected file path: {input_template}",
            hints=[
                "Use the desktop launcher to initialize workspace files.",
                "Confirm the workbook path in your run command or shortcut.",
            ],
        )

    try:
        config = load_config(input_template)
    except Exception as exc:
        _fatal(
            summary="Could not read control workbook.",
            code="OPT-1002",
            details=str(exc),
            hints=[
                "Close the workbook in Excel and retry.",
                "Check whether the file is a valid .xlsx workbook.",
            ],
            debug_details=traceback.format_exc(),
        )

    if output_name:
        config.output_file_name = output_name

    if mode:
        config.run_mode = _normalize_cli_mode(mode)
    if verbosity != "config":
        config.verbose = verbosity == "verbose"
    if validation_policy != "config":
        config.skip_validation_errors = validation_policy == "skip-errors"

    run_with_config(config, runtime_paths=runtime_paths, input_template=input_template)


def run_with_config(
    config: Config,
    *,
    runtime_paths=None,
    input_template: str | None = None,
) -> None:
    global _ACTIVE_LOG_PATH
    _ACTIVE_LOG_PATH = None
    _banner()
    runtime_paths = ensure_workspace_dirs(runtime_paths or resolve_runtime_paths())
    logger = get_app_logger()

    try:
        log_context = setup_run_file_logging(runtime_paths, run_label="optimizer_run")
        _set_active_log_path(log_context.log_path)
    except Exception as exc:
        click.echo(
            f"  Warning: could not initialize file logging ({exc}). Continuing without structured log.",
            err=True,
        )

    logger.info("Optimizer run started.")
    logger.debug(
        "Run source=%s | mode=%s output=%s",
        "workbook" if input_template else "launcher",
        config.run_mode,
        config.output_folder,
    )

    if input_template:
        click.echo(f"\n[1/4] Reading control workbook: {input_template}")
        logger.debug("Reading control workbook: %s", input_template)
    else:
        click.echo("\n[1/4] Reading launcher settings")
        logger.debug("Running without control workbook (launcher settings mode).")

    try:
        from app.license_validator import LicenseValidationError, validate_license_with_fallback
    except ModuleNotFoundError as exc:
        _fatal(
            summary="Required Python packages are missing or incomplete.",
            code="OPT-1003",
            details=str(exc),
            hints=[
                "Run runtime\\setup_requirements.bat, then retry.",
                "If packaged mode is used, re-run the installer package.",
            ],
            debug_details=traceback.format_exc(),
        )

    if runtime_paths.is_frozen:
        license_primary_root = str(runtime_paths.user_workspace_dir)
        license_fallback_roots = [
            config.project_root_folder,
            str(runtime_paths.app_install_dir),
        ]
    else:
        license_primary_root = config.project_root_folder
        license_fallback_roots = [str(runtime_paths.user_workspace_dir)]

    try:
        license_info = validate_license_with_fallback(
            primary_root=license_primary_root,
            fallback_roots=license_fallback_roots,
        )
    except LicenseValidationError as exc:
        _fatal(
            summary="License validation failed.",
            code="OPT-1201",
            details=str(exc),
            hints=[
                "Confirm licenses\\active\\license.json is present.",
                "If machine-locked, regenerate machine fingerprint and request a new license.",
            ],
        )
    except Exception as exc:
        _fatal(
            summary="License validation failed unexpectedly.",
            code="OPT-1202",
            details=str(exc),
            hints=[
                "Check whether the license file is complete and readable.",
                "Share the log file with support.",
            ],
            debug_details=traceback.format_exc(),
        )

    if license_info.project_root:
        config.project_root_folder = license_info.project_root

    config.license_status = license_info.status
    config.license_id = license_info.license_id
    config.license_type = license_info.license_type
    config.licensed_to = license_info.customer_name
    config.license_expiry = license_info.expiry_date
    config.license_binding_mode = license_info.binding_mode
    config.license_machine_label = license_info.machine_label

    if input_template:
        try:
            refresh_control_workbook_license_sheet(
                input_template,
                project_root=config.project_root_folder,
                license_info=license_info,
            )
        except PermissionError:
            click.echo(
                "  Note: could not refresh the License sheet because the control workbook is open in Excel.",
                err=False,
            )
            logger.warning("Could not refresh license sheet because workbook is open in Excel.")
        except Exception as exc:
            click.echo(
                f"  Note: could not refresh the License sheet: {exc}",
                err=False,
            )
            logger.warning("Could not refresh license sheet: %s", exc)

    config.run_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    selected_scenario = _selected_scenario(config.scenario_name)
    modes_to_run = _resolve_modes(config.run_mode)
    months = _build_month_list(config.start_month, config.horizon_months)

    click.echo(f"  Run mode     : {config.run_mode}")
    click.echo(f"  Scenario     : {config.scenario_name}")
    click.echo(f"  Start month  : {config.start_month}")
    click.echo(f"  Horizon      : {config.horizon_months} months")
    click.echo(f"  Input loads  : {config.input_load_folder}")
    click.echo(f"  Input master : {config.input_master_folder}")
    click.echo(f"  Output       : {config.output_folder}")
    click.echo(f"  License      : {config.license_status}")
    click.echo(f"  Licensed to  : {config.licensed_to}")
    click.echo(f"  Expires      : {config.license_expiry}")
    click.echo(f"  Binding      : {config.license_binding_mode}")
    click.echo(f"  Workspace    : {runtime_paths.user_workspace_dir}")
    if _ACTIVE_LOG_PATH is not None:
        click.echo(f"  Log file     : {_ACTIVE_LOG_PATH}")
    logger.debug(
        "Resolved config | run_mode=%s scenario=%s",
        config.run_mode,
        config.scenario_name,
    )

    try:
        _validate_input_setup(config, modes_to_run, runtime_paths.is_frozen)
    except Exception as exc:
        _fatal(
            summary="Input folder validation failed.",
            code="OPT-1301",
            details=str(exc),
            hints=[
                "Check Project Root / Input / Output folders in launcher settings.",
                "Confirm all required planner and master files exist.",
            ],
            debug_details=traceback.format_exc(),
        )

    click.echo("\n[2/4] Loading, validating, and optimizing")
    run_payloads: dict[str, dict[str, Any]] = {}
    metrics_by_mode: dict[str, dict[str, Any]] = {}

    for selected_mode in modes_to_run:
        click.echo(f"\n  --- {selected_mode} ---")
        logger.info("Mode started: %s", selected_mode)
        try:
            if selected_mode == "ModeA":
                loads, capacities_by_basis, routings = load_direct_mode_a_with_capacity_bases(
                    load_folder=config.input_load_folder,
                    master_folder=config.input_master_folder,
                    selected_scenario=selected_scenario,
                )
            else:
                loads, baseline_capacities_by_basis, capacities_by_basis, routings = load_direct_mode_b_with_capacity_bases(
                    load_folder=config.input_load_folder,
                    master_folder=config.input_master_folder,
                    selected_scenario=selected_scenario,
                )
        except Exception as exc:
            _fatal(
                summary=f"{selected_mode} failed during data loading.",
                code="OPT-1401",
                details=str(exc),
                hints=[
                    "Check CSV/Excel column names and source file formats.",
                    "Confirm scenario names and month values are valid.",
                ],
                debug_details=traceback.format_exc(),
            )

        click.echo(f"    Load records  : {len(loads):,}")
        if selected_mode == "ModeA":
            click.echo(
                "    Capacity rows : "
                f"Max={len(capacities_by_basis[MAX_BASIS]):,} | Planned={len(capacities_by_basis[PLANNED_BASIS]):,}"
            )
        else:
            click.echo(
                "    Capacity rows : "
                f"Baseline Max={len(baseline_capacities_by_basis[MAX_BASIS]):,} | "
                f"Baseline Planned={len(baseline_capacities_by_basis[PLANNED_BASIS]):,} | "
                f"ModeB Max={len(capacities_by_basis[MAX_BASIS]):,} | "
                f"ModeB Planned={len(capacities_by_basis[PLANNED_BASIS]):,}"
            )
        click.echo(f"    Routing rows  : {len(routings):,}")
        basis_payloads: dict[str, dict[str, Any]] = {}
        combined_issues: list[ValidationIssue] = []

        for capacity_basis in CAPACITY_BASES:
            click.echo(f"    [{capacity_basis}]")
            capacities = capacities_by_basis[capacity_basis]
            baseline_capacities = (
                capacities
                if selected_mode == "ModeA"
                else baseline_capacities_by_basis[capacity_basis]
            )
            issues = validate(
                loads,
                baseline_capacities,
                routings,
                mode=selected_mode,
                routing_capacities=capacities,
            )
            _basis_print_issues(capacity_basis, issues)
            combined_issues.extend(_prefix_issues(capacity_basis, issues))

            if has_errors(issues) and not config.skip_validation_errors:
                _fatal(
                    summary="Validation found ERROR-level issues.",
                    code="OPT-1501",
                    details=f"{selected_mode} / {capacity_basis} stopped because validation failed.",
                    hints=[
                        "Fix source data issues reported above and rerun.",
                        "If you need a forced run, set Skip Validation Errors = Yes.",
                    ],
                )

            if selected_mode == "ModeA":
                results = run_optimization_mode_a(
                    months=months,
                    loads=loads,
                    capacities=capacities,
                    verbose=config.verbose,
                )
                toller_products = set()
            else:
                results, toller_products = run_optimization_mode_b(
                    months=months,
                    loads=loads,
                    baseline_capacities=baseline_capacities,
                    routing_capacities=capacities,
                    routings=routings,
                    verbose=config.verbose,
                )

            metrics = _build_run_metrics(
                selected_mode=selected_mode,
                capacity_basis=capacity_basis,
                loads=loads,
                results=results,
                months=months,
                scenario_name=config.scenario_name,
            )
            basis_payloads[capacity_basis] = {
                "capacities": capacities,
                "results": results,
                "issues": issues,
                "toller_products": toller_products,
                "metrics": metrics,
            }
            _print_run_metrics(metrics)

        metrics_by_mode[selected_mode] = basis_payloads[PLANNED_BASIS]["metrics"]
        logger.debug("Mode metrics [%s]: %s", selected_mode, metrics_by_mode[selected_mode])
        run_payloads[selected_mode] = {
            "loads": loads,
            "baseline_capacities_by_basis": {
                basis: capacities_by_basis[basis] if selected_mode == "ModeA" else baseline_capacities_by_basis[basis]
                for basis in CAPACITY_BASES
            },
            "capacities_by_basis": capacities_by_basis,
            "routings": routings,
            "basis_payloads": basis_payloads,
            "issues": combined_issues,
        }

    click.echo("\n[3/4] Writing Excel result workbooks")
    output_paths: dict[str, str] = {}
    for selected_mode in modes_to_run:
        payload = run_payloads[selected_mode]
        try:
            output_paths[selected_mode] = write_capacity_basis_results(
                basis_results={
                    basis: payload["basis_payloads"][basis]["results"]
                    for basis in CAPACITY_BASES
                },
                loads=payload["loads"],
                basis_capacities=payload["capacities_by_basis"],
                routings=payload["routings"],
                config=config,
                issues=payload["issues"],
                months=months,
                mode=selected_mode,
                toller_products_by_basis={
                    basis: payload["basis_payloads"][basis]["toller_products"]
                    for basis in CAPACITY_BASES
                },
                unmet_capacities_by_basis=payload["baseline_capacities_by_basis"],
            )
        except Exception as exc:
            _fatal(
                summary=f"Failed to write {selected_mode} output workbook.",
                code="OPT-1601",
                details=str(exc),
                hints=[
                    "Close any output workbook that is open in Excel.",
                    "Check output folder permissions and available disk space.",
                ],
                debug_details=traceback.format_exc(),
            )
        click.echo(f"  {selected_mode}: {output_paths[selected_mode]}")
        logger.info("Workbook written for %s: %s", selected_mode, output_paths[selected_mode])

    if set(modes_to_run) == {"ModeA", "ModeB"}:
        try:
            comparison_path = write_mode_comparison_summary(
                mode_results={
                    mode_name: _mode_results_for_summary(run_payloads[mode_name])
                    for mode_name in ("ModeA", "ModeB")
                },
                mode_loads={mode_name: run_payloads[mode_name]["loads"] for mode_name in ("ModeA", "ModeB")},
                mode_capacities={
                    mode_name: _mode_capacities_for_summary(run_payloads[mode_name])
                    for mode_name in ("ModeA", "ModeB")
                },
                mode_routings={mode_name: run_payloads[mode_name]["routings"] for mode_name in ("ModeA", "ModeB")},
                config=config,
                months=months,
                metrics_by_mode=metrics_by_mode,
                dashboard_facts_by_mode={
                    mode_name: build_dashboard_fact_frame(
                    mode=mode_name,
                    results=_mode_results_for_summary(run_payloads[mode_name]),
                    loads=run_payloads[mode_name]["loads"],
                    capacities=_mode_capacities_for_summary(run_payloads[mode_name]),
                    routings=run_payloads[mode_name]["routings"],
                    unmet_capacities=_mode_unmet_capacities_for_summary(run_payloads[mode_name]),
                )
                for mode_name in ("ModeA", "ModeB")
            },
            capacity_basis_payloads_by_mode={
                mode_name: _mode_capacity_basis_payload(run_payloads[mode_name])
                for mode_name in ("ModeA", "ModeB")
            },
            mode_unmet_capacities={
                mode_name: _mode_unmet_capacities_for_summary(run_payloads[mode_name])
                for mode_name in ("ModeA", "ModeB")
            },
        )
        except Exception as exc:
            _fatal(
                summary="Failed to write ModeA/ModeB summary workbook.",
                code="OPT-1602",
                details=str(exc),
                hints=[
                    "Close any summary workbook that is open in Excel.",
                    "Check output folder permissions and available disk space.",
                ],
                debug_details=traceback.format_exc(),
            )
        click.echo(f"  Summary : {comparison_path}")
        logger.info("Comparison summary workbook written: %s", comparison_path)

    click.echo("\n[4/4] Completed")
    click.echo("  Result workbooks contain dashboard and analysis sheets in Excel.")
    click.echo("")
    logger.info("Optimizer run completed successfully.")


def _build_run_metrics(
    *,
    selected_mode: str,
    capacity_basis: str,
    loads,
    results,
    months: list[str],
    scenario_name: str,
) -> dict[str, Any]:
    total_demand = _total_demand(loads, months)
    total_internal_allocated = _total_internal_allocated(results)
    total_outsourced = _total_outsourced(results)
    total_unmet = _total_unmet(results)
    total_supplied = total_internal_allocated + total_outsourced
    service_level = 100.0 * total_supplied / total_demand if total_demand > 0 else 0.0
    return {
        "mode": selected_mode,
        "capacity_basis": capacity_basis,
        "scenario_name": scenario_name,
        "total_demand": total_demand,
        "total_internal_allocated": total_internal_allocated,
        "total_outsourced": total_outsourced,
        "total_unmet": total_unmet,
        "service_level": service_level,
        "result_rows": len(results),
        "months": len(months),
    }


def _print_run_metrics(metrics: dict[str, Any]) -> None:
    click.echo(f"      Total demand    : {metrics['total_demand']:>12,.1f} tons")
    click.echo(f"      Internal alloc. : {metrics['total_internal_allocated']:>12,.1f} tons")
    click.echo(f"      Outsourced      : {metrics['total_outsourced']:>12,.1f} tons")
    click.echo(f"      Total unmet     : {metrics['total_unmet']:>12,.1f} tons")
    click.echo(f"      Service level   : {metrics['service_level']:>11.1f}%")


def _mode_results_for_summary(payload: dict[str, Any]):
    if "basis_payloads" in payload:
        return payload["basis_payloads"][PLANNED_BASIS]["results"]
    return payload["results"]


def _mode_capacities_for_summary(payload: dict[str, Any]):
    if "capacities_by_basis" in payload:
        return payload["capacities_by_basis"][PLANNED_BASIS]
    return payload["capacities"]


def _mode_unmet_capacities_for_summary(payload: dict[str, Any]):
    if "baseline_capacities_by_basis" in payload:
        return payload["baseline_capacities_by_basis"][PLANNED_BASIS]
    return payload["capacities"]


def _mode_capacity_basis_payload(payload: dict[str, Any]) -> dict[str, Any]:
    if "basis_payloads" in payload:
        return {
            "basis_results": {
                basis: payload["basis_payloads"][basis]["results"]
                for basis in CAPACITY_BASES
            },
            "basis_capacities": payload["capacities_by_basis"],
            "unmet_capacities_by_basis": payload["baseline_capacities_by_basis"],
            "loads": payload["loads"],
            "routings": payload["routings"],
        }

    return {
        "basis_results": {
            "Max": payload["results"],
            PLANNED_BASIS: payload["results"],
        },
        "basis_capacities": {
            "Max": payload["capacities"],
            PLANNED_BASIS: payload["capacities"],
        },
        "unmet_capacities_by_basis": {
            "Max": payload["capacities"],
            PLANNED_BASIS: payload["capacities"],
        },
        "loads": payload["loads"],
        "routings": payload["routings"],
    }

def _resolve_modes(run_mode: str) -> list[str]:
    normalized = str(run_mode or "").strip().lower()
    if normalized == "modea":
        return ["ModeA"]
    if normalized == "both":
        return ["ModeA", "ModeB"]
    return ["ModeB"]


def _normalize_cli_mode(value: str) -> str:
    text = value.strip().lower()
    if text == "mode-a":
        return "ModeA"
    if text == "both":
        return "Both"
    return "ModeB"


def _validate_input_setup(config, modes_to_run: list[str], is_frozen: bool = False) -> None:
    _validate_required_directory(
        "Project_Root_Folder",
        config.project_root_folder,
        required_entries=None if is_frozen else [],
        purpose="workspace root",
    )
    _validate_required_directory(
        "Input_Load_Folder",
        config.input_load_folder,
        purpose="planner input",
    )
    _validate_required_directory(
        "Input_Master_Folder",
        config.input_master_folder,
        purpose="master data",
    )
    _validate_output_folder(config.output_folder)
    _validate_planner_files(config.input_load_folder)
    _validate_capacity_master(config.input_master_folder)
    if "ModeB" in modes_to_run:
        _validate_routing_file(config.input_master_folder, ["alternative_routing", "master_routing"])


def _validate_required_directory(
    label: str,
    path: str,
    required_entries: list[str] | None = None,
    purpose: str = "directory",
) -> None:
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f"{label} does not exist: {path}")
    if not os.path.isdir(path):
        raise NotADirectoryError(f"{label} is not a folder: {path}")
    missing = [
        entry for entry in (required_entries or [])
        if not os.path.exists(os.path.join(path, entry))
    ]
    if missing:
        raise FileNotFoundError(
            f"{label} is not a valid {purpose}: {path}. Missing expected item(s): {', '.join(missing)}"
        )


def _validate_output_folder(path: str) -> None:
    if not path:
        raise FileNotFoundError("Output_Folder is empty.")
    normalized = os.path.abspath(path)
    parent = normalized if os.path.isdir(normalized) else os.path.dirname(normalized) or normalized
    if not os.path.exists(parent):
        raise FileNotFoundError(
            f"Output_Folder parent does not exist: {parent}"
        )
    if not os.path.isdir(parent):
        raise NotADirectoryError(
            f"Output_Folder parent is not a folder: {parent}"
        )


def _validate_planner_files(load_folder: str) -> None:
    from app.data_loader import _find_planner_files

    _find_planner_files(load_folder)


def _validate_master_file(master_folder: str, stem: str) -> None:
    for ext in (".xlsx", ".xls", ".csv"):
        if os.path.exists(os.path.join(master_folder, stem + ext)):
            return
    raise FileNotFoundError(
        f"Required master file '{stem}' not found in {master_folder}. "
        f"Expected {stem}.xlsx, {stem}.xls, or {stem}.csv."
    )


def _validate_capacity_master(master_folder: str) -> None:
    _validate_master_file(master_folder, "master_capacity")


def _validate_routing_file(master_folder: str, stems: list[str]) -> None:
    for stem in stems:
        for ext in (".xlsx", ".xls", ".csv"):
            if os.path.exists(os.path.join(master_folder, stem + ext)):
                return
    raise FileNotFoundError(
        f"Required routing file not found in {master_folder}. "
        f"Tried: {', '.join(stems)} with .xlsx/.xls/.csv."
    )


def _prefix_issues(capacity_basis: str, issues: list[ValidationIssue]) -> list[ValidationIssue]:
    return [
        ValidationIssue(
            severity=issue.severity,
            check=f"{capacity_basis}:{issue.check}",
            detail=f"[{capacity_basis}] {issue.detail}",
        )
        for issue in issues
    ]


def _basis_print_issues(capacity_basis: str, issues: list[ValidationIssue]) -> None:
    if not issues:
        click.echo(f"      Validation    : no issues for {capacity_basis}")
        return
    click.echo(f"      Validation    : {capacity_basis}")
    print_issues(issues)


def _build_month_list(start: str, count: int) -> list[str]:
    year, month = int(start[:4]), int(start[5:])
    months: list[str] = []
    for _ in range(count):
        months.append(f"{year}-{month:02d}")
        month += 1
        if month > 12:
            month = 1
            year += 1
    return months


def _total_demand(loads, months: list[str] | None = None) -> float:
    month_filter = set(months) if months is not None else None
    demand_by_month_product: dict[tuple[str, str], float] = {}
    for load in loads:
        if month_filter is not None and load.month not in month_filter:
            continue
        if load.forecast_tons <= 0:
            continue
        key = (load.month, load.product)
        demand_by_month_product[key] = demand_by_month_product.get(key, 0.0) + load.forecast_tons
    return sum(demand_by_month_product.values())


def _total_unmet(results) -> float:
    unmet_by_month_product: dict[tuple[str, str], float] = {}
    for result in results:
        key = (result.month, result.product)
        unmet_by_month_product[key] = max(
            unmet_by_month_product.get(key, 0.0),
            result.unmet_tons,
        )
    return sum(unmet_by_month_product.values())


def _total_internal_allocated(results) -> float:
    return sum(result.allocated_tons for result in results if result.allocation_type == "Internal")


def _total_outsourced(results) -> float:
    return sum(result.outsourced_tons for result in results if result.allocation_type == "Outsourced")


def _selected_scenario(configured_scenario: str | None) -> str | None:
    scenario = str(configured_scenario or "").strip()
    if not scenario or scenario.lower() in {"base", "base scenario"}:
        return None
    return scenario


def _banner() -> None:
    click.echo("=" * 60)
    click.echo(f"  Chemical Capacity Optimizer  {APP_VERSION}")
    click.echo("  Launcher Settings / Excel Workbook + Python Optimization + Excel Reports")
    click.echo("=" * 60)


def _set_active_log_path(log_path: Path) -> None:
    global _ACTIVE_LOG_PATH
    _ACTIVE_LOG_PATH = log_path


def _fatal(
    *,
    summary: str,
    code: str,
    details: str | None = None,
    hints: list[str] | None = None,
    debug_details: str | None = None,
) -> None:
    logger = get_app_logger()
    logger.error("[%s] %s | details=%s", code, summary, details or "")
    if debug_details:
        logger.debug("Traceback for [%s]:\n%s", code, debug_details)

    user_error = format_user_error(
        code=code,
        summary=summary,
        details=details,
        hints=hints,
        log_path=_ACTIVE_LOG_PATH,
    )
    click.echo(f"\n{user_error}", err=True)
    sys.exit(1)


if __name__ == "__main__":
    main()
