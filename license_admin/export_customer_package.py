"""
Build a clean customer delivery package from the development repository.

This internal tool keeps developer-only assets out of the exported package
and regenerates a fresh control workbook for the delivery folder.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from app.create_template import write_control_workbook
from license_admin.license_tools.common import sanitize_path_component

DEFAULT_DELIVERY_ROOT = REPO_ROOT / "delivery_packages"
TOOL_NAME = "capacity_optimizer"
CUSTOMER_DOCS = [
    "docs/CUSTOMER_LICENSE_QUICKSTART_CN.md",
    "docs/IT_DEPLOYMENT_CHECKLIST_CN.md",
]
ROOT_FILES = [
    "requirements.txt",
    "LICENSE",
]
PACKAGE_DIRS = [
    "app",
    "runtime",
]


def build_customer_package(
    *,
    project_root: Path | str,
    destination_root: Path | str,
    customer_name: str,
    package_name: str | None = None,
    license_file: Path | str | None = None,
    include_demo_data: bool = True,
    overwrite: bool = False,
) -> Path:
    project_root = Path(project_root)
    destination_root = Path(destination_root)
    license_file = Path(license_file) if license_file is not None else None

    customer_slug = sanitize_path_component(customer_name)
    package_dir_name = package_name or f"{TOOL_NAME}_{customer_slug}"
    package_path = destination_root / package_dir_name

    if package_path.exists():
        if not overwrite:
            raise FileExistsError(f"Delivery package already exists: {package_path}")
        shutil.rmtree(package_path)

    package_path.mkdir(parents=True, exist_ok=True)

    for relative_path in ROOT_FILES:
        _copy_file(project_root / relative_path, package_path / relative_path)

    for relative_dir in PACKAGE_DIRS:
        _copy_tree(project_root / relative_dir, package_path / relative_dir)

    if include_demo_data:
        _copy_tree(project_root / "Data_Input", package_path / "Data_Input")
    else:
        (package_path / "Data_Input").mkdir(parents=True, exist_ok=True)

    docs_dir = package_path / "docs"
    docs_dir.mkdir(parents=True, exist_ok=True)
    for relative_path in CUSTOMER_DOCS:
        _copy_file(project_root / relative_path, package_path / relative_path)

    (package_path / "output").mkdir(parents=True, exist_ok=True)
    (package_path / "licenses" / "active").mkdir(parents=True, exist_ok=True)
    (package_path / "licenses" / "requests").mkdir(parents=True, exist_ok=True)
    (package_path / "Tooling Control Panel").mkdir(parents=True, exist_ok=True)

    workbook_path = package_path / "Tooling Control Panel" / "Capacity_Optimizer_Control.xlsx"
    write_control_workbook(str(workbook_path), load_dir=str(package_path / "Data_Input"))

    if license_file is not None:
        if not license_file.exists():
            raise FileNotFoundError(f"License file not found: {license_file}")
        shutil.copy2(license_file, package_path / "licenses" / "active" / "license.json")

    _write_delivery_readme(
        destination=package_path,
        customer_name=customer_name,
        license_included=license_file is not None,
        include_demo_data=include_demo_data,
    )
    _write_manifest(
        destination=package_path,
        customer_name=customer_name,
        package_name=package_dir_name,
        license_included=license_file is not None,
        include_demo_data=include_demo_data,
    )

    return package_path


def _copy_file(source: Path, destination: Path) -> None:
    if not source.exists():
        raise FileNotFoundError(f"Required source file not found: {source}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def _copy_tree(source: Path, destination: Path) -> None:
    if not source.exists():
        raise FileNotFoundError(f"Required source directory not found: {source}")
    shutil.copytree(
        source,
        destination,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo"),
        dirs_exist_ok=True,
    )


def _write_delivery_readme(
    *,
    destination: Path,
    customer_name: str,
    license_included: bool,
    include_demo_data: bool,
) -> None:
    license_text = (
        "A signed license file has already been included in `licenses\\active\\license.json`."
        if license_included
        else "No license file is bundled. Put the signed license into `licenses\\active\\license.json` before running."
    )
    data_text = (
        "Demo input data is included under `Data_Input\\`."
        if include_demo_data
        else "No demo input data was included; place customer files under `Data_Input\\` or point the control workbook to another folder."
    )

    content = f"""# Chemical Capacity Optimizer Delivery Package

Prepared for: `{customer_name}`

This package is the customer-facing runtime bundle. It intentionally excludes
internal development, test, and license-administration files.

## First Run

1. Run `runtime\\setup_requirements.bat`
2. Confirm or place the license at `licenses\\active\\license.json`
3. Open `Tooling Control Panel\\Capacity_Optimizer_Control.xlsx`
4. Save the workbook after editing settings
5. Run `runtime\\run_optimizer.bat`

## Package Notes

- {license_text}
- {data_text}
- Machine fingerprint requests are written to `licenses\\requests\\`
- Customer instructions: `docs\\CUSTOMER_LICENSE_QUICKSTART_CN.md`
- IT deployment guide: `docs\\IT_DEPLOYMENT_CHECKLIST_CN.md`
"""
    (destination / "README.md").write_text(content, encoding="utf-8")


def _write_manifest(
    *,
    destination: Path,
    customer_name: str,
    package_name: str,
    license_included: bool,
    include_demo_data: bool,
) -> None:
    manifest = {
        "tool_name": TOOL_NAME,
        "customer_name": customer_name,
        "package_name": package_name,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "license_included": license_included,
        "demo_data_included": include_demo_data,
    }
    with open(destination / "delivery_manifest.json", "w", encoding="utf-8") as handle:
        json.dump(manifest, handle, ensure_ascii=False, indent=2)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Export a clean customer delivery package.")
    parser.add_argument("--customer-name", required=True, help="Customer name shown in the package output.")
    parser.add_argument(
        "--destination-root",
        default=str(DEFAULT_DELIVERY_ROOT),
        help="Folder where the delivery package directory will be created.",
    )
    parser.add_argument(
        "--package-name",
        default=None,
        help="Optional package directory name. Defaults to capacity_optimizer_<CustomerName>.",
    )
    parser.add_argument(
        "--license-file",
        default=None,
        help="Optional signed license.json to include under licenses\\active\\license.json.",
    )
    parser.add_argument(
        "--no-demo-data",
        action="store_true",
        help="Create an empty Data_Input folder instead of copying the bundled demo data.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite the destination package directory if it already exists.",
    )
    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    package_path = build_customer_package(
        project_root=REPO_ROOT,
        destination_root=Path(args.destination_root),
        customer_name=args.customer_name,
        package_name=args.package_name,
        license_file=Path(args.license_file) if args.license_file else None,
        include_demo_data=not args.no_demo_data,
        overwrite=args.overwrite,
    )
    print(f"Customer delivery package created: {package_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
