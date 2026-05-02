"""
Minimal layout verification for the PyInstaller one-folder output.
"""
from __future__ import annotations

import argparse
from pathlib import Path


def _resolve_resource_root(dist_root: Path) -> Path:
    candidates = [
        dist_root / "resources",
        dist_root / "_internal" / "resources",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    candidate_text = "\n".join(str(path) for path in candidates)
    raise SystemExit(
        "Packaged layout verification failed. Could not find bundled resources in any of:\n"
        f"{candidate_text}"
    )


def verify_dist_layout(
    dist_root: Path,
    *,
    app_name: str,
    required_resource_subpaths: tuple[str, ...] = (),
) -> None:
    required_paths = [dist_root / f"{app_name}.exe"]
    resource_root: Path | None = None
    if required_resource_subpaths:
        resource_root = _resolve_resource_root(dist_root)
        required_paths.extend(resource_root / subpath for subpath in required_resource_subpaths)
    missing = [path for path in required_paths if not path.exists()]
    if missing:
        missing_text = "\n".join(str(path) for path in missing)
        raise SystemExit(f"Packaged layout verification failed. Missing:\n{missing_text}")
    print(f"Verified one-folder layout: {dist_root}")
    if resource_root is not None:
        print(f"Bundled resources root: {resource_root}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify the one-folder dist layout for a packaged desktop app.")
    parser.add_argument(
        "--dist-root",
        default="dist/CapacityOptimizer",
        help="Path to the one-folder application root.",
    )
    parser.add_argument(
        "--app-name",
        default="CapacityOptimizer",
        help="Expected executable name without the .exe suffix.",
    )
    parser.add_argument(
        "--require-resource-subpath",
        action="append",
        default=[],
        help="Relative path expected under the bundled resources root. Repeat as needed.",
    )
    args = parser.parse_args()
    verify_dist_layout(
        Path(args.dist_root).resolve(),
        app_name=args.app_name,
        required_resource_subpaths=tuple(args.require_resource_subpath),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
