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


def verify_dist_layout(dist_root: Path) -> None:
    resource_root = _resolve_resource_root(dist_root)
    required_paths = [
        dist_root / "CapacityOptimizer.exe",
        resource_root / "Data_Input",
        resource_root / "docs",
    ]
    missing = [path for path in required_paths if not path.exists()]
    if missing:
        missing_text = "\n".join(str(path) for path in missing)
        raise SystemExit(f"Packaged layout verification failed. Missing:\n{missing_text}")
    print(f"Verified one-folder layout: {dist_root}")
    print(f"Bundled resources root: {resource_root}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify the one-folder dist layout for CapacityOptimizer.")
    parser.add_argument(
        "--dist-root",
        default="dist/CapacityOptimizer",
        help="Path to the one-folder application root.",
    )
    args = parser.parse_args()
    verify_dist_layout(Path(args.dist_root).resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
