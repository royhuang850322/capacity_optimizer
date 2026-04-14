from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path


def _load_verify_dist_layout():
    module_path = Path(__file__).resolve().parents[1] / "packaging" / "verify_onefolder_dist.py"
    spec = importlib.util.spec_from_file_location("verify_onefolder_dist", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load packaging verifier module from {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.verify_dist_layout


verify_dist_layout = _load_verify_dist_layout()


class PackagingVerifierTests(unittest.TestCase):
    def test_verify_accepts_top_level_resources_layout(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dist_root = Path(temp_dir) / "CapacityOptimizer"
            (dist_root / "resources" / "Data_Input").mkdir(parents=True)
            (dist_root / "resources" / "docs").mkdir(parents=True)
            (dist_root / "CapacityOptimizer.exe").write_text("", encoding="utf-8")
            verify_dist_layout(dist_root)

    def test_verify_accepts_internal_resources_layout(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dist_root = Path(temp_dir) / "CapacityOptimizer"
            (dist_root / "_internal" / "resources" / "Data_Input").mkdir(parents=True)
            (dist_root / "_internal" / "resources" / "docs").mkdir(parents=True)
            (dist_root / "CapacityOptimizer.exe").write_text("", encoding="utf-8")
            verify_dist_layout(dist_root)


if __name__ == "__main__":
    unittest.main()
