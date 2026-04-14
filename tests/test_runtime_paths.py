import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.runtime_paths import APP_NAME, ensure_workspace_dirs, resolve_runtime_paths


class RuntimePathsTests(unittest.TestCase):
    def test_source_mode_uses_repository_root_as_workspace(self):
        paths = resolve_runtime_paths()

        self.assertFalse(paths.is_frozen)
        self.assertEqual(paths.app_install_dir, Path(__file__).resolve().parents[1])
        self.assertEqual(paths.bundled_resources_dir, paths.app_install_dir)
        self.assertEqual(paths.bundled_docs_dir, paths.app_install_dir / "docs")
        self.assertEqual(paths.user_workspace_dir, paths.app_install_dir)
        self.assertEqual(paths.templates_dir, paths.app_install_dir / "Tooling Control Panel")
        self.assertEqual(paths.workspace_docs_dir, paths.app_install_dir / "docs")
        self.assertEqual(paths.outputs_dir, paths.app_install_dir / "output")
        self.assertEqual(paths.logs_dir, paths.app_install_dir / "logs")
        self.assertEqual(paths.license_active_dir, paths.app_install_dir / "licenses" / "active")
        self.assertEqual(paths.workspace_manifest_path, paths.app_install_dir / "workspace_manifest.json")
        self.assertEqual(paths.sample_data_dir, paths.app_install_dir / "Data_Input")

    def test_env_override_wins_in_source_mode(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict(os.environ, {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = resolve_runtime_paths()

            self.assertEqual(paths.user_workspace_dir, Path(tmpdir).resolve())
            self.assertEqual(paths.templates_dir, Path(tmpdir).resolve() / "Tooling Control Panel")
            self.assertEqual(paths.workspace_docs_dir, Path(tmpdir).resolve() / "docs")
            self.assertEqual(paths.outputs_dir, Path(tmpdir).resolve() / "output")

    def test_frozen_mode_uses_install_and_local_appdata_workspace(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            install_dir = Path(tmpdir) / "dist" / "CapacityOptimizer"
            resource_dir = install_dir / "resources"
            resource_dir.mkdir(parents=True, exist_ok=True)
            local_appdata = Path(tmpdir) / "LocalAppData"
            local_appdata.mkdir(parents=True, exist_ok=True)

            with patch("app.runtime_paths.is_frozen_runtime", return_value=True), patch(
                "app.runtime_paths.sys.executable",
                str(install_dir / "CapacityOptimizer.exe"),
            ), patch.dict(os.environ, {"LOCALAPPDATA": str(local_appdata)}, clear=False):
                paths = resolve_runtime_paths()

            self.assertTrue(paths.is_frozen)
            self.assertEqual(paths.app_install_dir, install_dir.resolve())
            self.assertEqual(paths.bundled_resources_dir, resource_dir.resolve())
            self.assertEqual(paths.user_workspace_dir, (local_appdata / APP_NAME).resolve())
            self.assertEqual(paths.sample_data_dir, resource_dir.resolve() / "Data_Input")

    def test_ensure_workspace_dirs_creates_user_writable_structure(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict(os.environ, {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = ensure_workspace_dirs(resolve_runtime_paths())

            self.assertTrue(paths.templates_dir.exists())
            self.assertTrue(paths.workspace_docs_dir.exists())
            self.assertTrue(paths.outputs_dir.exists())
            self.assertTrue(paths.logs_dir.exists())
            self.assertTrue(paths.license_dir.exists())
            self.assertTrue(paths.license_active_dir.exists())
            self.assertTrue(paths.license_requests_dir.exists())


if __name__ == "__main__":
    unittest.main()
