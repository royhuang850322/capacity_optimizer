import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.desktop_launcher import generate_machine_fingerprint_request, run_optimizer_from_launcher
from app.models import Config
from app.runtime_paths import resolve_runtime_paths
from app.workspace_init import initialize_user_workspace


class DesktopLauncherTests(unittest.TestCase):
    def test_initialize_user_workspace_creates_workbook_and_sample_data(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            fake_install = Path(tmpdir) / "dist" / "CapacityOptimizer"
            bundled_data = fake_install / "resources" / "Data_Input"
            bundled_docs = fake_install / "resources" / "docs"
            bundled_data.mkdir(parents=True, exist_ok=True)
            bundled_docs.mkdir(parents=True, exist_ok=True)
            (bundled_data / "planner1_load.csv").write_text(
                "Month,PlannerName,Product,ProductFamily,Plant,Forecast_Tons\n"
                "2026-01,PlannerA,P1,F1,PLT1,10\n",
                encoding="utf-8",
            )
            (bundled_data / "master_capacity.csv").write_text(
                "Product,WorkCenter,Annual_Capacity_Tons,Utilization_Target\n"
                "P1,WC1,120,1\n",
                encoding="utf-8",
            )
            (bundled_docs / "CUSTOMER_LICENSE_QUICKSTART_CN.md").write_text("customer quickstart", encoding="utf-8")
            (bundled_docs / "IT_DEPLOYMENT_CHECKLIST_CN.md").write_text("it checklist", encoding="utf-8")
            (bundled_docs / "desktop_launcher_usage.md").write_text("desktop launcher", encoding="utf-8")
            (bundled_docs / "PYTHON_INSTALL_GUIDE_CN.md").write_text("python guide", encoding="utf-8")
            local_appdata = Path(tmpdir) / "LocalAppData"
            local_appdata.mkdir(parents=True, exist_ok=True)

            with patch("app.runtime_paths.is_frozen_runtime", return_value=True), patch(
                "app.runtime_paths.sys.executable",
                str(fake_install / "CapacityOptimizer.exe"),
            ), patch.dict("os.environ", {"LOCALAPPDATA": str(local_appdata)}, clear=False):
                result = initialize_user_workspace(resolve_runtime_paths())

            self.assertTrue(result.workbook_created)
            self.assertTrue(result.sample_data_copied)
            self.assertTrue(result.paths.control_workbook_path.exists())
            self.assertTrue((result.paths.workspace_input_dir / "planner1_load.csv").exists())
            self.assertTrue((result.paths.workspace_input_dir / "master_capacity.csv").exists())
            self.assertTrue((result.paths.workspace_docs_dir / "CUSTOMER_LICENSE_QUICKSTART_CN.md").exists())
            self.assertTrue(result.paths.workspace_manifest_path.exists())
            manifest = json.loads(result.paths.workspace_manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["workspace_dir"], str(result.paths.user_workspace_dir))
            self.assertEqual(manifest["install_dir"], str(result.paths.app_install_dir))

    def test_initialize_user_workspace_does_not_overwrite_existing_user_files(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict("os.environ", {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = resolve_runtime_paths()
                initialize_user_workspace(paths)
                paths.control_workbook_path.write_text("user workbook", encoding="utf-8")
                user_data = paths.workspace_input_dir / "planner1_load.csv"
                user_data.write_text("custom planner file", encoding="utf-8")
                customer_doc = paths.workspace_docs_dir / "CUSTOMER_LICENSE_QUICKSTART_CN.md"
                customer_doc.write_text("custom doc", encoding="utf-8")

                result = initialize_user_workspace(paths)

            self.assertFalse(result.workbook_created)
            self.assertFalse(result.sample_data_copied)
            self.assertEqual(paths.control_workbook_path.read_text(encoding="utf-8"), "user workbook")
            self.assertEqual(user_data.read_text(encoding="utf-8"), "custom planner file")
            self.assertEqual(customer_doc.read_text(encoding="utf-8"), "custom doc")

    def test_run_optimizer_from_launcher_writes_log_and_reports_success(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict("os.environ", {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = resolve_runtime_paths()
                initialize_user_workspace(paths)
                runtime_config = self._build_runtime_config(paths)

                def fake_executor(config, *, runtime_paths, input_template):
                    self.assertEqual(config.run_mode, "ModeA")
                    self.assertEqual(config.input_load_folder, str(paths.workspace_input_dir))
                    self.assertIsNone(input_template)
                    self.assertEqual(runtime_paths.user_workspace_dir, paths.user_workspace_dir)

                result = run_optimizer_from_launcher(paths, runtime_config=runtime_config, run_executor=fake_executor)

            self.assertTrue(result.success)
            self.assertTrue(result.log_path.exists())
            log_text = result.log_path.read_text(encoding="utf-8")
            self.assertIn("[Launcher] Run started", log_text)
            self.assertIn("Running with launcher settings", log_text)

    def test_run_optimizer_from_launcher_initializes_workspace_if_missing(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict("os.environ", {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = resolve_runtime_paths()
                runtime_config = self._build_runtime_config(paths)

                def fake_executor(config, *, runtime_paths, input_template):
                    self.assertEqual(config.run_mode, "ModeA")
                    self.assertIsNone(input_template)
                    self.assertEqual(runtime_paths.user_workspace_dir, paths.user_workspace_dir)

                result = run_optimizer_from_launcher(paths, runtime_config=runtime_config, run_executor=fake_executor)

            self.assertTrue(result.success)
            self.assertTrue(paths.workspace_input_dir.exists())
            self.assertTrue(paths.logs_dir.exists())

    def test_generate_machine_fingerprint_request_writes_workspace_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict("os.environ", {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = resolve_runtime_paths()
                initialize_user_workspace(paths)
                with patch("app.desktop_launcher.build_machine_identity_payload") as payload_builder:
                    payload_builder.return_value = {
                        "product_name": "Chemical Capacity Optimizer",
                        "generated_at": "2026-04-13 09:00:00",
                        "machine_label": "TEST-PC",
                        "machine_fingerprint": "sha256:test",
                    }
                    request_path = generate_machine_fingerprint_request(paths)

            self.assertTrue(request_path.exists())
            self.assertEqual(request_path.parent, paths.license_requests_dir)
            self.assertIn("machine_fingerprint_TEST-PC_", request_path.name)

    def test_run_optimizer_from_launcher_supports_legacy_cli_runner(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict("os.environ", {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = resolve_runtime_paths()
                initialize_user_workspace(paths)

                def fake_runner(*, args, standalone_mode):
                    self.assertFalse(standalone_mode)
                    self.assertEqual(args, ["--input-template", str(paths.control_workbook_path)])

                result = run_optimizer_from_launcher(paths, cli_runner=fake_runner)

            self.assertTrue(result.success)
            self.assertTrue(result.log_path.exists())
            self.assertIn("[Launcher] Running legacy workbook mode.", result.log_path.read_text(encoding="utf-8"))

    @staticmethod
    def _build_runtime_config(paths) -> Config:
        return Config(
            project_root_folder=str(paths.user_workspace_dir),
            input_load_folder=str(paths.workspace_input_dir),
            input_master_folder=str(paths.workspace_input_dir),
            output_folder=str(paths.outputs_dir),
            output_file_name="test_result.xlsx",
            scenario_name="Baseline",
            start_month="2026-01",
            horizon_months=1,
            run_mode="ModeA",
            direct_mode=True,
            verbose=False,
            skip_validation_errors=False,
        )


if __name__ == "__main__":
    unittest.main()
