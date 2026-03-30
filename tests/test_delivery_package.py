import json
import os
import tempfile
import unittest
from contextlib import contextmanager

from license_admin.export_customer_package import build_customer_package


TEST_TMP_ROOT = os.path.join(os.path.dirname(__file__), "_tmp")
os.makedirs(TEST_TMP_ROOT, exist_ok=True)


@contextmanager
def workspace_tempdir():
    with tempfile.TemporaryDirectory(dir=TEST_TMP_ROOT) as tmpdir:
        yield tmpdir


def _write_text(path: str, text: str = "demo") -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(text)


class DeliveryPackageTests(unittest.TestCase):
    def _build_fake_project(self, root: str) -> None:
        _write_text(os.path.join(root, "requirements.txt"), "openpyxl>=3.1.5\n")
        _write_text(os.path.join(root, "LICENSE"), "MIT\n")
        _write_text(os.path.join(root, "app", "main.py"), "print('app')\n")
        _write_text(os.path.join(root, "app", "create_template.py"), "print('template')\n")
        _write_text(os.path.join(root, "runtime", "run_optimizer.bat"), "@echo off\n")
        _write_text(os.path.join(root, "runtime", "setup_requirements.bat"), "@echo off\n")
        _write_text(os.path.join(root, "runtime", "get_machine_fingerprint.bat"), "@echo off\n")
        _write_text(os.path.join(root, "Data_Input", "planner1_load.csv"), "Month,PlannerName,Product,ProductFamily,Plant,Forecast_Tons\n")
        _write_text(os.path.join(root, "docs", "CUSTOMER_LICENSE_QUICKSTART_CN.md"), "customer quickstart\n")
        _write_text(os.path.join(root, "docs", "IT_DEPLOYMENT_CHECKLIST_CN.md"), "it checklist\n")

    def test_build_customer_package_creates_clean_structure(self):
        with workspace_tempdir() as tmpdir:
            project_root = os.path.join(tmpdir, "repo")
            destination_root = os.path.join(tmpdir, "packages")
            self._build_fake_project(project_root)

            package_path = build_customer_package(
                project_root=os.path.abspath(project_root),
                destination_root=os.path.abspath(destination_root),
                customer_name="DuPont",
                include_demo_data=True,
                overwrite=False,
            )

            self.assertTrue(os.path.isdir(package_path))
            self.assertTrue(os.path.exists(os.path.join(package_path, "app", "main.py")))
            self.assertTrue(os.path.exists(os.path.join(package_path, "runtime", "run_optimizer.bat")))
            self.assertTrue(os.path.exists(os.path.join(package_path, "Tooling Control Panel", "Capacity_Optimizer_Control.xlsx")))
            self.assertTrue(os.path.isdir(os.path.join(package_path, "licenses", "active")))
            self.assertTrue(os.path.isdir(os.path.join(package_path, "licenses", "requests")))
            self.assertTrue(os.path.exists(os.path.join(package_path, "docs", "CUSTOMER_LICENSE_QUICKSTART_CN.md")))
            self.assertTrue(os.path.exists(os.path.join(package_path, "docs", "IT_DEPLOYMENT_CHECKLIST_CN.md")))
            self.assertTrue(os.path.exists(os.path.join(package_path, "README.md")))
            self.assertTrue(os.path.exists(os.path.join(package_path, "delivery_manifest.json")))
            self.assertFalse(os.path.exists(os.path.join(package_path, "license_admin")))
            self.assertFalse(os.path.exists(os.path.join(package_path, "tests")))

    def test_build_customer_package_can_include_license_file(self):
        with workspace_tempdir() as tmpdir:
            project_root = os.path.join(tmpdir, "repo")
            destination_root = os.path.join(tmpdir, "packages")
            self._build_fake_project(project_root)
            license_source = os.path.join(tmpdir, "issued_license.json")
            _write_text(license_source, json.dumps({"license_id": "LIC-001"}, ensure_ascii=False))

            package_path = build_customer_package(
                project_root=os.path.abspath(project_root),
                destination_root=os.path.abspath(destination_root),
                customer_name="DuPont",
                license_file=os.path.abspath(license_source),
                overwrite=False,
            )

            with open(os.path.join(package_path, "licenses", "active", "license.json"), "r", encoding="utf-8") as handle:
                payload = json.load(handle)
            self.assertEqual(payload["license_id"], "LIC-001")


if __name__ == "__main__":
    unittest.main()
