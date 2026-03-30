import os
import unittest

from license_admin.delivery_exporter_ui import PROJECT_ROOT


class DeliveryExporterPathTests(unittest.TestCase):
    def test_project_root_points_to_repository_root(self):
        self.assertTrue(os.path.exists(os.path.join(PROJECT_ROOT, "README.md")))
        self.assertTrue(os.path.exists(os.path.join(PROJECT_ROOT, "license_admin", "export_customer_package.py")))


if __name__ == "__main__":
    unittest.main()
