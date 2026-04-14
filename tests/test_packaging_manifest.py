import unittest
from pathlib import Path

from build_support.packaging_manifest import (
    APP_NAME,
    ENTRY_SCRIPT,
    HIDDEN_IMPORT_PACKAGES,
    RESOURCE_DIR_MAPPINGS,
    RESOURCE_FILE_MAPPINGS,
    iter_data_mappings,
)


class PackagingManifestTests(unittest.TestCase):
    def test_entry_script_matches_desktop_launcher(self):
        self.assertEqual(ENTRY_SCRIPT, "CapacityOptimizerLauncher.pyw")
        self.assertTrue((Path(__file__).resolve().parents[1] / ENTRY_SCRIPT).exists())

    def test_resource_manifest_includes_runtime_assets(self):
        project_root = Path(__file__).resolve().parents[1]
        mappings = iter_data_mappings(project_root)
        mapping_text = "\n".join(f"{source} -> {target}" for source, target in mappings)

        self.assertIn("Data_Input", mapping_text)
        self.assertIn("resources/docs", mapping_text)
        self.assertIn(APP_NAME, "CapacityOptimizer")
        self.assertIn("ortools", HIDDEN_IMPORT_PACKAGES)
        self.assertTrue(any(relative == "Data_Input" for relative, _ in RESOURCE_DIR_MAPPINGS))
        self.assertTrue(any(relative == "docs/desktop_launcher_usage.md" for relative, _ in RESOURCE_FILE_MAPPINGS))


if __name__ == "__main__":
    unittest.main()
