import json
import os
import tempfile
import unittest
from contextlib import contextmanager

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives import serialization

from license_admin.license_tools.common import (
    activate_issued_license,
    build_issued_license_path,
    copy_machine_request_to_admin,
    create_signed_license,
    create_signed_trial_license,
    ensure_customer_tool_dirs,
    load_machine_identity_json,
)


TEST_TMP_ROOT = os.path.join(os.path.dirname(__file__), "_tmp")
os.makedirs(TEST_TMP_ROOT, exist_ok=True)


@contextmanager
def workspace_tempdir():
    with tempfile.TemporaryDirectory(dir=TEST_TMP_ROOT) as tmpdir:
        yield tmpdir


def _write_private_key(path: str) -> None:
    private_key = Ed25519PrivateKey.generate()
    pem = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )
    with open(path, "wb") as handle:
        handle.write(pem)


class LicenseToolTests(unittest.TestCase):
    def test_create_signed_trial_license_writes_unbound_license(self):
        with workspace_tempdir() as tmpdir:
            private_key = os.path.join(tmpdir, "private.pem")
            output_path = os.path.join(tmpdir, "license.json")
            _write_private_key(private_key)

            payload = create_signed_trial_license(
                private_key_path=private_key,
                out_path=output_path,
                license_id="LIC-TRIAL-001",
                customer_name="Trial Customer",
                customer_id="TRIAL-001",
                days_valid=14,
                issue_date="2026-03-29",
            )

            self.assertEqual(payload["binding_mode"], "unbound")
            self.assertEqual(payload["license_type"], "trial")
            self.assertTrue(os.path.exists(output_path))

    def test_create_signed_machine_locked_license_requires_fingerprint(self):
        with workspace_tempdir() as tmpdir:
            private_key = os.path.join(tmpdir, "private.pem")
            output_path = os.path.join(tmpdir, "license.json")
            _write_private_key(private_key)

            with self.assertRaises(ValueError):
                create_signed_license(
                    private_key_path=private_key,
                    out_path=output_path,
                    license_id="LIC-COMM-001",
                    license_type="commercial",
                    customer_name="ABC",
                    customer_id="ABC-001",
                    issue_date="2026-03-29",
                    expiry_date="2027-03-28",
                    binding_mode="machine_locked",
                )

    def test_load_machine_identity_json_reads_expected_fields(self):
        with workspace_tempdir() as tmpdir:
            payload_path = os.path.join(tmpdir, "machine_fingerprint.json")
            with open(payload_path, "w", encoding="utf-8") as handle:
                json.dump(
                    {
                        "machine_label": "DEMO-PC",
                        "machine_fingerprint": "sha256:abcdef",
                    },
                    handle,
                )

            payload = load_machine_identity_json(payload_path)

            self.assertEqual(payload["machine_label"], "DEMO-PC")
            self.assertEqual(payload["machine_fingerprint"], "sha256:abcdef")

    def test_managed_license_paths_create_customer_tool_structure(self):
        with workspace_tempdir() as tmpdir:
            dirs = ensure_customer_tool_dirs("DuPont", admin_root=tmpdir, tool_name="capacity_optimizer")
            self.assertTrue(dirs["requests"].exists())
            self.assertTrue(dirs["issued"].exists())
            self.assertTrue(dirs["active"].exists())
            self.assertTrue(dirs["archive"].exists())

            issued_path = build_issued_license_path(
                "DuPont",
                "LIC-DUPONT-COMM-2026-0001",
                admin_root=tmpdir,
                tool_name="capacity_optimizer",
            )
            self.assertIn("DuPont", str(issued_path))
            self.assertIn("capacity_optimizer", str(issued_path))
            self.assertIn("issued", str(issued_path))

    def test_activate_issued_license_copies_to_active_and_archives_previous(self):
        with workspace_tempdir() as tmpdir:
            private_key = os.path.join(tmpdir, "private.pem")
            _write_private_key(private_key)

            first_path = build_issued_license_path("DuPont", "LIC-001", admin_root=tmpdir)
            second_path = build_issued_license_path("DuPont", "LIC-002", admin_root=tmpdir)

            create_signed_trial_license(
                private_key_path=private_key,
                out_path=str(first_path),
                license_id="LIC-001",
                customer_name="DuPont",
                customer_id="DUP-001",
                days_valid=7,
                issue_date="2026-03-30",
            )
            active_path = activate_issued_license(first_path, "DuPont", admin_root=tmpdir)
            self.assertTrue(active_path.exists())

            create_signed_trial_license(
                private_key_path=private_key,
                out_path=str(second_path),
                license_id="LIC-002",
                customer_name="DuPont",
                customer_id="DUP-001",
                days_valid=7,
                issue_date="2026-03-30",
            )
            activate_issued_license(second_path, "DuPont", admin_root=tmpdir)

            archive_dir = ensure_customer_tool_dirs("DuPont", admin_root=tmpdir)["archive"]
            self.assertTrue(any(p.name.startswith("LIC-001") for p in archive_dir.iterdir()))

    def test_copy_machine_request_to_admin_stores_request_file(self):
        with workspace_tempdir() as tmpdir:
            source_path = os.path.join(tmpdir, "machine_fingerprint.json")
            with open(source_path, "w", encoding="utf-8") as handle:
                json.dump(
                    {
                        "machine_label": "DUPONT-PC01",
                        "machine_fingerprint": "sha256:abcdef",
                    },
                    handle,
                )

            dest = copy_machine_request_to_admin(
                source_path,
                "DuPont",
                admin_root=tmpdir,
                machine_label="DUPONT-PC01",
            )

            self.assertTrue(dest.exists())
            self.assertIn("requests", str(dest))
            self.assertIn("DUPONT-PC01", dest.name)


if __name__ == "__main__":
    unittest.main()
