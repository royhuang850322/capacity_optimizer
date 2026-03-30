import base64
import json
import os
import tempfile
import unittest
from contextlib import contextmanager
from datetime import date
from unittest.mock import patch

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

import license_validator
from license_validator import LicenseValidationError, validate_license


TEST_TMP_ROOT = os.path.join(os.path.dirname(__file__), "_tmp")
os.makedirs(TEST_TMP_ROOT, exist_ok=True)


@contextmanager
def workspace_tempdir():
    with tempfile.TemporaryDirectory(dir=TEST_TMP_ROOT) as tmpdir:
        yield tmpdir


def _signed_payload(private_key: Ed25519PrivateKey, **overrides):
    payload = {
        "license_version": 1,
        "product_name": "Chemical Capacity Optimizer",
        "license_id": "LIC-TEST-0001",
        "license_type": "trial",
        "customer_name": "Test Customer",
        "customer_id": "TEST-001",
        "issue_date": "2026-03-29",
        "expiry_date": "2026-06-30",
        "binding_mode": "unbound",
        "machine_fingerprint": "",
        "machine_label": "",
        "features": {"mode_a": True, "mode_b": True, "comparison_workbook": True},
        "note": "Test license",
    }
    payload.update(overrides)
    payload["signature"] = base64.b64encode(
        private_key.sign(license_validator._canonical_license_bytes(payload))
    ).decode("ascii")
    return payload


class LicenseValidatorTests(unittest.TestCase):
    def test_validate_unbound_license(self):
        private_key = Ed25519PrivateKey.generate()
        public_key = private_key.public_key()
        payload = _signed_payload(private_key)

        with workspace_tempdir() as tmpdir:
            license_path = os.path.join(tmpdir, "license.json")
            with open(license_path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=2)

            with patch("license_validator._load_public_key", return_value=public_key):
                info = validate_license(tmpdir, today=date(2026, 4, 1))

        self.assertEqual(info.status, "Valid")
        self.assertEqual(info.license_id, "LIC-TEST-0001")
        self.assertEqual(info.customer_name, "Test Customer")
        self.assertEqual(info.binding_mode, "unbound")

    def test_validate_expired_license_raises(self):
        private_key = Ed25519PrivateKey.generate()
        public_key = private_key.public_key()
        payload = _signed_payload(private_key, expiry_date="2026-03-31")

        with workspace_tempdir() as tmpdir:
            license_path = os.path.join(tmpdir, "license.json")
            with open(license_path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=2)

            with patch("license_validator._load_public_key", return_value=public_key):
                with self.assertRaises(LicenseValidationError):
                    validate_license(tmpdir, today=date(2026, 4, 1))

    def test_validate_machine_locked_license_rejects_other_machine(self):
        private_key = Ed25519PrivateKey.generate()
        public_key = private_key.public_key()
        payload = _signed_payload(
            private_key,
            binding_mode="machine_locked",
            machine_fingerprint="sha256:expectedmachinefingerprint",
            machine_label="LOCKED-PC",
        )

        with workspace_tempdir() as tmpdir:
            license_path = os.path.join(tmpdir, "license.json")
            with open(license_path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=2)

            with patch("license_validator._load_public_key", return_value=public_key):
                with patch("license_validator.get_machine_fingerprint", return_value="sha256:othermachine"):
                    with self.assertRaises(LicenseValidationError):
                        validate_license(tmpdir, today=date(2026, 4, 1))


if __name__ == "__main__":
    unittest.main()
