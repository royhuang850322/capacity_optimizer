"""
Offline signed-license validation for the Chemical Capacity Optimizer.
"""
from __future__ import annotations

import base64
import json
import os
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

from machine_fingerprint import get_machine_fingerprint, get_machine_label


PRODUCT_NAME = "Chemical Capacity Optimizer"
LICENSE_FILENAME = "license.json"
LICENSE_VERSION = 1
PUBLIC_KEY_PEM = b"""-----BEGIN PUBLIC KEY-----
MCowBQYDK2VwAyEAGtlVAoaB7aq+Dx6gQ70uZ+mIRR2WPK6spcFCcQoq+pw=
-----END PUBLIC KEY-----
"""


class LicenseValidationError(RuntimeError):
    """Raised when the local license is missing, invalid, expired, or mismatched."""


@dataclass(frozen=True)
class LicenseInfo:
    license_id: str
    license_type: str
    customer_name: str
    customer_id: str
    issue_date: str
    expiry_date: str
    binding_mode: str
    machine_fingerprint: str
    machine_label: str
    note: str
    features: dict[str, Any]
    license_path: str
    status: str = "Valid"


def _license_path(project_root: str) -> str:
    return os.path.join(project_root, LICENSE_FILENAME)


def _load_license_payload(path: str) -> dict[str, Any]:
    if not os.path.exists(path):
        raise LicenseValidationError(
            f"License file not found.\nExpected file: {path}\nPlease contact RSCP for a valid license file."
        )
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except Exception as exc:
        raise LicenseValidationError(f"Could not read license file: {path}") from exc
    if not isinstance(payload, dict):
        raise LicenseValidationError("License file format is invalid.")
    return payload


def _required_text(payload: dict[str, Any], key: str) -> str:
    value = payload.get(key)
    if value is None:
        raise LicenseValidationError(f"License field '{key}' is missing.")
    text = str(value).strip()
    if not text:
        raise LicenseValidationError(f"License field '{key}' is empty.")
    return text


def _parse_iso_date(raw_value: str, field_name: str) -> date:
    try:
        return datetime.strptime(raw_value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise LicenseValidationError(
            f"License field '{field_name}' must use YYYY-MM-DD format."
        ) from exc


def _canonical_license_bytes(payload: dict[str, Any]) -> bytes:
    signable = {key: value for key, value in payload.items() if key != "signature"}
    return json.dumps(
        signable,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=False,
    ).encode("utf-8")


def _load_public_key() -> Ed25519PublicKey:
    if b"PLACEHOLDER" in PUBLIC_KEY_PEM:
        raise LicenseValidationError(
            "Embedded license public key is not configured. Generate the signing key pair first."
        )
    return serialization.load_pem_public_key(PUBLIC_KEY_PEM)


def _verify_signature(payload: dict[str, Any]) -> None:
    signature_b64 = _required_text(payload, "signature")
    try:
        signature = base64.b64decode(signature_b64, validate=True)
    except Exception as exc:
        raise LicenseValidationError("License signature is not valid Base64 data.") from exc

    public_key = _load_public_key()
    try:
        public_key.verify(signature, _canonical_license_bytes(payload))
    except InvalidSignature as exc:
        raise LicenseValidationError(
            "License is invalid or has been modified.\nPlease contact RSCP for a valid license."
        ) from exc


def validate_license(project_root: str, today: date | None = None) -> LicenseInfo:
    payload = _load_license_payload(_license_path(project_root))

    version = int(payload.get("license_version", 0))
    if version != LICENSE_VERSION:
        raise LicenseValidationError(
            f"Unsupported license version: {version}. Expected {LICENSE_VERSION}."
        )

    product_name = _required_text(payload, "product_name")
    if product_name != PRODUCT_NAME:
        raise LicenseValidationError(
            f"License product mismatch: expected '{PRODUCT_NAME}', got '{product_name}'."
        )

    _verify_signature(payload)

    issue_date_text = _required_text(payload, "issue_date")
    expiry_date_text = _required_text(payload, "expiry_date")
    issue_date_value = _parse_iso_date(issue_date_text, "issue_date")
    expiry_date_value = _parse_iso_date(expiry_date_text, "expiry_date")
    current_date = today or date.today()
    if expiry_date_value < current_date:
        raise LicenseValidationError(
            f"License expired on {expiry_date_text}.\nPlease contact RSCP to renew the license."
        )
    if issue_date_value > expiry_date_value:
        raise LicenseValidationError("License issue_date is later than expiry_date.")

    binding_mode = _required_text(payload, "binding_mode").lower()
    if binding_mode not in {"unbound", "machine_locked"}:
        raise LicenseValidationError(
            "License binding_mode must be either 'unbound' or 'machine_locked'."
        )

    machine_fingerprint = str(payload.get("machine_fingerprint") or "").strip()
    machine_label = str(payload.get("machine_label") or "").strip()
    if binding_mode == "machine_locked":
        if not machine_fingerprint:
            raise LicenseValidationError(
                "Machine-locked license is missing machine_fingerprint."
            )
        current_fingerprint = get_machine_fingerprint()
        if current_fingerprint != machine_fingerprint:
            display_fingerprint = current_fingerprint.split(":", 1)[-1][:12]
            licensed_machine = machine_label or "UNKNOWN_MACHINE"
            raise LicenseValidationError(
                "This license is not valid for this computer.\n"
                f"Licensed machine: {licensed_machine}\n"
                f"Current machine fingerprint: {display_fingerprint}\n"
                "Please contact RSCP if you need to move the license."
            )
        if not machine_label:
            machine_label = get_machine_label()

    return LicenseInfo(
        license_id=_required_text(payload, "license_id"),
        license_type=_required_text(payload, "license_type"),
        customer_name=_required_text(payload, "customer_name"),
        customer_id=_required_text(payload, "customer_id"),
        issue_date=issue_date_text,
        expiry_date=expiry_date_text,
        binding_mode=binding_mode,
        machine_fingerprint=machine_fingerprint,
        machine_label=machine_label,
        note=str(payload.get("note") or "").strip(),
        features=payload.get("features") if isinstance(payload.get("features"), dict) else {},
        license_path=_license_path(project_root),
    )
