# License Tools

These scripts are for RSCP internal use. Do not ship private keys to customers.

## GUI launcher

For day-to-day use, the easiest entry point is:

```text
license_admin\open_license_generator.bat
```

It opens a local desktop form where you can:

- choose `Trial / Unbound` or `Manual / Custom`
- fill customer information
- load the fingerprint request file generated under `licenses\requests\` when needed
- click `Generate License`

The GUI uses the same signing logic as the command-line scripts below.

## Managed repository structure

By default, internal license artifacts are stored under:

```text
D:\RSCP_License_Admin\<CustomerName>\capacity_optimizer\
```

and then split into:

- `requests`
- `issued`
- `active`
- `archive`
- `notes`

If you need a different root, set the environment variable `RSCP_LICENSE_ADMIN_ROOT`
or choose another folder in the GUI.

## 1. Short-term trial license

Use this flow when you want to send a portable evaluation package without binding
the license to one computer.

```powershell
python license_admin\license_tools\generate_trial_license.py `
  --private-key license_admin\private_keys\license_signing_ed25519_private.pem `
  --license-id LIC-TRIAL-2026-0001 `
  --customer-name "Demo Customer" `
  --customer-id "DEMO-001" `
  --days-valid 14 `
  --note "14-day trial"
```

Result:

- `license_type = trial`
- `binding_mode = unbound`
- no machine fingerprint is required
- default output goes to `...\issued\`
- the same license is also copied to `...\active\license.json`

## 2. Machine-locked commercial license

Use this flow when the customer sends back `machine_fingerprint.json` from
`runtime\get_machine_fingerprint.bat`.

```powershell
python license_admin\license_tools\generate_license.py `
  --private-key license_admin\private_keys\license_signing_ed25519_private.pem `
  --license-id LIC-COMM-2026-0001 `
  --license-type commercial `
  --customer-name "ABC Chemical" `
  --customer-id "ABC-001" `
  --issue-date 2026-03-29 `
  --expiry-date 2027-03-28 `
  --binding-mode machine_locked `
  --machine-fingerprint "sha256:..." `
  --machine-label "ABC-LAPTOP-01" `
  --note "Commercial annual license"
```

Result:

- `license_type = commercial`
- `binding_mode = machine_locked`
- customer can run only on the licensed machine
- default output goes to `...\issued\`
- the same license is also copied to `...\active\license.json`
