"""
Internal Tkinter GUI for exporting clean customer delivery packages without
typing PowerShell commands.
"""
from __future__ import annotations

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

BOOTSTRAP_ROOT = Path(__file__).resolve().parents[1]
if str(BOOTSTRAP_ROOT) not in sys.path:
    sys.path.insert(0, str(BOOTSTRAP_ROOT))

from app.runtime_paths import resolve_runtime_paths

PROJECT_ROOT = resolve_runtime_paths().app_install_dir
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from license_admin.export_customer_package import DEFAULT_DELIVERY_ROOT, TOOL_NAME, build_customer_package
from license_admin.license_tools.common import sanitize_path_component


class DeliveryExporterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Chemical Capacity Optimizer - Delivery Package Exporter")
        self.root.geometry("980x720")
        self.root.minsize(940, 680)

        self.customer_name_var = tk.StringVar()
        self.destination_root_var = tk.StringVar(value=str(DEFAULT_DELIVERY_ROOT))
        self.package_name_var = tk.StringVar()
        self.license_file_var = tk.StringVar()
        self.include_demo_data_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)

        self.package_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Fill in the form and click Export Package.")
        self.last_exported_path: Path | None = None

        self._build_ui()
        self._bind_events()
        self._refresh_preview()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        container = ttk.Frame(self.root, padding=14)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)

        ttk.Label(
            container,
            text="Chemical Capacity Optimizer - Internal Delivery Package Exporter",
            font=("Segoe UI", 15, "bold"),
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            container,
            text=(
                "This tool creates a clean customer-facing runtime package. "
                "It copies only the runtime files, regenerates a fresh control workbook, "
                "and can optionally include a signed license file."
            ),
            wraplength=920,
        ).grid(row=1, column=0, sticky="w", pady=(4, 10))

        form_frame = ttk.LabelFrame(container, text="Export Settings", padding=10)
        form_frame.grid(row=2, column=0, sticky="ew")
        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="Customer Name").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form_frame, textvariable=self.customer_name_var).grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)

        ttk.Label(form_frame, text="Destination Root").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form_frame, textvariable=self.destination_root_var).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(form_frame, text="Browse...", command=self._browse_destination_root).grid(row=1, column=2, padx=(8, 0), pady=4)

        ttk.Label(form_frame, text="Package Name").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form_frame, textvariable=self.package_name_var).grid(row=2, column=1, columnspan=2, sticky="ew", pady=4)

        ttk.Label(form_frame, text="License File (Optional)").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form_frame, textvariable=self.license_file_var).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Button(form_frame, text="Browse...", command=self._browse_license_file).grid(row=3, column=2, padx=(8, 0), pady=4)

        ttk.Label(form_frame, text="Package Path Preview").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form_frame, textvariable=self.package_path_var, state="readonly").grid(
            row=4, column=1, columnspan=2, sticky="ew", pady=4
        )

        options_frame = ttk.LabelFrame(container, text="Options", padding=10)
        options_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        options_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            options_frame,
            text="Include demo data in Data_Input",
            variable=self.include_demo_data_var,
        ).grid(row=0, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            options_frame,
            text="Overwrite existing package folder if it already exists",
            variable=self.overwrite_var,
        ).grid(row=1, column=0, sticky="w", pady=2)

        notes_frame = ttk.LabelFrame(container, text="What This Export Includes", padding=10)
        notes_frame.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(
            notes_frame,
            text=(
                "Included: app, runtime, Tooling Control Panel, docs for customers, Data_Input "
                "(optional), output, licenses folders, README, LICENSE, requirements.txt.\n"
                "Excluded: tests, internal license administration files, private keys, and internal SOP files."
            ),
            wraplength=920,
        ).grid(row=0, column=0, sticky="w")

        actions_frame = ttk.Frame(container)
        actions_frame.grid(row=5, column=0, sticky="ew", pady=(12, 0))

        ttk.Button(actions_frame, text="Export Package", command=self._export_package).grid(row=0, column=0, sticky="w")
        ttk.Button(actions_frame, text="Open Exported Package", command=self._open_exported_package).grid(
            row=0, column=1, padx=(10, 0), sticky="w"
        )
        ttk.Button(actions_frame, text="Open Delivery Root", command=self._open_delivery_root).grid(
            row=0, column=2, padx=(10, 0), sticky="w"
        )
        ttk.Button(actions_frame, text="Clear Form", command=self._clear_form).grid(
            row=0, column=3, padx=(10, 0), sticky="w"
        )

        status_frame = ttk.LabelFrame(container, text="Status", padding=10)
        status_frame.grid(row=6, column=0, sticky="ew", pady=(12, 0))
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status_var, wraplength=900, foreground="#1f4e79").grid(
            row=0, column=0, sticky="w"
        )

    def _bind_events(self) -> None:
        self.customer_name_var.trace_add("write", lambda *_args: self._refresh_preview())
        self.destination_root_var.trace_add("write", lambda *_args: self._refresh_preview())
        self.package_name_var.trace_add("write", lambda *_args: self._refresh_preview())

    def _browse_destination_root(self) -> None:
        selected = filedialog.askdirectory(
            title="Select delivery package root",
            initialdir=self.destination_root_var.get().strip() or str(PROJECT_ROOT),
        )
        if selected:
            self.destination_root_var.set(selected)

    def _browse_license_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select signed license.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=self.license_file_var.get().strip() or str(PROJECT_ROOT),
        )
        if selected:
            self.license_file_var.set(selected)

    def _refresh_preview(self) -> None:
        destination_root = Path(self.destination_root_var.get().strip() or DEFAULT_DELIVERY_ROOT)
        customer_name = self.customer_name_var.get().strip()
        package_name = self.package_name_var.get().strip()
        if package_name:
            preview_name = sanitize_path_component(package_name, f"{TOOL_NAME}_PACKAGE")
        else:
            preview_name = f"{TOOL_NAME}_{sanitize_path_component(customer_name, 'CUSTOMER_NAME')}"
        self.package_path_var.set(str(destination_root / preview_name))

    def _clear_form(self) -> None:
        self.customer_name_var.set("")
        self.destination_root_var.set(str(DEFAULT_DELIVERY_ROOT))
        self.package_name_var.set("")
        self.license_file_var.set("")
        self.include_demo_data_var.set(True)
        self.overwrite_var.set(False)
        self.last_exported_path = None
        self.status_var.set("Form cleared.")
        self._refresh_preview()

    def _open_exported_package(self) -> None:
        if self.last_exported_path is None:
            messagebox.showinfo(
                "No package yet",
                "Export a package first, then this button will open the generated folder.",
                parent=self.root,
            )
            return
        os.startfile(str(self.last_exported_path))

    def _open_delivery_root(self) -> None:
        destination_root = Path(self.destination_root_var.get().strip() or DEFAULT_DELIVERY_ROOT)
        destination_root.mkdir(parents=True, exist_ok=True)
        os.startfile(str(destination_root))

    def _export_package(self) -> None:
        customer_name = self.customer_name_var.get().strip()
        if not customer_name:
            messagebox.showerror("Missing customer", "Customer Name is required.", parent=self.root)
            return

        destination_root = self.destination_root_var.get().strip()
        if not destination_root:
            messagebox.showerror("Missing destination", "Destination Root is required.", parent=self.root)
            return

        license_file = self.license_file_var.get().strip() or None
        if license_file and not Path(license_file).exists():
            messagebox.showerror("Missing license file", f"License file not found:\n{license_file}", parent=self.root)
            return

        try:
            package_path = build_customer_package(
                project_root=PROJECT_ROOT,
                destination_root=Path(destination_root),
                customer_name=customer_name,
                package_name=self.package_name_var.get().strip() or None,
                license_file=license_file,
                include_demo_data=self.include_demo_data_var.get(),
                overwrite=self.overwrite_var.get(),
            )
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc), parent=self.root)
            self.status_var.set(f"Export failed: {exc}")
            return

        self.last_exported_path = package_path
        self._refresh_preview()
        success_message = (
            f"Customer delivery package created successfully.\n\n"
            f"Customer: {customer_name}\n"
            f"Package folder: {package_path}\n"
            f"Demo data included: {'Yes' if self.include_demo_data_var.get() else 'No'}\n"
            f"License included: {'Yes' if license_file else 'No'}"
        )
        self.status_var.set(success_message.replace("\n", " "))
        messagebox.showinfo("Delivery package created", success_message, parent=self.root)


def main() -> int:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    DeliveryExporterApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
