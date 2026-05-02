"""
Windows desktop launcher for the Capacity Optimizer.

Backend orchestration is unchanged; only the UI layer is modernized with PySide6.
"""
from __future__ import annotations

import json
import os
import sys
import threading
import traceback
from contextlib import redirect_stderr, redirect_stdout
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.machine_fingerprint import build_machine_identity_payload, sanitize_machine_label
from app.main import run_with_config
from app.models import Config
from app.data_loader import discover_planner_scenarios
from app.i18n import normalize_language, ui_text
from app.run_logging import RUN_LOG_PATH_ENV_VAR
from app.runtime_paths import RuntimePaths, ensure_workspace_dirs, resolve_runtime_paths, with_workspace_dir
from app.workspace_init import initialize_user_workspace

try:
    from PySide6.QtCore import QObject, Qt, QUrl, Signal
    from PySide6.QtGui import QDesktopServices, QFont, QPalette
    from PySide6.QtWidgets import (
        QAbstractItemView,
        QApplication,
        QComboBox,
        QFileDialog,
        QFormLayout,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QLabel,
        QLayout,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QScrollArea,
        QStackedWidget,
        QSplitter,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )

    PYSIDE6_AVAILABLE = True
except ModuleNotFoundError:
    PYSIDE6_AVAILABLE = False
    QObject = object  # type: ignore[assignment]
    Signal = None  # type: ignore[assignment]


APP_TITLE = "Capacity Optimizer"
SETTINGS_FILENAME = "launcher_settings.json"
APP_VERSION = "v1.1.3"

LIGHT_QSS = """
QWidget {
    background: #f4f6fa;
    color: #1d2530;
    font-family: "Segoe UI";
    font-size: 13px;
}
QFrame#Card {
    background: #ffffff;
    border: 1px solid #d9e1ec;
    border-radius: 12px;
}
QLabel#Title {
    font-size: 30px;
    font-weight: 700;
}
QLabel#Subtitle {
    color: #5b6a80;
    font-size: 14px;
}
QLabel#SectionTitle {
    font-size: 16px;
    font-weight: 600;
}
QLabel#StatusBadge {
    background: #e8eff9;
    border: 1px solid #d0dcee;
    border-radius: 10px;
    padding: 6px 10px;
    font-weight: 600;
}
QPushButton {
    background: #e9edf5;
    border: 1px solid #cfd8e6;
    border-radius: 8px;
    padding: 8px 12px;
}
QPushButton:hover {
    background: #dfe7f2;
}
QPushButton#Primary {
    background: #1a6fb3;
    color: white;
    border: 1px solid #14598f;
    font-size: 14px;
    font-weight: 700;
    padding: 12px 14px;
}
QPushButton#Primary:hover {
    background: #165f99;
}
QLineEdit, QComboBox, QTextEdit {
    background: #ffffff;
    border: 1px solid #ced8e8;
    border-radius: 8px;
    padding: 6px 8px;
}
QLineEdit:disabled {
    background: #f1f4f9;
    color: #54657f;
}
QSplitter::handle {
    background: #dce4ef;
}
"""

DARK_QSS = """
QWidget {
    background: #151b23;
    color: #e8eef7;
    font-family: "Segoe UI";
    font-size: 13px;
}
QFrame#Card {
    background: #202a36;
    border: 1px solid #334357;
    border-radius: 12px;
}
QLabel#Title {
    font-size: 30px;
    font-weight: 700;
}
QLabel#Subtitle {
    color: #9ba9bf;
    font-size: 14px;
}
QLabel#SectionTitle {
    font-size: 16px;
    font-weight: 600;
}
QLabel#StatusBadge {
    background: #2d3d52;
    border: 1px solid #455f7d;
    border-radius: 10px;
    padding: 6px 10px;
    font-weight: 600;
}
QPushButton {
    background: #2a3645;
    color: #e8eef7;
    border: 1px solid #3a4d63;
    border-radius: 8px;
    padding: 8px 12px;
}
QPushButton:hover {
    background: #334357;
}
QPushButton#Primary {
    background: #1d79c4;
    color: white;
    border: 1px solid #1765a4;
    font-size: 14px;
    font-weight: 700;
    padding: 12px 14px;
}
QPushButton#Primary:hover {
    background: #1a6db0;
}
QLineEdit, QComboBox, QTextEdit {
    background: #1c2531;
    color: #eef4fc;
    border: 1px solid #3a4d63;
    border-radius: 8px;
    padding: 6px 8px;
}
QLineEdit:disabled {
    background: #253142;
    color: #9eb0ca;
}
QSplitter::handle {
    background: #3a4d63;
}
"""


@dataclass(frozen=True)
class LauncherRunResult:
    success: bool
    log_path: Path
    message: str


def _default_settings(paths: RuntimePaths) -> dict[str, str]:
    now = datetime.now()
    return {
        "workspace_root": str(paths.user_workspace_dir),
        "output_file_name": "capacity_result.xlsx",
        "scenario_name": "Baseline",
        "start_year": str(now.year),
        "start_month": str(now.month),
        "horizon_months": "12",
        "run_mode": "ModeB",
        "verbose": "No",
        "skip_validation_errors": "No",
        "theme": "System",
        "language": "en",
    }


def _to_bool(value: str | bool | int | None) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}


def generate_machine_fingerprint_request(paths: RuntimePaths | None = None) -> Path:
    runtime_paths = ensure_workspace_dirs(paths or resolve_runtime_paths())
    payload = build_machine_identity_payload()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"machine_fingerprint_{sanitize_machine_label(payload['machine_label'])}_{timestamp}.json"
    output_path = runtime_paths.license_requests_dir / filename
    with open(output_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    return output_path


def _settings_path(paths: RuntimePaths) -> Path:
    return paths.app_install_dir / SETTINGS_FILENAME


def load_launcher_settings(paths: RuntimePaths) -> dict[str, str]:
    settings_path = _settings_path(paths)
    defaults = _default_settings(paths)
    if not settings_path.exists():
        return defaults
    try:
        payload = json.loads(settings_path.read_text(encoding="utf-8"))
    except Exception:
        return defaults
    if not isinstance(payload, dict):
        return defaults
    return {key: str(payload.get(key, defaults[key])) for key in defaults}


def save_launcher_settings(paths: RuntimePaths, settings: dict[str, str]) -> Path:
    settings_path = _settings_path(paths)
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    payload = dict(settings)
    payload["saved_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return settings_path


def run_optimizer_from_launcher(
    paths: RuntimePaths | None = None,
    *,
    runtime_config: Config | None = None,
    run_executor=None,
    cli_runner=None,
) -> LauncherRunResult:
    runtime_paths = initialize_user_workspace(paths or resolve_runtime_paths()).paths
    log_path = runtime_paths.logs_dir / f"optimizer_run_{datetime.now():%Y%m%d_%H%M%S}.log"
    executor = run_executor or run_with_config
    legacy_runner = cli_runner

    with open(log_path, "w", encoding="utf-8") as handle, redirect_stdout(handle), redirect_stderr(handle):
        print(f"[Launcher] Run started at {datetime.now():%Y-%m-%d %H:%M:%S}")
        print(f"[Launcher] Workspace: {runtime_paths.user_workspace_dir}")
        previous_log_override = os.environ.get(RUN_LOG_PATH_ENV_VAR)
        os.environ[RUN_LOG_PATH_ENV_VAR] = str(log_path)
        try:
            if runtime_config is not None:
                print("[Launcher] Running with launcher settings (no control workbook).")
                executor(runtime_config, runtime_paths=runtime_paths, input_template=None)
            elif legacy_runner is not None:
                print("[Launcher] Running legacy workbook mode.")
                legacy_runner(args=["--input-template", str(runtime_paths.control_workbook_path)], standalone_mode=False)
            else:
                raise RuntimeError("No run configuration was provided.")
        except SystemExit as exc:
            code = int(exc.code) if isinstance(exc.code, int) else 1
            if code == 0:
                print("[Launcher] Run completed successfully.")
                return LauncherRunResult(True, log_path, "Optimization completed successfully.")
            print(f"[Launcher] Run failed with exit code {code}.")
            return LauncherRunResult(False, log_path, "The optimizer stopped before finishing. Please review the log file.")
        except Exception:
            print("[Launcher] Unexpected exception during run:")
            traceback.print_exc()
            return LauncherRunResult(False, log_path, "The optimizer hit an unexpected error. Please review the log file.")
        finally:
            if previous_log_override is None:
                os.environ.pop(RUN_LOG_PATH_ENV_VAR, None)
            else:
                os.environ[RUN_LOG_PATH_ENV_VAR] = previous_log_override

    return LauncherRunResult(True, log_path, "Optimization completed successfully.")


if PYSIDE6_AVAILABLE:

    class _RunSignals(QObject):
        finished = Signal(object, object)

    class CardFrame(QFrame):
        def __init__(self, title: str, parent: QWidget | None = None) -> None:
            super().__init__(parent)
            self.setObjectName("Card")
            layout = QVBoxLayout(self)
            layout.setContentsMargins(14, 12, 14, 12)
            layout.setSpacing(10)
            self.header = QLabel(title)
            self.header.setObjectName("SectionTitle")
            layout.addWidget(self.header)
            self.body = QWidget(self)
            self.body_layout = QVBoxLayout(self.body)
            self.body_layout.setContentsMargins(0, 0, 0, 0)
            self.body_layout.setSpacing(8)
            layout.addWidget(self.body)

        def set_title(self, title: str) -> None:
            self.header.setText(title)

    class LauncherMainWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle(f"{APP_TITLE} Launcher")
            self.setMinimumSize(1180, 760)
            self.resize(1360, 860)
            self.install_paths = resolve_runtime_paths()
            initial_settings = load_launcher_settings(self.install_paths)
            self.current_language = normalize_language(initial_settings.get("language", "en"))
            self.paths = self._paths_from_workspace_setting(initial_settings.get("workspace_root", ""))
            self.last_log_path: Path | None = None
            self.run_in_progress = False
            self.signals = _RunSignals()
            self._card_refs: dict[str, CardFrame] = {}
            self._label_refs: dict[str, QLabel] = {}
            self._button_refs: dict[str, QPushButton] = {}
            self.signals.finished.connect(self._finish_run)
            self._build_ui()
            self._wire_events()
            self._apply_settings_to_form(initial_settings)
            self._refresh_runtime_displays()
            self._apply_language(self.current_language)
            self._append_status(ui_text(self.current_language, "status_ready"))

        def _build_ui(self) -> None:
            self.window_scroll = QScrollArea(self)
            self.window_scroll.setWidgetResizable(True)
            self.window_scroll.setFrameShape(QFrame.NoFrame)
            self.window_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.window_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.setCentralWidget(self.window_scroll)

            root = QWidget(self.window_scroll)
            self.window_scroll.setWidget(root)
            root_layout = QVBoxLayout(root)
            root_layout.setContentsMargins(14, 14, 14, 14)
            root_layout.setSpacing(10)
            root_layout.setSizeConstraint(QLayout.SetMinimumSize)

            root_layout.addWidget(self._build_header())
            self.splitter = QSplitter(Qt.Horizontal)
            self.splitter.setChildrenCollapsible(False)
            self.splitter.setHandleWidth(8)
            self.sidebar = self._build_sidebar()
            self.main_stack = self._build_main_area()
            self.splitter.addWidget(self.sidebar)
            self.splitter.addWidget(self.main_stack)
            self.splitter.setSizes([380, 980])
            root_layout.addWidget(self.splitter, 1)
            root_layout.addWidget(self._build_footer())
            self._apply_theme("System")

        def _build_header(self) -> QWidget:
            card = CardFrame("Header")
            self._card_refs["header"] = card
            card.layout().setContentsMargins(18, 14, 18, 14)
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(10)

            left = QVBoxLayout()
            self.header_title = QLabel(APP_TITLE)
            self.header_title.setObjectName("Title")
            self.header_subtitle = QLabel()
            self.header_subtitle.setObjectName("Subtitle")
            self.header_subtitle.setWordWrap(True)
            left.addWidget(self.header_title)
            left.addWidget(self.header_subtitle)
            left.addStretch(1)

            right = QHBoxLayout()
            right.setContentsMargins(0, 0, 0, 0)
            right.setSpacing(10)
            right.setAlignment(Qt.AlignTop | Qt.AlignRight)

            self.theme_label = QLabel("Theme")
            self.theme_combo = QComboBox()
            self.theme_combo.addItems(["System", "Light", "Dark"])
            self.theme_combo.setMinimumWidth(132)

            self.language_label = QLabel("Language")
            self.language_combo = QComboBox()
            self.language_combo.addItems(["English", "Chinese"])
            self.language_combo.setMinimumWidth(132)

            self.help_button = QPushButton("Help")
            self._button_refs["help"] = self.help_button

            theme_group = QHBoxLayout()
            theme_group.setContentsMargins(0, 0, 0, 0)
            theme_group.setSpacing(6)
            theme_group.addWidget(self.theme_label, 0, Qt.AlignVCenter)
            theme_group.addWidget(self.theme_combo, 0, Qt.AlignVCenter)

            language_group = QHBoxLayout()
            language_group.setContentsMargins(0, 0, 0, 0)
            language_group.setSpacing(6)
            language_group.addWidget(self.language_label, 0, Qt.AlignVCenter)
            language_group.addWidget(self.language_combo, 0, Qt.AlignVCenter)

            right.addLayout(theme_group)
            right.addLayout(language_group)
            right.addWidget(self.help_button, 0, Qt.AlignVCenter)

            row.addLayout(left, 1)
            row.addLayout(right)
            card.body_layout.addLayout(row)
            return card

        def _build_sidebar(self) -> QWidget:
            card = CardFrame("Navigation & Settings")
            self._card_refs["nav_settings"] = card
            card.setMinimumWidth(360)
            card.setMaximumWidth(380)
            body = card.body_layout
            self.nav_list = QListWidget()
            self.nav_list.setSelectionMode(QAbstractItemView.SingleSelection)
            self.nav_list.setSpacing(4)
            self.nav_list.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.nav_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            for label in ("Home", "Configuration", "License & Diagnostics"):
                QListWidgetItem(label, self.nav_list)
            self.nav_list.setCurrentRow(0)
            row_height = max(30, self.nav_list.sizeHintForRow(0))
            fixed_height = row_height * self.nav_list.count() + 14
            self.nav_list.setFixedHeight(fixed_height)
            body.addWidget(self.nav_list)
            body.addSpacing(10)

            self.sidebar_settings_scroll = QScrollArea()
            self.sidebar_settings_scroll.setWidgetResizable(True)
            self.sidebar_settings_scroll.setFrameShape(QFrame.NoFrame)
            self.sidebar_settings_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            settings_container = QWidget()
            settings_layout = QVBoxLayout(settings_container)
            settings_layout.setContentsMargins(0, 0, 0, 0)
            settings_layout.setSpacing(8)

            self.sidebar_planning_card = CardFrame("Planning")
            self.sidebar_runtime_card = CardFrame("Runtime")
            self._card_refs["planning"] = self.sidebar_planning_card
            self._card_refs["runtime"] = self.sidebar_runtime_card
            self._build_planning_fields()
            self._build_runtime_fields()
            settings_layout.addWidget(self.sidebar_planning_card)
            settings_layout.addWidget(self.sidebar_runtime_card)
            settings_layout.addStretch(1)
            self.sidebar_settings_scroll.setWidget(settings_container)
            body.addWidget(self.sidebar_settings_scroll, 1)

            self.sidebar_status = QLabel("Ready")
            self.sidebar_status.setObjectName("StatusBadge")
            body.addWidget(self.sidebar_status)
            return card

        def _build_main_area(self) -> QStackedWidget:
            stack = QStackedWidget()
            stack.addWidget(self._make_scroll_page(self._build_home_page()))
            stack.addWidget(self._make_scroll_page(self._build_configuration_page()))
            stack.addWidget(self._make_scroll_page(self._build_license_page()))
            return stack

        def _build_footer(self) -> QWidget:
            card = CardFrame("Footer")
            self._card_refs["footer"] = card
            row = QHBoxLayout()
            self.footer_version = QLabel(f"Version: {APP_VERSION}")
            self.footer_license = QLabel("License: validated at runtime")
            self.footer_workspace = QLabel("Workspace: (initializing)")
            row.addWidget(self.footer_version)
            row.addWidget(self.footer_license)
            row.addStretch(1)
            row.addWidget(self.footer_workspace)
            card.body_layout.addLayout(row)
            return card

        def _build_home_page(self) -> QWidget:
            page = QWidget()
            layout = QVBoxLayout(page)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(10)

            actions_card = CardFrame("Actions")
            self._card_refs["actions"] = actions_card
            actions_grid = QGridLayout()
            actions_grid.setHorizontalSpacing(12)
            actions_grid.setVerticalSpacing(10)
            self.run_button = QPushButton("Run Optimization")
            self.run_button.setObjectName("Primary")
            self.open_output_button = QPushButton("Open Output Folder")
            self.open_logs_button = QPushButton("Open Log Folder")
            self.open_last_log_button = QPushButton("Open Latest Log")
            self.open_workspace_button = QPushButton("Open Workspace Folder")
            self._button_refs["run"] = self.run_button
            self._button_refs["open_output"] = self.open_output_button
            self._button_refs["open_logs"] = self.open_logs_button
            self._button_refs["open_last_log"] = self.open_last_log_button
            self._button_refs["open_workspace"] = self.open_workspace_button
            actions_grid.addWidget(self.run_button, 0, 0, 1, 2)
            actions_grid.addWidget(self.open_output_button, 1, 0)
            actions_grid.addWidget(self.open_logs_button, 1, 1)
            actions_grid.addWidget(self.open_last_log_button, 2, 0)
            actions_grid.addWidget(self.open_workspace_button, 2, 1)
            actions_card.body_layout.addLayout(actions_grid)
            layout.addWidget(actions_card)

            self.status_card = CardFrame("Run Status")
            self._card_refs["status"] = self.status_card
            self._build_status_fields()
            layout.addWidget(self.status_card)

            layout.addStretch(1)
            return page

        def _build_configuration_page(self) -> QWidget:
            container = QWidget()
            layout = QVBoxLayout(container)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(10)

            action_card = CardFrame("Configuration Actions")
            self._card_refs["config_actions"] = action_card
            action_row = QHBoxLayout()
            self.initialize_button = QPushButton("Initialize Workspace")
            self.save_settings_button = QPushButton("Save Settings")
            self._button_refs["initialize"] = self.initialize_button
            self._button_refs["save_settings"] = self.save_settings_button
            action_row.addWidget(self.initialize_button)
            action_row.addWidget(self.save_settings_button)
            action_row.addStretch(1)
            action_card.body_layout.addLayout(action_row)

            self.config_workspace_card = CardFrame("Workspace")
            self._card_refs["workspace"] = self.config_workspace_card
            layout.addWidget(action_card)
            layout.addWidget(self.config_workspace_card)
            layout.addStretch(1)

            self._build_workspace_fields()
            return container

        def _build_license_page(self) -> QWidget:
            page = QWidget()
            layout = QVBoxLayout(page)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(10)

            license_card = CardFrame("License & Diagnostics")
            self._card_refs["license"] = license_card
            grid = QGridLayout()
            grid.setHorizontalSpacing(12)
            grid.setVerticalSpacing(10)
            self.generate_fingerprint_button = QPushButton("Generate Machine Fingerprint")
            self.open_requests_button = QPushButton("Open License Requests")
            self.open_license_folder_button = QPushButton("Open License Folder")
            self.open_docs_button = QPushButton("Open Workspace Docs")
            self._button_refs["generate_fingerprint"] = self.generate_fingerprint_button
            self._button_refs["open_requests"] = self.open_requests_button
            self._button_refs["open_license_folder"] = self.open_license_folder_button
            self._button_refs["open_docs"] = self.open_docs_button
            grid.addWidget(self.generate_fingerprint_button, 0, 0)
            grid.addWidget(self.open_requests_button, 0, 1)
            grid.addWidget(self.open_license_folder_button, 1, 0)
            grid.addWidget(self.open_docs_button, 1, 1)
            license_card.body_layout.addLayout(grid)
            layout.addWidget(license_card)
            layout.addStretch(1)
            return page

        def _make_scroll_page(self, content: QWidget) -> QScrollArea:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setWidget(content)
            return scroll

        def _build_workspace_fields(self) -> None:
            form = QFormLayout()
            form.setLabelAlignment(Qt.AlignRight)
            form.setHorizontalSpacing(10)
            form.setVerticalSpacing(8)
            workspace_row = QHBoxLayout()
            self.workspace_root_edit = QLineEdit()
            self.workspace_root_browse = QPushButton("Browse...")
            self.workspace_root_browse.setFixedWidth(106)
            workspace_row.addWidget(self.workspace_root_edit, 1)
            workspace_row.addWidget(self.workspace_root_browse)
            self.config_input_display = self._readonly_line()
            self.config_output_display = self._readonly_line()
            self.workspace_label = QLabel("Workspace")
            self.data_input_label = QLabel("Data_Input")
            self.output_label = QLabel("Output")
            form.addRow(self.workspace_label, workspace_row)
            form.addRow(self.data_input_label, self.config_input_display)
            form.addRow(self.output_label, self.config_output_display)
            self.config_workspace_card.body_layout.addLayout(form)

        def _build_planning_fields(self) -> None:
            form = QFormLayout()
            form.setLabelAlignment(Qt.AlignRight)
            form.setHorizontalSpacing(10)
            form.setVerticalSpacing(8)
            self.scenario_combo = QComboBox()
            self.scenario_combo.setEditable(False)
            self.output_name_edit = QLineEdit()
            self.start_year_edit = QLineEdit()
            self.start_month_combo = QComboBox()
            self.start_month_combo.addItems([str(i) for i in range(1, 13)])
            self.horizon_edit = QLineEdit()
            self.run_mode_combo = QComboBox()
            self.run_mode_combo.addItems(["ModeA", "ModeB", "Both"])
            self.scenario_name_label = QLabel("Scenario Name")
            self.output_file_name_label = QLabel("Output File Name")
            self.start_year_label = QLabel("Start Year")
            self.start_month_label = QLabel("Start Month")
            self.horizon_months_label = QLabel("Horizon Months")
            self.run_mode_label = QLabel("Run Mode")
            form.addRow(self.scenario_name_label, self.scenario_combo)
            form.addRow(self.output_file_name_label, self.output_name_edit)
            form.addRow(self.start_year_label, self.start_year_edit)
            form.addRow(self.start_month_label, self.start_month_combo)
            form.addRow(self.horizon_months_label, self.horizon_edit)
            form.addRow(self.run_mode_label, self.run_mode_combo)
            self.sidebar_planning_card.body_layout.addLayout(form)

        def _build_runtime_fields(self) -> None:
            form = QFormLayout()
            form.setLabelAlignment(Qt.AlignRight)
            form.setHorizontalSpacing(10)
            form.setVerticalSpacing(8)
            self.verbose_combo = QComboBox()
            self.verbose_combo.addItems(["No", "Yes"])
            self.skip_validation_combo = QComboBox()
            self.skip_validation_combo.addItems(["No", "Yes"])
            self.verbose_label = QLabel("Verbose")
            self.skip_validation_label = QLabel("Skip Validation Errors")
            form.addRow(self.verbose_label, self.verbose_combo)
            form.addRow(self.skip_validation_label, self.skip_validation_combo)
            self.sidebar_runtime_card.body_layout.addLayout(form)

        def _build_status_fields(self) -> None:
            self.status_text = QTextEdit()
            self.status_text.setReadOnly(True)
            self.status_text.setMinimumHeight(150)
            self.status_text.setPlaceholderText("")
            self.status_card.body_layout.addWidget(self.status_text)

        def _readonly_line(self) -> QLineEdit:
            line = QLineEdit()
            line.setReadOnly(True)
            return line

        def _path_row(self, grid: QGridLayout, row: int, label: str) -> tuple[QLineEdit, QPushButton]:
            field = QLineEdit()
            browse = QPushButton("Browse...")
            browse.setFixedWidth(106)
            grid.addWidget(QLabel(label), row, 0)
            grid.addWidget(field, row, 1)
            grid.addWidget(browse, row, 2)
            return field, browse

        def _grid_pair(
            self,
            grid: QGridLayout,
            row: int,
            left_label: str,
            left_widget: QWidget,
            right_label: str,
            right_widget: QWidget,
        ) -> None:
            grid.addWidget(QLabel(left_label), row, 0)
            grid.addWidget(left_widget, row, 1)
            grid.addWidget(QLabel(right_label), row, 2)
            grid.addWidget(right_widget, row, 3)

        def _wire_events(self) -> None:
            self.theme_combo.currentTextChanged.connect(self._apply_theme)
            self.language_combo.currentTextChanged.connect(self._language_changed)
            self.help_button.clicked.connect(self._show_help)
            self.nav_list.currentRowChanged.connect(self.main_stack.setCurrentIndex)
            self.run_button.clicked.connect(self._run_clicked)
            self.open_output_button.clicked.connect(lambda: self._open_folder(self.paths.outputs_dir))
            self.open_logs_button.clicked.connect(lambda: self._open_folder(self.paths.logs_dir))
            self.save_settings_button.clicked.connect(self._save_settings_clicked)
            self.initialize_button.clicked.connect(lambda: self._initialize_workspace(show_message=True))
            self.generate_fingerprint_button.clicked.connect(self._generate_fingerprint_clicked)
            self.open_workspace_button.clicked.connect(lambda: self._open_folder(self.paths.user_workspace_dir))
            self.open_requests_button.clicked.connect(lambda: self._open_folder(self.paths.license_requests_dir))
            self.open_last_log_button.clicked.connect(self._open_latest_log_clicked)
            self.open_license_folder_button.clicked.connect(lambda: self._open_folder(self.paths.license_dir))
            self.open_docs_button.clicked.connect(lambda: self._open_folder(self.paths.workspace_docs_dir))
            self.workspace_root_browse.clicked.connect(lambda: self._browse_directory(self.workspace_root_edit))
            self.workspace_root_edit.editingFinished.connect(self._on_workspace_root_changed)

        def _initialize_workspace(self, *, show_message: bool) -> None:
            result = initialize_user_workspace(self.paths)
            self.paths = result.paths
            settings = load_launcher_settings(self.install_paths)
            self._apply_settings_to_form(settings)
            self._refresh_runtime_displays()
            detail = f"sample_data_copied={result.sample_data_copied}"
            self._append_status(f"Workspace ready ({detail}).")
            self._set_sidebar_status(ui_text(self.current_language, "status_workspace_ready"))
            if show_message:
                QMessageBox.information(
                    self,
                    ui_text(self.current_language, "workspace_ready_title"),
                    ui_text(self.current_language, "workspace_ready_message"),
                )

        def _refresh_runtime_displays(self) -> None:
            self.config_input_display.setText(str(self.paths.workspace_input_dir))
            self.config_output_display.setText(str(self.paths.outputs_dir))
            self.footer_workspace.setText(
                ui_text(self.current_language, "workspace_footer", path=self.paths.user_workspace_dir)
            )

        def _apply_settings_to_form(self, settings: dict[str, str]) -> None:
            workspace_root = settings["workspace_root"]
            self.workspace_root_edit.setText(workspace_root)
            self.paths = self._paths_from_workspace_setting(workspace_root)
            self.output_name_edit.setText(settings["output_file_name"])
            self._refresh_scenario_options(
                preferred=settings["scenario_name"],
                load_folder=str(self.paths.workspace_input_dir),
            )
            self.start_year_edit.setText(settings["start_year"])
            self.start_month_combo.setCurrentText(settings["start_month"])
            self.horizon_edit.setText(settings["horizon_months"])
            self._set_combo_value(self.run_mode_combo, settings["run_mode"])
            self._set_combo_value(self.verbose_combo, "Yes" if _to_bool(settings["verbose"]) else "No")
            self._set_combo_value(
                self.skip_validation_combo,
                "Yes" if _to_bool(settings["skip_validation_errors"]) else "No",
            )
            self._set_combo_value(self.theme_combo, settings.get("theme", "System"))
            self._set_combo_value(self.language_combo, normalize_language(settings.get("language", "en")))

        def _collect_settings(self) -> dict[str, str]:
            return {
                "workspace_root": self.workspace_root_edit.text().strip(),
                "output_file_name": self.output_name_edit.text().strip(),
                "scenario_name": self.scenario_combo.currentText().strip(),
                "start_year": self.start_year_edit.text().strip(),
                "start_month": self.start_month_combo.currentText().strip(),
                "horizon_months": self.horizon_edit.text().strip(),
                "run_mode": self._combo_value(self.run_mode_combo, "ModeB"),
                "verbose": self._combo_value(self.verbose_combo, "No"),
                "skip_validation_errors": self._combo_value(self.skip_validation_combo, "No"),
                "theme": self._combo_value(self.theme_combo, "System"),
                "language": self.current_language,
            }

        def _refresh_scenario_options(
            self,
            *,
            preferred: str | None = None,
            load_folder: str | None = None,
        ) -> None:
            folder = (load_folder or str(self.paths.workspace_input_dir)).strip()
            scenarios: list[str] = []
            if folder and os.path.isdir(folder):
                try:
                    scenarios = discover_planner_scenarios(folder)
                except Exception:
                    scenarios = []
            if not scenarios:
                scenarios = ["Baseline"]

            current_value = self.scenario_combo.currentText().strip()
            target = (preferred or current_value or "").strip()
            if target and target not in scenarios:
                scenarios.append(target)

            self.scenario_combo.blockSignals(True)
            self.scenario_combo.clear()
            for value in scenarios:
                self.scenario_combo.addItem(value)
            if target:
                idx = self.scenario_combo.findText(target, Qt.MatchFixedString)
                if idx >= 0:
                    self.scenario_combo.setCurrentIndex(idx)
                else:
                    self.scenario_combo.setCurrentIndex(0)
            else:
                self.scenario_combo.setCurrentIndex(0)
            self.scenario_combo.blockSignals(False)

        def _save_settings_clicked(self) -> None:
            try:
                path = self._save_settings()
            except Exception as exc:
                self._show_error("Save Failed", str(exc))
                return
            self._append_status(f"Settings saved: {path}")
            QMessageBox.information(
                self,
                ui_text(self.current_language, "saved_title"),
                ui_text(self.current_language, "saved_message", path=path),
            )

        def _save_settings(self) -> Path:
            settings = self._collect_settings()
            workspace_root = settings["workspace_root"].strip()
            if not workspace_root:
                raise ValueError("Workspace cannot be empty.")
            self.paths = self._paths_from_workspace_setting(workspace_root)
            self._refresh_runtime_displays()
            self._refresh_scenario_options(
                preferred=settings["scenario_name"],
                load_folder=str(self.paths.workspace_input_dir),
            )
            return save_launcher_settings(self.paths, settings)

        def _build_config(self) -> Config:
            settings = self._collect_settings()
            workspace_root = settings["workspace_root"].strip()
            if not workspace_root:
                raise ValueError("Workspace cannot be empty.")
            self.paths = self._paths_from_workspace_setting(workspace_root)
            self._refresh_runtime_displays()
            year = int(settings["start_year"])
            month = int(settings["start_month"])
            horizon = int(settings["horizon_months"])
            if year < 1900 or year > 9999:
                raise ValueError("Start Year must be between 1900 and 9999.")
            if month < 1 or month > 12:
                raise ValueError("Start Month must be between 1 and 12.")
            if horizon <= 0:
                raise ValueError("Horizon Months must be greater than 0.")
            output_name = settings["output_file_name"].strip()
            if not output_name:
                raise ValueError("Output File Name cannot be empty.")
            if not output_name.lower().endswith(".xlsx"):
                output_name = f"{output_name}.xlsx"
            run_mode = settings["run_mode"]
            if run_mode not in {"ModeA", "ModeB", "Both"}:
                raise ValueError("Run Mode must be ModeA, ModeB, or Both.")

            return Config(
                project_root_folder=str(self.paths.user_workspace_dir),
                input_load_folder=str(self.paths.workspace_input_dir),
                input_master_folder=str(self.paths.workspace_input_dir),
                output_folder=str(self.paths.outputs_dir),
                output_file_name=output_name,
                scenario_name=settings["scenario_name"],
                start_month=f"{year:04d}-{month:02d}",
                horizon_months=horizon,
                run_mode=run_mode,
                verbose=_to_bool(settings["verbose"]),
                skip_validation_errors=_to_bool(settings["skip_validation_errors"]),
                language=self.current_language,
                notes="",
            )

        def _run_clicked(self) -> None:
            if self.run_in_progress:
                return
            try:
                self._save_settings()
                config = self._build_config()
            except Exception as exc:
                self._show_error(ui_text(self.current_language, "invalid_settings_title"), str(exc))
                return

            self.run_in_progress = True
            self.run_button.setEnabled(False)
            self._set_sidebar_status(ui_text(self.current_language, "status_running"))
            self._append_status(ui_text(self.current_language, "status_run_started"))
            thread = threading.Thread(target=self._run_worker, args=(config,), daemon=True)
            thread.start()

        def _run_worker(self, config: Config) -> None:
            try:
                result = run_optimizer_from_launcher(self.paths, runtime_config=config)
                self.signals.finished.emit(result, None)
            except Exception as exc:
                self.signals.finished.emit(None, exc)

        def _finish_run(self, result: LauncherRunResult | None, error: Exception | None) -> None:
            self.run_in_progress = False
            self.run_button.setEnabled(True)
            if error is not None:
                self._set_sidebar_status("Run Failed")
                self._append_status(f"{ui_text(self.current_language, 'status_run_failed')}: {error}")
                self._show_error(ui_text(self.current_language, "run_failed_title"), str(error))
                return
            if result is None:
                self._set_sidebar_status(ui_text(self.current_language, "status_run_failed"))
                self._show_error(ui_text(self.current_language, "run_failed_title"), "No result was returned.")
                return

            self.last_log_path = result.log_path
            self._append_status(result.message)
            self._append_status(f"Log: {result.log_path}")
            if result.success:
                self._set_sidebar_status(ui_text(self.current_language, "status_run_succeeded"))
                QMessageBox.information(
                    self,
                    ui_text(self.current_language, "run_completed_title"),
                    ui_text(self.current_language, "run_completed_message", message=result.message, path=result.log_path),
                )
            else:
                self._set_sidebar_status(ui_text(self.current_language, "status_run_failed"))
                QMessageBox.critical(
                    self,
                    ui_text(self.current_language, "run_failed_title"),
                    ui_text(self.current_language, "run_failed_message", message=result.message, path=result.log_path),
                )

        def _generate_fingerprint_clicked(self) -> None:
            try:
                request_path = generate_machine_fingerprint_request(self.paths)
            except Exception as exc:
                self._show_error(ui_text(self.current_language, "generate_fingerprint_failed_title"), str(exc))
                return
            self._append_status(f"Machine fingerprint generated: {request_path}")
            QMessageBox.information(
                self,
                ui_text(self.current_language, "machine_fingerprint_title"),
                ui_text(self.current_language, "machine_fingerprint_message", path=request_path),
            )

        def _open_latest_log_clicked(self) -> None:
            latest = self.last_log_path
            if latest is None or not latest.exists():
                logs = sorted(self.paths.logs_dir.glob("*.log"), key=lambda p: p.stat().st_mtime, reverse=True)
                latest = logs[0] if logs else None
            if latest is None:
                self._show_error(
                    ui_text(self.current_language, "no_logs_title"),
                    ui_text(self.current_language, "no_logs_message", path=self.paths.logs_dir),
                )
                return
            self._open_path(latest)

        def _browse_directory(self, target: QLineEdit) -> None:
            base = target.text().strip() or str(self.paths.user_workspace_dir)
            selected = QFileDialog.getExistingDirectory(self, "Select Folder", base)
            if selected:
                target.setText(selected)
                if target is self.workspace_root_edit:
                    self._on_workspace_root_changed()

        def _on_workspace_root_changed(self) -> None:
            workspace_root = self.workspace_root_edit.text().strip()
            if not workspace_root:
                return
            self.paths = self._paths_from_workspace_setting(workspace_root)
            self._refresh_runtime_displays()
            self._refresh_scenario_options(load_folder=str(self.paths.workspace_input_dir))

        def _open_folder(self, path: Path) -> None:
            resolved = Path(path)
            resolved.mkdir(parents=True, exist_ok=True)
            self._open_path(resolved)

        def _open_path(self, path: Path) -> None:
            if not path.exists():
                self._show_error(ui_text(self.current_language, "path_not_found_title"), ui_text(self.current_language, "path_not_found_message", path=path))
                return
            opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
            if not opened:
                self._show_error(ui_text(self.current_language, "open_failed_title"), ui_text(self.current_language, "open_failed_message", path=path))

        def _set_sidebar_status(self, text: str) -> None:
            self.sidebar_status.setText(text)

        def _append_status(self, text: str) -> None:
            timestamp = datetime.now().strftime("%H:%M:%S")
            if hasattr(self, "status_text"):
                self.status_text.append(f"[{timestamp}] {text}")

        def _show_help(self) -> None:
            docs_dir = self.paths.workspace_docs_dir
            QMessageBox.information(
                self,
                ui_text(self.current_language, "help_title"),
                ui_text(self.current_language, "help_message", docs_dir=docs_dir),
            )

        def _paths_from_workspace_setting(self, workspace_root: str) -> RuntimePaths:
            target = workspace_root.strip() or str(self.install_paths.user_workspace_dir)
            return with_workspace_dir(self.install_paths, target)

        def _show_error(self, title: str, message: str) -> None:
            self._append_status(f"{title}: {message}")
            QMessageBox.critical(self, title, message)

        def _apply_theme(self, theme_name: str) -> None:
            requested = self._combo_value(self.theme_combo, theme_name or "System").strip().lower()
            if requested == "dark":
                stylesheet = DARK_QSS
            elif requested == "light":
                stylesheet = LIGHT_QSS
            else:
                app = QApplication.instance()
                if app is not None:
                    window_color = app.palette().color(QPalette.Window)
                    stylesheet = DARK_QSS if window_color.lightness() < 128 else LIGHT_QSS
                else:
                    stylesheet = LIGHT_QSS
            self.setStyleSheet(stylesheet)

        def _language_changed(self, text: str) -> None:
            self.current_language = "zh" if text in {"Chinese", "中文"} else "en"
            self._apply_language(self.current_language)

        def _apply_language(self, language: str) -> None:
            self.current_language = normalize_language(language)
            self.setWindowTitle(ui_text(self.current_language, "launcher_title"))
            self.header_title.setText(ui_text(self.current_language, "app_title"))
            self.header_subtitle.setText(ui_text(self.current_language, "header_subtitle"))
            self.theme_label.setText(ui_text(self.current_language, "theme_label"))
            self.language_label.setText(ui_text(self.current_language, "language_label"))
            self.status_text.setPlaceholderText(ui_text(self.current_language, "status_placeholder"))
            self.sidebar_status.setText(ui_text(self.current_language, "status_ready"))
            self.footer_version.setText(ui_text(self.current_language, "version_label", version=APP_VERSION))
            self.footer_license.setText(ui_text(self.current_language, "license_label"))

            self._card_refs["header"].set_title(ui_text(self.current_language, "header_card"))
            self._card_refs["nav_settings"].set_title(ui_text(self.current_language, "nav_settings_card"))
            self._card_refs["planning"].set_title(ui_text(self.current_language, "planning_card"))
            self._card_refs["runtime"].set_title(ui_text(self.current_language, "runtime_card"))
            self._card_refs["actions"].set_title(ui_text(self.current_language, "actions_card"))
            self._card_refs["status"].set_title(ui_text(self.current_language, "status_card"))
            self._card_refs["config_actions"].set_title(ui_text(self.current_language, "config_actions_card"))
            self._card_refs["workspace"].set_title(ui_text(self.current_language, "workspace_card"))
            self._card_refs["license"].set_title(ui_text(self.current_language, "license_card"))
            self._card_refs["footer"].set_title(ui_text(self.current_language, "footer_card"))

            self.nav_list.item(0).setText(ui_text(self.current_language, "nav_home"))
            self.nav_list.item(1).setText(ui_text(self.current_language, "nav_configuration"))
            self.nav_list.item(2).setText(ui_text(self.current_language, "nav_license"))

            self.workspace_label.setText(ui_text(self.current_language, "workspace_label"))
            self.data_input_label.setText(ui_text(self.current_language, "data_input_label"))
            self.output_label.setText(ui_text(self.current_language, "output_label"))
            self.scenario_name_label.setText(ui_text(self.current_language, "scenario_name_label"))
            self.output_file_name_label.setText(ui_text(self.current_language, "output_file_name_label"))
            self.start_year_label.setText(ui_text(self.current_language, "start_year_label"))
            self.start_month_label.setText(ui_text(self.current_language, "start_month_label"))
            self.horizon_months_label.setText(ui_text(self.current_language, "horizon_months_label"))
            self.run_mode_label.setText(ui_text(self.current_language, "run_mode_label"))
            self.verbose_label.setText(ui_text(self.current_language, "verbose_label"))
            self.skip_validation_label.setText(ui_text(self.current_language, "skip_validation_label"))

            self.help_button.setText(ui_text(self.current_language, "help_button"))
            self.run_button.setText(ui_text(self.current_language, "run_button"))
            self.open_output_button.setText(ui_text(self.current_language, "open_output_button"))
            self.open_logs_button.setText(ui_text(self.current_language, "open_logs_button"))
            self.open_last_log_button.setText(ui_text(self.current_language, "open_last_log_button"))
            self.open_workspace_button.setText(ui_text(self.current_language, "open_workspace_button"))
            self.initialize_button.setText(ui_text(self.current_language, "initialize_button"))
            self.save_settings_button.setText(ui_text(self.current_language, "save_settings_button"))
            self.generate_fingerprint_button.setText(ui_text(self.current_language, "generate_fingerprint_button"))
            self.open_requests_button.setText(ui_text(self.current_language, "open_requests_button"))
            self.open_license_folder_button.setText(ui_text(self.current_language, "open_license_folder_button"))
            self.open_docs_button.setText(ui_text(self.current_language, "open_docs_button"))

            self._set_combo_items(
                self.language_combo,
                [ui_text("en", "language_english"), ui_text(self.current_language, "language_chinese")],
                ui_text(self.current_language, "language_chinese") if self.current_language == "zh" else ui_text("en", "language_english"),
            )
            self._set_combo_items(
                self.theme_combo,
                [ui_text(self.current_language, "theme_system"), ui_text(self.current_language, "theme_light"), ui_text(self.current_language, "theme_dark")],
                self.theme_combo.currentText(),
            )
            self._set_combo_items(
                self.run_mode_combo,
                ["ModeA", "ModeB", ui_text(self.current_language, "mode_both")],
                self.run_mode_combo.currentText(),
            )
            self._set_combo_items(
                self.verbose_combo,
                [ui_text(self.current_language, "no"), ui_text(self.current_language, "yes")],
                ui_text(self.current_language, "yes") if _to_bool(self.verbose_combo.currentText()) else ui_text(self.current_language, "no"),
            )
            self._set_combo_items(
                self.skip_validation_combo,
                [ui_text(self.current_language, "no"), ui_text(self.current_language, "yes")],
                ui_text(self.current_language, "yes") if _to_bool(self.skip_validation_combo.currentText()) else ui_text(self.current_language, "no"),
            )
            self._refresh_runtime_displays()

        def _set_combo_items(self, combo: QComboBox, items: list[str], current_text: str) -> None:
            combo.blockSignals(True)
            selected = current_text
            combo.clear()
            combo.addItems(items)
            idx = combo.findText(selected, Qt.MatchFixedString)
            if idx < 0:
                idx = 0
            combo.setCurrentIndex(idx)
            combo.blockSignals(False)

        # Overrides for stable bilingual internal values.
        def _apply_theme(self, theme_name: str) -> None:
            requested = self._combo_value(self.theme_combo, theme_name or "System").strip().lower()
            if requested == "dark":
                stylesheet = DARK_QSS
            elif requested == "light":
                stylesheet = LIGHT_QSS
            else:
                app = QApplication.instance()
                if app is not None:
                    window_color = app.palette().color(QPalette.Window)
                    stylesheet = DARK_QSS if window_color.lightness() < 128 else LIGHT_QSS
                else:
                    stylesheet = LIGHT_QSS
            self.setStyleSheet(stylesheet)

        def _language_changed(self, text: str) -> None:
            self.current_language = normalize_language(self._combo_value(self.language_combo, text))
            self._apply_language(self.current_language)

        def _apply_language(self, language: str) -> None:
            self.current_language = normalize_language(language)
            self.setWindowTitle(ui_text(self.current_language, "launcher_title"))
            self.header_title.setText(ui_text(self.current_language, "app_title"))
            self.header_subtitle.setText(ui_text(self.current_language, "header_subtitle"))
            self.theme_label.setText(ui_text(self.current_language, "theme_label"))
            self.language_label.setText(ui_text(self.current_language, "language_label"))
            self.status_text.setPlaceholderText(ui_text(self.current_language, "status_placeholder"))
            self.sidebar_status.setText(ui_text(self.current_language, "status_ready"))
            self.footer_version.setText(ui_text(self.current_language, "version_label", version=APP_VERSION))
            self.footer_license.setText(ui_text(self.current_language, "license_label"))

            self._card_refs["header"].set_title(ui_text(self.current_language, "header_card"))
            self._card_refs["nav_settings"].set_title(ui_text(self.current_language, "nav_settings_card"))
            self._card_refs["planning"].set_title(ui_text(self.current_language, "planning_card"))
            self._card_refs["runtime"].set_title(ui_text(self.current_language, "runtime_card"))
            self._card_refs["actions"].set_title(ui_text(self.current_language, "actions_card"))
            self._card_refs["status"].set_title(ui_text(self.current_language, "status_card"))
            self._card_refs["config_actions"].set_title(ui_text(self.current_language, "config_actions_card"))
            self._card_refs["workspace"].set_title(ui_text(self.current_language, "workspace_card"))
            self._card_refs["license"].set_title(ui_text(self.current_language, "license_card"))
            self._card_refs["footer"].set_title(ui_text(self.current_language, "footer_card"))

            self.nav_list.item(0).setText(ui_text(self.current_language, "nav_home"))
            self.nav_list.item(1).setText(ui_text(self.current_language, "nav_configuration"))
            self.nav_list.item(2).setText(ui_text(self.current_language, "nav_license"))

            self.workspace_label.setText(ui_text(self.current_language, "workspace_label"))
            self.data_input_label.setText(ui_text(self.current_language, "data_input_label"))
            self.output_label.setText(ui_text(self.current_language, "output_label"))
            self.scenario_name_label.setText(ui_text(self.current_language, "scenario_name_label"))
            self.output_file_name_label.setText(ui_text(self.current_language, "output_file_name_label"))
            self.start_year_label.setText(ui_text(self.current_language, "start_year_label"))
            self.start_month_label.setText(ui_text(self.current_language, "start_month_label"))
            self.horizon_months_label.setText(ui_text(self.current_language, "horizon_months_label"))
            self.run_mode_label.setText(ui_text(self.current_language, "run_mode_label"))
            self.verbose_label.setText(ui_text(self.current_language, "verbose_label"))
            self.skip_validation_label.setText(ui_text(self.current_language, "skip_validation_label"))

            self.help_button.setText(ui_text(self.current_language, "help_button"))
            self.run_button.setText(ui_text(self.current_language, "run_button"))
            self.open_output_button.setText(ui_text(self.current_language, "open_output_button"))
            self.open_logs_button.setText(ui_text(self.current_language, "open_logs_button"))
            self.open_last_log_button.setText(ui_text(self.current_language, "open_last_log_button"))
            self.open_workspace_button.setText(ui_text(self.current_language, "open_workspace_button"))
            self.initialize_button.setText(ui_text(self.current_language, "initialize_button"))
            self.save_settings_button.setText(ui_text(self.current_language, "save_settings_button"))
            self.generate_fingerprint_button.setText(ui_text(self.current_language, "generate_fingerprint_button"))
            self.open_requests_button.setText(ui_text(self.current_language, "open_requests_button"))
            self.open_license_folder_button.setText(ui_text(self.current_language, "open_license_folder_button"))
            self.open_docs_button.setText(ui_text(self.current_language, "open_docs_button"))

            current_theme = self._combo_value(self.theme_combo, "System")
            current_run_mode = self._combo_value(self.run_mode_combo, "ModeB")
            current_verbose = self._combo_value(self.verbose_combo, "No")
            current_skip = self._combo_value(self.skip_validation_combo, "No")

            self._set_combo_items(
                self.language_combo,
                [
                    (ui_text(self.current_language, "language_english"), "en"),
                    (ui_text(self.current_language, "language_chinese"), "zh"),
                ],
                self.current_language,
            )
            self._set_combo_items(
                self.theme_combo,
                [
                    (ui_text(self.current_language, "theme_system"), "System"),
                    (ui_text(self.current_language, "theme_light"), "Light"),
                    (ui_text(self.current_language, "theme_dark"), "Dark"),
                ],
                current_theme,
            )
            self._set_combo_items(
                self.run_mode_combo,
                [
                    (ui_text(self.current_language, "mode_modea"), "ModeA"),
                    (ui_text(self.current_language, "mode_modeb"), "ModeB"),
                    (ui_text(self.current_language, "mode_both"), "Both"),
                ],
                current_run_mode,
            )
            self._set_combo_items(
                self.verbose_combo,
                [
                    (ui_text(self.current_language, "no"), "No"),
                    (ui_text(self.current_language, "yes"), "Yes"),
                ],
                current_verbose,
            )
            self._set_combo_items(
                self.skip_validation_combo,
                [
                    (ui_text(self.current_language, "no"), "No"),
                    (ui_text(self.current_language, "yes"), "Yes"),
                ],
                current_skip,
            )
            self._refresh_runtime_displays()

        def _set_combo_items(self, combo: QComboBox, items: list[tuple[str, str]], current_value: str) -> None:
            combo.blockSignals(True)
            combo.clear()
            for label, value in items:
                combo.addItem(label, value)
            idx = combo.findData(current_value)
            if idx < 0:
                idx = 0
            combo.setCurrentIndex(idx)
            combo.blockSignals(False)

        def _combo_value(self, combo: QComboBox, fallback: str) -> str:
            value = combo.currentData()
            if value is None or str(value).strip() == "":
                text = combo.currentText().strip()
                return text or fallback
            return str(value)

        def _set_combo_value(self, combo: QComboBox, value: str) -> None:
            idx = combo.findData(value)
            if idx < 0:
                idx = combo.findText(value, Qt.MatchFixedString)
            if idx >= 0:
                combo.setCurrentIndex(idx)


def main() -> int:
    if not PYSIDE6_AVAILABLE:
        print(
            "PySide6 is required for the desktop launcher UI.\n"
            "Install it with: python -m pip install PySide6"
        )
        return 1
    app = QApplication.instance() or QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    app.setApplicationName(APP_TITLE)
    window = LauncherMainWindow()
    window.show()
    return app.exec()
