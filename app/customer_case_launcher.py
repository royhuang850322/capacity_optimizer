"""Standalone desktop launcher for product analysis reports."""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

from app.desktop_launcher import LIGHT_QSS
from app.modeb_customer_case_report import (
    DEFAULT_OUTPUT_NAME,
    ReportValidationError,
    generate_modeb_customer_case_report,
    infer_workspace_root_from_report,
    resolve_mode_report_selection,
)
from app.runtime_paths import RuntimePaths, resolve_runtime_paths, with_workspace_dir
from app.version import APP_VERSION

try:
    from PySide6.QtCore import Qt, QUrl
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFormLayout,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QScrollArea,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )

    PYSIDE6_AVAILABLE = True
except ModuleNotFoundError:
    PYSIDE6_AVAILABLE = False


APP_TITLE = f"Product Analysis Reporter {APP_VERSION}"
SETTINGS_FILENAME = "product_analysis_launcher_settings.json"
LEGACY_SETTINGS_FILENAME = "modeb_product_analysis_launcher_settings.json"


def _show_native_error(title: str, message: str) -> None:
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
    except Exception:
        pass


def is_capacity_optimizer_workspace(path: str | Path | None) -> bool:
    if not path:
        return False
    root = Path(path).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        return False
    data_input_dir = root / "Data_Input"
    if not data_input_dir.exists():
        return False
    return (
        (root / "output").exists()
        or (root / "CapacityOptimizer.exe").exists()
        or (root / "workspace_manifest.json").exists()
    )


def guess_workspace_root(base_paths: RuntimePaths, saved_workspace_root: str | Path | None = None) -> Path | None:
    candidates: list[Path] = []
    if saved_workspace_root:
        candidates.append(Path(saved_workspace_root))
    candidates.append(base_paths.app_install_dir)
    candidates.append(base_paths.app_install_dir.parent)

    seen: set[Path] = set()
    for candidate in candidates:
        resolved = candidate.expanduser().resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        if is_capacity_optimizer_workspace(resolved):
            return resolved
    return None


def _default_settings(_paths: RuntimePaths) -> dict[str, str]:
    return {
        "workspace_root": "",
        "report_mode": "ModeB",
        "use_latest_report": "Yes",
        "manual_report_path": "",
        "output_file_name": DEFAULT_OUTPUT_NAME,
        **{f"product_{index}": "" for index in range(1, 11)},
    }


def _settings_path(paths: RuntimePaths) -> Path:
    return paths.app_install_dir / SETTINGS_FILENAME


def _legacy_settings_path(paths: RuntimePaths) -> Path:
    return paths.app_install_dir / LEGACY_SETTINGS_FILENAME


def load_customer_case_settings(paths: RuntimePaths) -> dict[str, str]:
    defaults = _default_settings(paths)
    settings_path = _settings_path(paths)
    if not settings_path.exists():
        legacy_path = _legacy_settings_path(paths)
        if legacy_path.exists():
            settings_path = legacy_path
        else:
            return defaults
    if not settings_path.exists():
        return defaults
    try:
        payload = json.loads(settings_path.read_text(encoding="utf-8"))
    except Exception:
        return defaults
    if not isinstance(payload, dict):
        return defaults
    return {key: str(payload.get(key, defaults[key])) for key in defaults}


def save_customer_case_settings(paths: RuntimePaths, settings: dict[str, str]) -> Path:
    settings_path = _settings_path(paths)
    payload = {key: str(value) for key, value in settings.items()}
    settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return settings_path


def _to_bool(value: str | bool | None) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    return text in {"1", "true", "yes", "y", "on", "是"}


if PYSIDE6_AVAILABLE:
    class CustomerCaseMainWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.base_paths = resolve_runtime_paths()
            self.settings = load_customer_case_settings(self.base_paths)
            self.setWindowTitle(APP_TITLE)
            self.resize(1080, 860)
            self.setStyleSheet(LIGHT_QSS)
            self._build_ui()
            self._apply_settings()

        def _build_ui(self) -> None:
            central = QWidget()
            outer = QVBoxLayout(central)
            outer.setContentsMargins(18, 18, 18, 18)
            outer.setSpacing(14)

            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll_container = QWidget()
            self.scroll_layout = QVBoxLayout(scroll_container)
            self.scroll_layout.setContentsMargins(0, 0, 0, 0)
            self.scroll_layout.setSpacing(14)

            header = self._card()
            header_layout = QVBoxLayout(header)
            title = QLabel("产品分析工具")
            title.setObjectName("Title")
            subtitle = QLabel("读取已有的 ModeA 或 ModeB 单报告，并复用 CapacityOptimizer 的共享工作目录生成产品分析 Excel。")
            subtitle.setObjectName("Subtitle")
            subtitle.setWordWrap(True)
            header_layout.addWidget(title)
            header_layout.addWidget(subtitle)
            self.scroll_layout.addWidget(header)

            source_card = self._card()
            source_layout = QFormLayout(source_card)
            source_layout.setLabelAlignment(Qt.AlignLeft)
            source_layout.setFormAlignment(Qt.AlignTop)

            workspace_row = QWidget()
            workspace_layout = QHBoxLayout(workspace_row)
            workspace_layout.setContentsMargins(0, 0, 0, 0)
            workspace_layout.setSpacing(8)
            self.workspace_root_edit = QLineEdit()
            self.workspace_root_edit.setPlaceholderText("选择 CapacityOptimizer 的工作目录")
            self.workspace_root_edit.textChanged.connect(self._on_workspace_root_changed)
            self.browse_workspace_button = QPushButton("浏览...")
            self.browse_workspace_button.clicked.connect(self._browse_workspace_root)
            workspace_layout.addWidget(self.workspace_root_edit, 1)
            workspace_layout.addWidget(self.browse_workspace_button)

            self.report_mode_combo = QComboBox()
            self.report_mode_combo.addItems(["ModeA", "ModeB"])
            self.report_mode_combo.currentTextChanged.connect(self._refresh_latest_report_hint)

            self.output_display = self._readonly_line()
            self.data_input_display = self._readonly_line()
            self.use_latest_checkbox = QCheckBox("自动调用共享 output 中最新的单报告")
            self.use_latest_checkbox.stateChanged.connect(self._toggle_manual_report_state)

            manual_path_row = QWidget()
            manual_path_layout = QHBoxLayout(manual_path_row)
            manual_path_layout.setContentsMargins(0, 0, 0, 0)
            manual_path_layout.setSpacing(8)
            self.manual_report_path_edit = QLineEdit()
            self.manual_report_path_edit.setPlaceholderText("可输入完整路径，或只输入共享 output 目录下的文件名")
            self.manual_report_path_edit.editingFinished.connect(self._sync_workspace_from_manual_report)
            self.browse_report_button = QPushButton("浏览...")
            self.browse_report_button.clicked.connect(self._browse_report)
            manual_path_layout.addWidget(self.manual_report_path_edit, 1)
            manual_path_layout.addWidget(self.browse_report_button)

            self.latest_report_hint = QLabel("请先选择 CapacityOptimizer 的工作目录。")
            self.latest_report_hint.setWordWrap(True)
            self.output_name_edit = QLineEdit()
            self.output_name_edit.setPlaceholderText("例如 product_analysis.xlsx（会自动追加时间戳）")

            source_layout.addRow("共享工作目录", workspace_row)
            source_layout.addRow("报告模式", self.report_mode_combo)
            source_layout.addRow("共享 output 目录", self.output_display)
            source_layout.addRow("共享 Data_Input 目录", self.data_input_display)
            source_layout.addRow("", self.use_latest_checkbox)
            source_layout.addRow("手工结果报告", manual_path_row)
            source_layout.addRow("最新报告提示", self.latest_report_hint)
            source_layout.addRow("输出文件名", self.output_name_edit)
            self.scroll_layout.addWidget(source_card)

            product_card = self._card()
            product_layout = QGridLayout(product_card)
            self.product_edits: list[QLineEdit] = []
            for index in range(10):
                label = QLabel(f"产品 {index + 1}")
                edit = QLineEdit()
                edit.setPlaceholderText("输入产品号，留空则忽略")
                self.product_edits.append(edit)
                row = index // 2
                col = (index % 2) * 2
                product_layout.addWidget(label, row, col)
                product_layout.addWidget(edit, row, col + 1)
            self.scroll_layout.addWidget(product_card)

            action_card = self._card()
            action_layout = QVBoxLayout(action_card)
            button_row = QHBoxLayout()
            self.generate_button = QPushButton("生成产品分析报告")
            self.generate_button.setObjectName("Primary")
            self.generate_button.clicked.connect(self._generate_report)
            self.open_output_button = QPushButton("打开共享 output 目录")
            self.open_output_button.clicked.connect(self._open_current_output_dir)
            self.save_settings_button = QPushButton("保存设置")
            self.save_settings_button.clicked.connect(self._save_settings)
            button_row.addWidget(self.generate_button)
            button_row.addWidget(self.open_output_button)
            button_row.addWidget(self.save_settings_button)
            button_row.addStretch(1)
            action_layout.addLayout(button_row)

            self.status_box = QTextEdit()
            self.status_box.setReadOnly(True)
            self.status_box.setPlaceholderText("这里会显示工作目录校验、旧版本提示和生成结果。")
            action_layout.addWidget(self.status_box)
            self.scroll_layout.addWidget(action_card)

            scroll.setWidget(scroll_container)
            outer.addWidget(scroll)
            self.setCentralWidget(central)

        def _card(self) -> QFrame:
            frame = QFrame()
            frame.setObjectName("Card")
            return frame

        def _readonly_line(self) -> QLineEdit:
            line = QLineEdit()
            line.setReadOnly(True)
            return line

        def _apply_settings(self) -> None:
            inferred_root = guess_workspace_root(self.base_paths, self.settings.get("workspace_root", ""))
            self.workspace_root_edit.setText(str(inferred_root) if inferred_root else self.settings.get("workspace_root", ""))
            self.report_mode_combo.setCurrentText(self.settings.get("report_mode", "ModeB"))
            self.use_latest_checkbox.setChecked(_to_bool(self.settings.get("use_latest_report", "Yes")))
            self.manual_report_path_edit.setText(self.settings.get("manual_report_path", ""))
            self.output_name_edit.setText(self.settings.get("output_file_name", DEFAULT_OUTPUT_NAME))
            for index, edit in enumerate(self.product_edits, start=1):
                edit.setText(self.settings.get(f"product_{index}", ""))
            self._toggle_manual_report_state()
            self._refresh_workspace_state()

        def _resolved_workspace_root(self) -> Path | None:
            text = self.workspace_root_edit.text().strip()
            if not text:
                return None
            try:
                return Path(text).expanduser().resolve()
            except Exception:
                return None

        def _workspace_paths(self) -> RuntimePaths | None:
            root = self._resolved_workspace_root()
            if root is None:
                return None
            return with_workspace_dir(self.base_paths, root)

        def _collect_settings(self) -> dict[str, str]:
            payload = {
                "workspace_root": self.workspace_root_edit.text().strip(),
                "report_mode": self.report_mode_combo.currentText(),
                "use_latest_report": "Yes" if self.use_latest_checkbox.isChecked() else "No",
                "manual_report_path": self.manual_report_path_edit.text().strip(),
                "output_file_name": self.output_name_edit.text().strip() or DEFAULT_OUTPUT_NAME,
            }
            for index, edit in enumerate(self.product_edits, start=1):
                payload[f"product_{index}"] = edit.text().strip()
            return payload

        def _save_settings(self) -> None:
            settings_path = save_customer_case_settings(self.base_paths, self._collect_settings())
            self._append_status(f"设置已保存：{settings_path}")

        def _append_status(self, message: str) -> None:
            self.status_box.append(message)

        def _refresh_workspace_state(self) -> None:
            workspace_paths = self._workspace_paths()
            if workspace_paths is None:
                self.output_display.clear()
                self.data_input_display.clear()
                self.latest_report_hint.setText("请先选择 CapacityOptimizer 的工作目录。")
                return

            self.output_display.setText(str(workspace_paths.outputs_dir))
            self.data_input_display.setText(str(workspace_paths.workspace_input_dir))
            self._refresh_latest_report_hint()

        def _refresh_latest_report_hint(self) -> None:
            workspace_paths = self._workspace_paths()
            if workspace_paths is None:
                self.latest_report_hint.setText("请先选择 CapacityOptimizer 的工作目录。")
                return

            workspace_root = workspace_paths.user_workspace_dir
            if not is_capacity_optimizer_workspace(workspace_root):
                self.latest_report_hint.setText("当前目录不像是 CapacityOptimizer 工作目录。请选择包含 Data_Input 和 output 的目录。")
                return

            report_mode = self.report_mode_combo.currentText()
            try:
                selection = resolve_mode_report_selection(
                    output_dir=workspace_paths.outputs_dir,
                    manual_report_path=None,
                    use_latest_report=True,
                    report_mode=report_mode,
                )
            except Exception:
                self.latest_report_hint.setText(f"共享 output 目录下还没有可用的 {report_mode} 报告。")
                return
            stamp = datetime.fromtimestamp(selection.selected_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            self.latest_report_hint.setText(f"最新 {report_mode} 报告：{selection.selected_path.name} | 时间：{stamp}")

        def _toggle_manual_report_state(self) -> None:
            manual_enabled = not self.use_latest_checkbox.isChecked()
            self.manual_report_path_edit.setEnabled(manual_enabled)
            self.browse_report_button.setEnabled(manual_enabled)

        def _on_workspace_root_changed(self) -> None:
            self._refresh_workspace_state()

        def _browse_workspace_root(self) -> None:
            initial_dir = self.workspace_root_edit.text().strip() or str(self.base_paths.app_install_dir)
            path = QFileDialog.getExistingDirectory(self, "选择 CapacityOptimizer 工作目录", initial_dir)
            if path:
                self.workspace_root_edit.setText(path)

        def _browse_report(self) -> None:
            workspace_paths = self._workspace_paths()
            initial_dir = str(workspace_paths.outputs_dir) if workspace_paths is not None else str(self.base_paths.app_install_dir)
            report_mode = self.report_mode_combo.currentText()
            path, _ = QFileDialog.getOpenFileName(
                self,
                f"选择 {report_mode} 报告",
                initial_dir,
                "Excel Workbook (*.xlsx)",
            )
            if path:
                self.manual_report_path_edit.setText(path)
                self._sync_workspace_from_manual_report()

        def _sync_workspace_from_manual_report(self) -> None:
            manual_text = self.manual_report_path_edit.text().strip()
            if not manual_text:
                return
            inferred_root = infer_workspace_root_from_report(manual_text)
            if inferred_root is not None and is_capacity_optimizer_workspace(inferred_root):
                self.workspace_root_edit.setText(str(inferred_root))

        def _open_path(self, path: Path) -> None:
            if not path.exists():
                QMessageBox.warning(self, "路径不存在", f"以下路径不存在：\n{path}")
                return
            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
            if not ok:
                QMessageBox.warning(self, "打开失败", f"无法打开以下路径：\n{path}")

        def _open_current_output_dir(self) -> None:
            workspace_paths = self._workspace_paths()
            if workspace_paths is None:
                QMessageBox.warning(self, "缺少工作目录", "请先选择 CapacityOptimizer 的工作目录。")
                return
            self._open_path(workspace_paths.outputs_dir)

        def _collect_products(self) -> list[str]:
            seen: set[str] = set()
            products: list[str] = []
            for edit in self.product_edits:
                text = edit.text().strip()
                if not text or text in seen:
                    continue
                products.append(text)
                seen.add(text)
            return products

        def _generate_report(self) -> None:
            products = self._collect_products()
            if not products:
                QMessageBox.warning(self, "缺少产品号", "请至少输入 1 个产品号。")
                return
            report_mode = self.report_mode_combo.currentText()

            workspace_paths = self._workspace_paths()
            if workspace_paths is None:
                QMessageBox.warning(self, "缺少工作目录", "请先选择 CapacityOptimizer 的工作目录。")
                return
            if not is_capacity_optimizer_workspace(workspace_paths.user_workspace_dir):
                QMessageBox.warning(self, "工作目录无效", "当前目录不像是 CapacityOptimizer 工作目录。请选择包含 Data_Input 和 output 的目录。")
                return

            try:
                selection = resolve_mode_report_selection(
                    output_dir=workspace_paths.outputs_dir,
                    manual_report_path=self.manual_report_path_edit.text().strip(),
                    use_latest_report=self.use_latest_checkbox.isChecked(),
                    report_mode=report_mode,
                )
            except Exception as exc:
                QMessageBox.warning(self, "报告选择失败", str(exc))
                self._append_status(f"报告选择失败：{exc}")
                return

            if not self.use_latest_checkbox.isChecked() and selection.latest_path and not selection.is_latest:
                answer = QMessageBox.question(
                    self,
                    f"不是最新 {report_mode} 报告",
                    (
                        f"你当前选择的文件不是最新版本。\n\n"
                        f"当前选择：{selection.selected_path.name}\n"
                        f"最新文件：{selection.latest_path.name}\n\n"
                        "是否仍继续分析你选择的旧版本？"
                    ),
                )
                if answer != QMessageBox.Yes:
                    self._append_status("用户取消了旧版本报告分析。")
                    return

            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                output_path = generate_modeb_customer_case_report(
                    report_path=selection.selected_path,
                    products=products,
                    report_mode=report_mode,
                    output_dir=workspace_paths.outputs_dir,
                    output_name=self.output_name_edit.text().strip() or DEFAULT_OUTPUT_NAME,
                    runtime_paths=workspace_paths,
                    latest_report_path=selection.latest_path,
                )
                self._save_settings()
            except (ReportValidationError, FileNotFoundError, ValueError) as exc:
                QMessageBox.warning(self, "生成失败", str(exc))
                self._append_status(f"生成失败：{exc}")
            except Exception as exc:  # pragma: no cover - UI safeguard
                QMessageBox.critical(self, "生成失败", f"出现未预期错误：\n{exc}")
                self._append_status(f"未预期错误：{exc}")
            else:
                self._append_status(f"产品分析报告已生成：{output_path}")
                QMessageBox.information(self, "生成成功", f"产品分析报告已生成：\n{output_path}")
            finally:
                QApplication.restoreOverrideCursor()
                self._refresh_latest_report_hint()


def main() -> int:
    if not PYSIDE6_AVAILABLE:
        message = (
            "PySide6 is required for the product analysis launcher UI.\n"
            "Install it with: python -m pip install PySide6"
        )
        _show_native_error(APP_TITLE, message)
        raise SystemExit(message)
    app = QApplication.instance() or QApplication(sys.argv)
    try:
        window = CustomerCaseMainWindow()
    except Exception as exc:
        message = f"启动工具时发生错误：\n{exc}"
        try:
            QMessageBox.critical(None, APP_TITLE, message)
        except Exception:
            _show_native_error(APP_TITLE, message)
        return 1
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
