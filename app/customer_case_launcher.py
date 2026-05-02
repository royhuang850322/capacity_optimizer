"""Standalone desktop launcher for ModeB product analysis reports."""
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
    resolve_modeb_report_selection,
)
from app.runtime_paths import RuntimePaths, ensure_workspace_dirs, resolve_runtime_paths

try:
    from PySide6.QtCore import Qt, QUrl
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
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


APP_TITLE = "ModeB Product Analysis Reporter"
SETTINGS_FILENAME = "modeb_product_analysis_launcher_settings.json"


def _show_native_error(title: str, message: str) -> None:
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
    except Exception:
        pass


def _default_settings(paths: RuntimePaths) -> dict[str, str]:
    return {
        "workspace_root": str(paths.user_workspace_dir),
        "use_latest_report": "Yes",
        "manual_report_path": "",
        "output_file_name": DEFAULT_OUTPUT_NAME,
        **{f"product_{index}": "" for index in range(1, 11)},
    }


def _settings_path(paths: RuntimePaths) -> Path:
    return paths.app_install_dir / SETTINGS_FILENAME


def load_customer_case_settings(paths: RuntimePaths) -> dict[str, str]:
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
            self.paths = ensure_workspace_dirs(resolve_runtime_paths())
            self.settings = load_customer_case_settings(self.paths)
            self.setWindowTitle(f"{APP_TITLE}")
            self.resize(1080, 840)
            self.setStyleSheet(LIGHT_QSS)
            self._build_ui()
            self._apply_settings()
            self._refresh_latest_report_hint()

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
            title = QLabel("ModeB 产品分析工具")
            title.setObjectName("Title")
            subtitle = QLabel("读取一个 ModeB 输出报告，再为输入的 1 到 10 个产品自动生成分析版 Excel。")
            subtitle.setObjectName("Subtitle")
            subtitle.setWordWrap(True)
            header_layout.addWidget(title)
            header_layout.addWidget(subtitle)
            self.scroll_layout.addWidget(header)

            source_card = self._card()
            source_layout = QFormLayout(source_card)
            source_layout.setLabelAlignment(Qt.AlignLeft)
            source_layout.setFormAlignment(Qt.AlignTop)

            self.workspace_display = self._readonly_line()
            self.output_display = self._readonly_line()
            self.data_input_display = self._readonly_line()
            self.use_latest_checkbox = QCheckBox("自动调用 output 中最新的 ModeB 报告")
            self.use_latest_checkbox.stateChanged.connect(self._toggle_manual_report_state)

            manual_path_row = QWidget()
            manual_path_layout = QHBoxLayout(manual_path_row)
            manual_path_layout.setContentsMargins(0, 0, 0, 0)
            manual_path_layout.setSpacing(8)
            self.manual_report_path_edit = QLineEdit()
            self.manual_report_path_edit.setPlaceholderText("可输入完整路径，或只输入 output 目录下的文件名")
            self.browse_report_button = QPushButton("浏览...")
            self.browse_report_button.clicked.connect(self._browse_report)
            manual_path_layout.addWidget(self.manual_report_path_edit, 1)
            manual_path_layout.addWidget(self.browse_report_button)

            self.latest_report_hint = QLabel("正在检查最新的 ModeB 报告...")
            self.latest_report_hint.setWordWrap(True)
            self.output_name_edit = QLineEdit()
            self.output_name_edit.setPlaceholderText("例如 product_analysis.xlsx（会自动追加时间戳）")

            source_layout.addRow("Workspace", self.workspace_display)
            source_layout.addRow("Output 目录", self.output_display)
            source_layout.addRow("Data_Input 目录", self.data_input_display)
            source_layout.addRow("", self.use_latest_checkbox)
            source_layout.addRow("手工 ModeB 报告", manual_path_row)
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
            self.open_output_button = QPushButton("打开输出目录")
            self.open_output_button.clicked.connect(lambda: self._open_path(self.paths.outputs_dir))
            self.save_settings_button = QPushButton("保存设置")
            self.save_settings_button.clicked.connect(self._save_settings)
            button_row.addWidget(self.generate_button)
            button_row.addWidget(self.open_output_button)
            button_row.addWidget(self.save_settings_button)
            button_row.addStretch(1)
            action_layout.addLayout(button_row)

            self.status_box = QTextEdit()
            self.status_box.setReadOnly(True)
            self.status_box.setPlaceholderText("这里会显示路径校验、旧版本提醒和生成结果。")
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
            self.workspace_display.setText(str(self.paths.user_workspace_dir))
            self.output_display.setText(str(self.paths.outputs_dir))
            self.data_input_display.setText(str(self.paths.workspace_input_dir))
            self.use_latest_checkbox.setChecked(_to_bool(self.settings.get("use_latest_report", "Yes")))
            self.manual_report_path_edit.setText(self.settings.get("manual_report_path", ""))
            self.output_name_edit.setText(self.settings.get("output_file_name", DEFAULT_OUTPUT_NAME))
            for index, edit in enumerate(self.product_edits, start=1):
                edit.setText(self.settings.get(f"product_{index}", ""))
            self._toggle_manual_report_state()

        def _collect_settings(self) -> dict[str, str]:
            payload = {
                "workspace_root": str(self.paths.user_workspace_dir),
                "use_latest_report": "Yes" if self.use_latest_checkbox.isChecked() else "No",
                "manual_report_path": self.manual_report_path_edit.text().strip(),
                "output_file_name": self.output_name_edit.text().strip() or DEFAULT_OUTPUT_NAME,
            }
            for index, edit in enumerate(self.product_edits, start=1):
                payload[f"product_{index}"] = edit.text().strip()
            return payload

        def _save_settings(self) -> None:
            settings_path = save_customer_case_settings(self.paths, self._collect_settings())
            self._append_status(f"设置已保存：{settings_path}")

        def _append_status(self, message: str) -> None:
            self.status_box.append(message)

        def _refresh_latest_report_hint(self) -> None:
            try:
                selection = resolve_modeb_report_selection(
                    output_dir=self.paths.outputs_dir,
                    manual_report_path=None,
                    use_latest_report=True,
                )
            except Exception:
                self.latest_report_hint.setText("当前 output 目录下还没有可用的 ModeB 报告。")
                return
            stamp = datetime.fromtimestamp(selection.selected_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            self.latest_report_hint.setText(f"最新 ModeB 报告：{selection.selected_path.name} | 时间：{stamp}")

        def _toggle_manual_report_state(self) -> None:
            manual_enabled = not self.use_latest_checkbox.isChecked()
            self.manual_report_path_edit.setEnabled(manual_enabled)
            self.browse_report_button.setEnabled(manual_enabled)

        def _browse_report(self) -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "选择 ModeB 报告",
                str(self.paths.outputs_dir),
                "Excel Workbook (*.xlsx)",
            )
            if path:
                self.manual_report_path_edit.setText(path)

        def _open_path(self, path: Path) -> None:
            if not path.exists():
                QMessageBox.warning(self, "路径不存在", f"以下路径不存在：\n{path}")
                return
            ok = QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
            if not ok:
                QMessageBox.warning(self, "打开失败", f"无法打开以下路径：\n{path}")

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

            try:
                selection = resolve_modeb_report_selection(
                    output_dir=self.paths.outputs_dir,
                    manual_report_path=self.manual_report_path_edit.text().strip(),
                    use_latest_report=self.use_latest_checkbox.isChecked(),
                )
            except Exception as exc:
                QMessageBox.warning(self, "报告选择失败", str(exc))
                self._append_status(f"报告选择失败：{exc}")
                return

            if not self.use_latest_checkbox.isChecked() and selection.latest_path and not selection.is_latest:
                answer = QMessageBox.question(
                    self,
                    "不是最新 ModeB 报告",
                    (
                        f"你当前选择的文件不是最新版本。\n\n"
                        f"当前选择：{selection.selected_path.name}\n"
                        f"最新文件：{selection.latest_path.name}\n\n"
                        f"是否仍继续分析你选择的旧版本？"
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
                    output_dir=self.paths.outputs_dir,
                    output_name=self.output_name_edit.text().strip() or DEFAULT_OUTPUT_NAME,
                    runtime_paths=self.paths,
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
        _show_native_error("ModeB Product Analysis Reporter", message)
        raise SystemExit(message)
    app = QApplication.instance() or QApplication(sys.argv)
    try:
        window = CustomerCaseMainWindow()
    except Exception as exc:
        message = f"启动工具时发生错误：\n{exc}"
        try:
            QMessageBox.critical(None, "ModeB Product Analysis Reporter", message)
        except Exception:
            _show_native_error("ModeB Product Analysis Reporter", message)
        return 1
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
