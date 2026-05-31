import csv
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QProcess, QProcessEnvironment, QTimer, Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)


def get_runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_runtime_base_dir()
SELECTIONS_FILE = BASE_DIR / "selections.json"
COURSE_CSV_FILE = BASE_DIR / "courses_schedule.csv"
QUESTIONS_SNAPSHOT_FILE = BASE_DIR / "questions_snapshot.json"
LOGIN_SCRIPT = BASE_DIR / "login.py"
DO_TABLE_SCRIPT = BASE_DIR / "do_table.py"
IS_FROZEN = bool(getattr(sys, "frozen", False))
PLAYWRIGHT_BROWSERS_DIR = BASE_DIR / "ms-playwright"

if IS_FROZEN:
    # EXE 模式固定瀏覽器下載位置，避免指向 _internal/.local-browsers 後找不到檔案。
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(PLAYWRIGHT_BROWSERS_DIR)


def _run_python_module(args: list[str]) -> tuple[bool, str]:
    """執行 python -m ...，回傳成功與輸出。"""
    cmd_candidates: list[list[str]] = []
    cmd_candidates.append([sys.executable, "-m", *args])
    if getattr(sys, "frozen", False):
        cmd_candidates.append(["python", "-m", *args])

    last_output = ""
    for cmd in cmd_candidates:
        try:
            result = subprocess.run(
                cmd,
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
            )
        except Exception as exc:
            last_output = str(exc)
            continue

        output = (result.stdout or "") + (result.stderr or "")
        if result.returncode == 0:
            return True, output.strip()
        last_output = output.strip()

    return False, last_output


def _run_playwright_install_chromium() -> tuple[bool, str]:
    """安裝 chromium。EXE 模式用 bundled driver；一般模式用 python -m playwright。"""
    if not IS_FROZEN:
        return _run_python_module(["playwright", "install", "chromium"])

    node_path = BASE_DIR / "_internal" / "playwright" / "driver" / "node.exe"
    cli_path = BASE_DIR / "_internal" / "playwright" / "driver" / "package" / "cli.js"
    if not node_path.exists() or not cli_path.exists():
        return False, "找不到 bundled playwright driver (node.exe 或 cli.js)。"

    env = os.environ.copy()
    env["PLAYWRIGHT_BROWSERS_PATH"] = str(PLAYWRIGHT_BROWSERS_DIR)

    try:
        result = subprocess.run(
            [str(node_path), str(cli_path), "install", "chromium"],
            cwd=str(BASE_DIR),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
    except Exception as exc:
        return False, str(exc)

    output = (result.stdout or "") + (result.stderr or "")
    return result.returncode == 0, output.strip()


def _check_playwright_and_chromium() -> tuple[bool, bool, str]:
    """檢查 playwright 套件與 chromium 是否可用。"""
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        return False, False, f"無法匯入 playwright: {exc}"

    try:
        with sync_playwright() as p:
            chromium_path = Path(p.chromium.executable_path)
            chromium_ok = chromium_path.exists()
            if chromium_ok:
                return True, True, ""
            return True, False, f"找不到 chromium 執行檔: {chromium_path}"
    except Exception as exc:
        return True, False, f"檢查 chromium 失敗: {exc}"


def ensure_runtime_dependencies() -> bool:
    """啟動前檢查執行環境，必要時可自動安裝。"""
    has_playwright, has_chromium, detail = _check_playwright_and_chromium()
    if has_playwright and has_chromium:
        return True

    missing: list[str] = []
    if not has_playwright:
        missing.append("playwright 套件")
    if has_playwright and not has_chromium:
        missing.append("playwright chromium")

    missing_lines = "\n- ".join(missing)

    message = (
        "啟動前檢查到缺少必要元件:\n"
        f"- {missing_lines}\n\n"
        "是否要立即自動安裝？"
    )
    if detail:
        message += f"\n\n詳細資訊:\n{detail}"

    reply = QMessageBox.question(
        None,
        "環境檢查",
        message,
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.Yes,
    )
    if reply != QMessageBox.Yes:
        return False

    if not has_playwright:
        if IS_FROZEN:
            QMessageBox.critical(
                None,
                "環境檢查失敗",
                "EXE 版本缺少 playwright 套件，請重新打包（需包含 playwright）。",
            )
            return False

        ok, output = _run_python_module(["pip", "install", "playwright"])
        if not ok:
            QMessageBox.critical(
                None,
                "安裝失敗",
                "安裝 playwright 失敗，請先執行 0_install.bat。\n\n"
                f"輸出:\n{output[-1200:] if output else '(無輸出)'}",
            )
            return False

    ok, output = _run_playwright_install_chromium()
    if not ok:
        QMessageBox.critical(
            None,
            "安裝失敗",
            "安裝 chromium 失敗，請先執行 0_install.bat。\n\n"
            f"輸出:\n{output[-1200:] if output else '(無輸出)'}",
        )
        return False

    has_playwright, has_chromium, detail = _check_playwright_and_chromium()
    if not (has_playwright and has_chromium):
        QMessageBox.critical(
            None,
            "環境檢查失敗",
            "安裝後仍無法使用 playwright/chromium，請手動執行 0_install.bat。\n\n"
            f"詳細資訊:\n{detail}",
        )
        return False

    QMessageBox.information(None, "完成", "執行環境檢查完成，playwright 與 chromium 可用。")
    return True


@dataclass
class CourseRow:
    status: str
    course_name: str
    time_slot: str


class ProfileTab(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.inputs: dict[str, QWidget] = {}
        self.form_layout = QFormLayout()
        self.form_layout.setLabelAlignment(Qt.AlignRight)
        self.form_layout.setFormAlignment(Qt.AlignTop)

        self.scroll_content = QWidget()
        self.scroll_content.setLayout(self.form_layout)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.scroll_content)

        self.status_label = QLabel("尚未載入")
        self.status_label.setStyleSheet("color: #3a4a5a;")

        btn_reload = QPushButton("重新載入")
        btn_reload.clicked.connect(self.load_data)

        btn_save = QPushButton("儲存 selections.json")
        btn_save.clicked.connect(self.save_data)

        button_row = QHBoxLayout()
        button_row.addWidget(btn_reload)
        button_row.addWidget(btn_save)
        button_row.addStretch()

        layout = QVBoxLayout(self)
        layout.addWidget(self.status_label)
        layout.addLayout(button_row)
        layout.addWidget(scroll)

        self.load_data()

    def _clear_form(self) -> None:
        while self.form_layout.rowCount() > 0:
            self.form_layout.removeRow(0)
        self.inputs.clear()

    def _create_input(self, key: str, value: str, options_map: dict[str, list[str]]) -> QWidget:
        options = options_map.get(key, [])
        if options:
            combo = QComboBox()
            combo.setEditable(False)
            combo.addItem("")
            for opt in options:
                combo.addItem(opt)
            idx = combo.findText(value)
            if idx >= 0:
                combo.setCurrentIndex(idx)
            return combo

        edit = QLineEdit()
        edit.setText(value)
        return edit

    def _read_selections(self) -> dict[str, str]:
        if not SELECTIONS_FILE.exists():
            return {}
        try:
            raw = json.loads(SELECTIONS_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
        if not isinstance(raw, dict):
            return {}
        return {str(k): str(v) if v is not None else "" for k, v in raw.items()}

    def _read_options_map(self) -> dict[str, list[str]]:
        if not QUESTIONS_SNAPSHOT_FILE.exists():
            return {}
        try:
            raw = json.loads(QUESTIONS_SNAPSHOT_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}

        options_map: dict[str, list[str]] = {}
        questions = raw.get("questions", []) if isinstance(raw, dict) else []
        for item in questions:
            if not isinstance(item, dict):
                continue
            label = item.get("label")
            options = item.get("options")
            if isinstance(label, str) and isinstance(options, list) and options:
                options_map[label] = [str(o) for o in options if isinstance(o, str)]

        email_label = raw.get("email_label") if isinstance(raw, dict) else None
        if isinstance(email_label, str) and email_label:
            options_map.setdefault(email_label, [])

        return options_map

    def load_data(self) -> None:
        selections = self._read_selections()
        options_map = self._read_options_map()
        self._clear_form()

        for key, value in selections.items():
            widget = self._create_input(key, value, options_map)
            self.inputs[key] = widget
            self.form_layout.addRow(QLabel(key), widget)

        self.status_label.setText(f"已載入 {len(selections)} 個欄位")

    def save_data(self) -> None:
        data: dict[str, str] = {}
        for key, widget in self.inputs.items():
            if isinstance(widget, QLineEdit):
                data[key] = widget.text().strip()
            elif isinstance(widget, QComboBox):
                data[key] = widget.currentText().strip()
            else:
                data[key] = ""

        SELECTIONS_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.status_label.setText("儲存完成")
        QMessageBox.information(self, "完成", "selections.json 已儲存")


class CourseTab(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[CourseRow] = []
        self.visible_source_indices: list[int] = []
        self.load_process: QProcess | None = None
        self.load_output_chunks: list[str] = []
        self.course_filter_checks: dict[str, QCheckBox] = {}
        self.weekday_filter_checks: dict[str, QCheckBox] = {}
        self._updating_table = False

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["選取", "狀態", "課程名", "時間"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.itemChanged.connect(self.on_table_item_changed)

        self.course_filter_box = QGroupBox("課程名篩選")
        self.course_filter_layout = QHBoxLayout(self.course_filter_box)

        self.weekday_filter_box = QGroupBox("星期關鍵字篩選")
        self.weekday_filter_layout = QHBoxLayout(self.weekday_filter_box)
        weekday_keywords = ["(一)", "(二)", "(三)", "(四)", "(五)", "(六)", "(日)"]
        for key in weekday_keywords:
            checkbox = QCheckBox(key)
            checkbox.stateChanged.connect(self.on_filter_changed)
            self.weekday_filter_checks[key] = checkbox
            self.weekday_filter_layout.addWidget(checkbox)
        self.weekday_filter_layout.addStretch()

        self.status_label = QLabel("尚未載入")

        self.btn_reload = QPushButton("重新載入")
        self.btn_reload.clicked.connect(self.load_data)

        self.btn_load_from_web = QPushButton("從網頁載入清單")
        self.btn_load_from_web.clicked.connect(self.load_list_from_web)

        self.btn_select_all = QPushButton("全選")
        self.btn_select_all.clicked.connect(self.select_all)

        self.btn_clear_all = QPushButton("全不選")
        self.btn_clear_all.clicked.connect(self.clear_all)

        self.btn_save = QPushButton("儲存到 courses_schedule.csv")
        self.btn_save.clicked.connect(self.save_data)

        row = QHBoxLayout()
        row.addWidget(self.btn_reload)
        row.addWidget(self.btn_load_from_web)
        row.addWidget(self.btn_select_all)
        row.addWidget(self.btn_clear_all)
        row.addWidget(self.btn_save)
        row.addStretch()

        layout = QVBoxLayout(self)
        layout.addWidget(self.status_label)
        layout.addLayout(row)
        layout.addWidget(self.course_filter_box)
        layout.addWidget(self.weekday_filter_box)
        layout.addWidget(self.table)

        self.load_data()

    def set_busy_state(self, busy: bool) -> None:
        self.btn_reload.setEnabled(not busy)
        self.btn_load_from_web.setEnabled(not busy)
        self.btn_select_all.setEnabled(not busy)
        self.btn_clear_all.setEnabled(not busy)
        self.btn_save.setEnabled(not busy)
        for checkbox in self.course_filter_checks.values():
            checkbox.setEnabled(not busy)
        for checkbox in self.weekday_filter_checks.values():
            checkbox.setEnabled(not busy)
        self.table.setEnabled(not busy)

    def rebuild_course_filters(self) -> None:
        old_checked = {name for name, check in self.course_filter_checks.items() if check.isChecked()}

        while self.course_filter_layout.count():
            item = self.course_filter_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        self.course_filter_checks = {}
        course_names = sorted({row.course_name for row in self.rows if row.course_name})
        for name in course_names:
            checkbox = QCheckBox(name)
            checkbox.setChecked(name in old_checked if old_checked else True)
            checkbox.stateChanged.connect(self.on_filter_changed)
            self.course_filter_checks[name] = checkbox
            self.course_filter_layout.addWidget(checkbox)

        self.course_filter_layout.addStretch()

    def selected_course_filters(self) -> set[str]:
        return {name for name, check in self.course_filter_checks.items() if check.isChecked()}

    def selected_weekday_filters(self) -> set[str]:
        return {key for key, check in self.weekday_filter_checks.items() if check.isChecked()}

    def row_matches_filter(self, row: CourseRow) -> bool:
        selected_courses = self.selected_course_filters()
        if selected_courses and row.course_name not in selected_courses:
            return False

        selected_weekdays = self.selected_weekday_filters()
        if selected_weekdays:
            if not any(key in row.time_slot for key in selected_weekdays):
                return False

        return True

    def sync_visible_checks_to_rows(self) -> None:
        for table_row, source_idx in enumerate(self.visible_source_indices):
            check_item = self.table.item(table_row, 0)
            if check_item is None:
                continue
            checked = check_item.checkState() == Qt.Checked
            original_status = self.rows[source_idx].status.strip()
            if checked:
                self.rows[source_idx].status = "V"
            elif original_status.upper() == "V":
                self.rows[source_idx].status = ""
            else:
                self.rows[source_idx].status = original_status

    def refresh_table_by_filters(self) -> None:
        self.sync_visible_checks_to_rows()

        self.visible_source_indices = []
        for idx, row in enumerate(self.rows):
            if self.row_matches_filter(row):
                self.visible_source_indices.append(idx)

        self._updating_table = True
        self.table.setRowCount(len(self.visible_source_indices))
        for table_idx, source_idx in enumerate(self.visible_source_indices):
            row = self.rows[source_idx]
            checked = Qt.Checked if row.status.upper() == "V" else Qt.Unchecked

            check_item = QTableWidgetItem()
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            check_item.setCheckState(checked)
            self.table.setItem(table_idx, 0, check_item)

            status_item = QTableWidgetItem(row.status)
            status_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(table_idx, 1, status_item)

            course_item = QTableWidgetItem(row.course_name)
            course_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(table_idx, 2, course_item)

            time_item = QTableWidgetItem(row.time_slot)
            time_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(table_idx, 3, time_item)
        self._updating_table = False

        self.status_label.setText(f"已載入 {len(self.rows)} 筆課程，篩選顯示 {len(self.visible_source_indices)} 筆")

    def on_filter_changed(self, *_args) -> None:
        if self._updating_table:
            return
        self.refresh_table_by_filters()

    def on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._updating_table:
            return
        if item.column() != 0:
            return

        row_idx = item.row()
        if row_idx < 0 or row_idx >= len(self.visible_source_indices):
            return

        source_idx = self.visible_source_indices[row_idx]
        checked = item.checkState() == Qt.Checked
        original_status = self.rows[source_idx].status.strip()
        if checked:
            self.rows[source_idx].status = "V"
        elif original_status.upper() == "V":
            self.rows[source_idx].status = ""
        else:
            self.rows[source_idx].status = original_status

        self._updating_table = True
        status_item = self.table.item(row_idx, 1)
        if status_item is not None:
            status_item.setText(self.rows[source_idx].status)
        self._updating_table = False

    def load_list_from_web(self) -> None:
        if self.load_process is not None:
            QMessageBox.warning(self, "忙碌中", "目前正在從網頁載入清單")
            return

        if not LOGIN_SCRIPT.exists():
            QMessageBox.critical(self, "錯誤", "找不到 login.py")
            return

        reply = QMessageBox.question(
            self,
            "確認載入",
            "將開啟瀏覽器並重新抓取課程清單，確定要繼續嗎？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self.load_output_chunks = []
        process = QProcess(self)
        process.setProgram(sys.executable)
        process.setArguments([str(LOGIN_SCRIPT), "--auto-close"])
        process.setWorkingDirectory(str(BASE_DIR))
        env = QProcessEnvironment.systemEnvironment()
        env.insert("BOOKING_BASE_DIR", str(BASE_DIR))
        process.setProcessEnvironment(env)
        process.setProcessChannelMode(QProcess.MergedChannels)
        process.readyReadStandardOutput.connect(self.on_load_from_web_output)
        process.finished.connect(self.on_load_from_web_finished)

        self.load_process = process
        self.set_busy_state(True)
        self.status_label.setText("從網頁載入中，請在瀏覽器完成登入或操作...")
        process.start()

    def on_load_from_web_output(self) -> None:
        if self.load_process is None:
            return
        data = self.load_process.readAllStandardOutput().data()
        text = data.decode(errors="ignore")
        if text:
            self.load_output_chunks.append(text)

    def on_load_from_web_finished(self, exit_code: int, _status) -> None:
        if self.load_process is not None:
            data = self.load_process.readAllStandardOutput().data()
            text = data.decode(errors="ignore")
            if text:
                self.load_output_chunks.append(text)

        self.load_process = None
        self.set_busy_state(False)

        if exit_code == 0:
            self.load_data()
            QMessageBox.information(self, "完成", "已從網頁更新課程清單")
            return

        logs = "".join(self.load_output_chunks).strip()
        if logs:
            logs = logs[-800:]
            QMessageBox.critical(self, "載入失敗", f"從網頁載入失敗 (exit_code={exit_code})\n\n最近輸出:\n{logs}")
        else:
            QMessageBox.critical(self, "載入失敗", f"從網頁載入失敗 (exit_code={exit_code})")
        self.status_label.setText("從網頁載入失敗")

    def load_data(self) -> None:
        self.rows = []
        self.visible_source_indices = []
        self.table.setRowCount(0)

        if not COURSE_CSV_FILE.exists():
            self.status_label.setText("找不到 courses_schedule.csv")
            return

        with COURSE_CSV_FILE.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                status = str(row.get("狀態", "")).strip()
                course_name = str(row.get("課程名", "")).strip()
                time_slot = str(row.get("時間", "")).strip()
                self.rows.append(CourseRow(status=status, course_name=course_name, time_slot=time_slot))

        self.rebuild_course_filters()
        self.refresh_table_by_filters()

    def select_all(self) -> None:
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item is not None:
                item.setCheckState(Qt.Checked)

    def clear_all(self) -> None:
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item is not None:
                item.setCheckState(Qt.Unchecked)

    def save_data(self) -> None:
        self.sync_visible_checks_to_rows()

        output_rows: list[CourseRow] = []
        for row in self.rows:
            output_rows.append(CourseRow(status=row.status, course_name=row.course_name, time_slot=row.time_slot))

        with COURSE_CSV_FILE.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["狀態", "課程名", "時間"])
            for row in output_rows:
                writer.writerow([row.status, row.course_name, row.time_slot])

        self.status_label.setText("儲存完成")
        QMessageBox.information(self, "完成", "courses_schedule.csv 已儲存")


class RunTab(QWidget):
    CONFIRM_ONCE_THEN_AUTO = "第一次需要按確認之後全自動"
    CONFIRM_EVERY_TIME = "每次都需要按確認"
    FULL_AUTO = "全自動執行"
    DELAY_0_1 = "0-1秒"
    DELAY_5_10 = "5-10秒"
    DELAY_30_60 = "30-60秒"
    DELAY_60_120 = "60-120秒"

    def __init__(self) -> None:
        super().__init__()
        self.process: QProcess | None = None
        self.current_script = ""
        self.execution_targets: list[tuple[str, str]] = []
        self.current_running_index: int = -1
        self.batch_progress_re = re.compile(r"\[BATCH\]\s*第\s*(\d+)\s*/\s*(\d+)\s*筆")

        self.status_label = QLabel("待命")

        self.exec_list_label = QLabel("本次執行清單：尚未建立")

        self.exec_table = QTableWidget(0, 3)
        self.exec_table.setHorizontalHeaderLabels(["課程名", "時間", "即時狀態"])
        self.exec_table.horizontalHeader().setStretchLastSection(True)
        self.exec_table.setAlternatingRowColors(True)
        self.exec_table.setEditTriggers(QTableWidget.NoEditTriggers)

        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)

        self.status_poll_timer = QTimer(self)
        self.status_poll_timer.setInterval(1000)
        self.status_poll_timer.timeout.connect(self.refresh_execution_status_from_csv)

        self.confirm_mode_combo = QComboBox()
        self.confirm_mode_combo.addItems([
            self.CONFIRM_ONCE_THEN_AUTO,
            self.CONFIRM_EVERY_TIME,
            self.FULL_AUTO,
        ])
        self.confirm_mode_combo.setCurrentText(self.CONFIRM_ONCE_THEN_AUTO)

        self.delay_mode_combo = QComboBox()
        self.delay_mode_combo.addItems([
            self.DELAY_0_1,
            self.DELAY_5_10,
            self.DELAY_30_60,
            self.DELAY_60_120,
        ])
        self.delay_mode_combo.setCurrentText(self.DELAY_0_1)

        self.antibot_jitter_checkbox = QCheckBox("隨機等防機器人認證")
        self.antibot_jitter_checkbox.setChecked(False)

        self.btn_run_do_table = QPushButton("執行 do_table.py")
        self.btn_run_do_table.clicked.connect(self.run_do_table)

        self.btn_stop = QPushButton("停止執行")
        self.btn_stop.clicked.connect(self.stop_script)
        self.btn_stop.setEnabled(False)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("執行確認模式"))
        mode_row.addWidget(self.confirm_mode_combo)
        mode_row.addSpacing(12)
        mode_row.addWidget(QLabel("Submit前隨機延遲"))
        mode_row.addWidget(self.delay_mode_combo)
        mode_row.addSpacing(12)
        mode_row.addWidget(self.antibot_jitter_checkbox)
        mode_row.addStretch()

        row = QHBoxLayout()
        row.addWidget(self.btn_run_do_table)
        row.addWidget(self.btn_stop)
        row.addStretch()

        help_box = QGroupBox("提示")
        help_layout = QVBoxLayout(help_box)
        help_layout.addWidget(QLabel("1. do_table.py 會依 courses_schedule.csv 狀態=V 進行批次填表。"))
        help_layout.addWidget(QLabel("2. 確認模式可控制每次執行前是否顯示確認視窗。"))

        layout = QVBoxLayout(self)
        layout.addWidget(self.status_label)
        layout.addLayout(mode_row)
        layout.addLayout(row)
        layout.addWidget(help_box)
        layout.addWidget(self.exec_list_label)
        layout.addWidget(self.exec_table)
        layout.addWidget(self.log_output)

    def append_log(self, text: str) -> None:
        if not text:
            return
        self.log_output.appendPlainText(text.rstrip())
        sb = self.log_output.verticalScrollBar()
        sb.setValue(sb.maximum())

    def set_running_ui(self, running: bool) -> None:
        self.confirm_mode_combo.setEnabled(not running)
        self.delay_mode_combo.setEnabled(not running)
        self.antibot_jitter_checkbox.setEnabled(not running)
        self.btn_run_do_table.setEnabled(not running)
        self.btn_stop.setEnabled(running)

    def _current_confirm_mode(self) -> str:
        mode = self.confirm_mode_combo.currentText()
        if mode == self.CONFIRM_EVERY_TIME:
            return "every"
        if mode == self.FULL_AUTO:
            return "none"
        return "once"

    def _current_delay_mode(self) -> str:
        mode = self.delay_mode_combo.currentText()
        if mode == self.DELAY_5_10:
            return "5-10"
        if mode == self.DELAY_30_60:
            return "30-60"
        if mode == self.DELAY_60_120:
            return "60-120"
        return "0-1"

    def _read_pending_targets(self) -> list[tuple[str, str]]:
        if not COURSE_CSV_FILE.exists():
            return []

        targets: list[tuple[str, str]] = []
        with COURSE_CSV_FILE.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                status = str(row.get("狀態", "")).strip().upper()
                if status != "V":
                    continue
                course_name = str(row.get("課程名", "")).strip()
                time_slot = str(row.get("時間", "")).strip()
                targets.append((course_name, time_slot))
        return targets

    def _set_exec_row_status(self, row: int, text: str) -> None:
        item = self.exec_table.item(row, 2)
        if item is None:
            item = QTableWidgetItem()
            self.exec_table.setItem(row, 2, item)
        item.setText(text)

    def build_execution_list(self) -> None:
        self.execution_targets = self._read_pending_targets()
        self.current_running_index = -1
        self.exec_table.setRowCount(0)

        if not self.execution_targets:
            self.exec_list_label.setText("本次執行清單：目前沒有狀態=V 的課程")
            return

        self.exec_table.setRowCount(len(self.execution_targets))
        for idx, (course_name, time_slot) in enumerate(self.execution_targets):
            self.exec_table.setItem(idx, 0, QTableWidgetItem(course_name))
            self.exec_table.setItem(idx, 1, QTableWidgetItem(time_slot))
            self.exec_table.setItem(idx, 2, QTableWidgetItem("待執行"))

        self.exec_list_label.setText(f"本次執行清單：共 {len(self.execution_targets)} 筆")

    def _build_status_queue_from_csv(self) -> dict[tuple[str, str], list[str]]:
        queue_map: dict[tuple[str, str], list[str]] = {}
        if not COURSE_CSV_FILE.exists():
            return queue_map

        with COURSE_CSV_FILE.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                course_name = str(row.get("課程名", "")).strip()
                time_slot = str(row.get("時間", "")).strip()
                status = str(row.get("狀態", "")).strip().upper()
                key = (course_name, time_slot)
                queue_map.setdefault(key, []).append(status)
        return queue_map

    def refresh_execution_status_from_csv(self) -> None:
        if not self.execution_targets:
            return

        queue_map = self._build_status_queue_from_csv()
        completed = 0
        for idx, target in enumerate(self.execution_targets):
            status_list = queue_map.get(target, [])
            csv_status = status_list.pop(0) if status_list else ""

            if csv_status == "D":
                ui_status = "已完成"
                completed += 1
            elif idx == self.current_running_index and self.process is not None:
                ui_status = "執行中"
            else:
                ui_status = "待執行"

            self._set_exec_row_status(idx, ui_status)

        self.exec_list_label.setText(f"本次執行清單：共 {len(self.execution_targets)} 筆，已完成 {completed} 筆")

    def run_do_table(self) -> None:
        if self.process is not None:
            QMessageBox.warning(self, "忙碌中", "目前已有執行中的程序")
            return

        if not DO_TABLE_SCRIPT.exists():
            QMessageBox.critical(self, "錯誤", "找不到 do_table.py")
            return

        self.build_execution_list()

        self.run_script(
            DO_TABLE_SCRIPT,
            {
                "BOOKING_CONFIRM_MODE": self._current_confirm_mode(),
                "BOOKING_SUBMIT_DELAY_RANGE": self._current_delay_mode(),
                "BOOKING_ANTIBOT_JITTER": "1" if self.antibot_jitter_checkbox.isChecked() else "0",
            },
        )

    def run_script(self, script_path: Path, extra_env: dict[str, str] | None = None) -> None:
        if self.process is not None:
            QMessageBox.warning(self, "忙碌中", "目前已有執行中的程序")
            return

        if not script_path.exists():
            QMessageBox.critical(self, "錯誤", f"找不到檔案: {script_path.name}")
            return

        self.current_script = script_path.name
        self.log_output.clear()
        self.append_log(f"[INFO] 開始執行 {self.current_script}")

        process = QProcess(self)
        process.setProgram(sys.executable)
        process.setArguments([str(script_path)])
        process.setWorkingDirectory(str(BASE_DIR))
        env = QProcessEnvironment.systemEnvironment()
        env.insert("BOOKING_BASE_DIR", str(BASE_DIR))
        if extra_env:
            for key, value in extra_env.items():
                env.insert(str(key), str(value))
        process.setProcessEnvironment(env)
        process.setProcessChannelMode(QProcess.MergedChannels)
        process.readyReadStandardOutput.connect(self.on_output)
        process.finished.connect(self.on_finished)

        self.process = process
        self.set_running_ui(True)
        self.status_label.setText(f"執行中: {self.current_script}")
        process.start()
        self.status_poll_timer.start()

    def stop_script(self) -> None:
        if self.process is None:
            return
        self.append_log("[INFO] 使用者要求停止程序")
        self.process.kill()

    def on_output(self) -> None:
        if self.process is None:
            return
        data = self.process.readAllStandardOutput().data()
        text = data.decode(errors="ignore")
        self.append_log(text)

        for line in text.splitlines():
            matched = self.batch_progress_re.search(line)
            if matched:
                idx = int(matched.group(1)) - 1
                if 0 <= idx < len(self.execution_targets):
                    self.current_running_index = idx
        self.refresh_execution_status_from_csv()

    def on_finished(self, exit_code: int, _status) -> None:
        self.append_log(f"[INFO] 執行結束，exit_code={exit_code}")
        self.status_label.setText("待命")
        self.status_poll_timer.stop()
        self.current_running_index = -1
        self.refresh_execution_status_from_csv()
        self.process = None
        self.current_script = ""
        self.set_running_ui(False)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Taipei Arena Ice Booking Tool")
        self.resize(1100, 760)

        tabs = QTabWidget()
        self.profile_tab = ProfileTab()
        self.course_tab = CourseTab()
        self.run_tab = RunTab()

        tabs.addTab(self.profile_tab, "基本資料")
        tabs.addTab(self.course_tab, "選課清單")
        tabs.addTab(self.run_tab, "執行中心")

        self.setCentralWidget(tabs)
        self._build_menu()

    def _build_menu(self) -> None:
        menu = self.menuBar().addMenu("檔案")

        reload_action = QAction("全部重載", self)
        reload_action.triggered.connect(self.reload_all)
        menu.addAction(reload_action)

        quit_action = QAction("離開", self)
        quit_action.triggered.connect(self.close)
        menu.addAction(quit_action)

    def reload_all(self) -> None:
        self.profile_tab.load_data()
        self.course_tab.load_data()
        QMessageBox.information(self, "完成", "資料已重載")


def main() -> None:
    app = QApplication(sys.argv)
    if not ensure_runtime_dependencies():
        sys.exit(1)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
