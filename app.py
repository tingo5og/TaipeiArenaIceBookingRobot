import csv
import os
import re
import subprocess
import sys
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QProcess, QProcessEnvironment, QTimer, Qt
from PySide6.QtGui import QAction, QColor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)


BASE_DIR = Path(__file__).resolve().parent
COURSE_CSV_FILE = BASE_DIR / "courses_schedule.csv"
SCAN_SCRIPT = BASE_DIR / "scan_courses.py"
DO_TABLE_SCRIPT = BASE_DIR / "do_table.py"
# 需與 metro_site.BROWSER_PROFILE_DIR 相同。
BROWSER_PROFILE_DIR = BASE_DIR / "browser_profile"
BOOKING_LOGIN_URL = "https://ssl.metro.taipei/TAPICELand/Student/Login.aspx"
# 需與 metro_site.DEBUG_PORT 相同；掃描/預約腳本透過此埠接上登入用瀏覽器。
BROWSER_DEBUG_PORT = int(os.environ.get("BOOKING_DEBUG_PORT", "9333"))

# 統一 Python 子程序輸出編碼，降低 Windows 碼頁造成的亂碼風險。
os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")


def _run_python_module(args: list[str]) -> tuple[bool, str]:
    """執行 python -m ...，回傳成功與輸出。"""
    try:
        result = subprocess.run(
            [sys.executable, "-m", *args],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except Exception as exc:
        return False, str(exc)

    output = ((result.stdout or "") + (result.stderr or "")).strip()
    return result.returncode == 0, output


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


def _resolve_chromium_executable() -> tuple[Path | None, str]:
    """取得 Playwright Chromium 執行檔路徑。"""
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        return None, f"無法匯入 playwright: {exc}"

    try:
        with sync_playwright() as p:
            chromium_path = Path(p.chromium.executable_path)
    except Exception as exc:
        return None, f"無法取得 Chromium 路徑: {exc}"

    if not chromium_path.exists():
        return None, f"找不到 Chromium 執行檔: {chromium_path}"
    return chromium_path, ""


def decode_process_output(data: bytes) -> str:
    try:
        return data.decode("utf-8", errors="replace")
    except Exception:
        return data.decode(errors="replace")


def force_close_profile_browsers(profile_dir: Path) -> tuple[bool, str]:
    """強制關閉使用指定 user-data-dir 的瀏覽器程序。"""
    profile_path = str(profile_dir.resolve()).replace("'", "''")
    ps_script = f"""
$target = [System.IO.Path]::GetFullPath('{profile_path}').ToLowerInvariant()
$targets = @("chrome.exe", "msedge.exe", "chromium.exe")
$killed = @()

Get-CimInstance Win32_Process | Where-Object {{
    $_.Name -and $targets -contains $_.Name.ToLowerInvariant() -and $_.CommandLine -and $_.CommandLine.ToLowerInvariant().Contains("--user-data-dir=" + $target)
}} | ForEach-Object {{
    try {{
        Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop
        $killed += $_.ProcessId
    }} catch {{
    }}
}}

if ($killed.Count -gt 0) {{
    Write-Output ("KILLED=" + (($killed | Sort-Object -Unique) -join ","))
}} else {{
    Write-Output "KILLED="
}}
"""

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except Exception as exc:
        return False, str(exc)

    output = ((result.stdout or "") + (result.stderr or "")).strip()
    if result.returncode != 0:
        return False, output or f"PowerShell 結束碼: {result.returncode}"

    return True, output


def is_login_browser_running() -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{BROWSER_DEBUG_PORT}/json/version", timeout=1) as resp:
            return resp.status == 200
    except Exception:
        return False


def open_login_browser(parent: QWidget) -> None:
    """開啟（或沿用）帶遠端除錯埠的 Chromium 並進入登入頁，讓使用者手動登入。"""
    if is_login_browser_running():
        try:
            req = urllib.request.Request(
                f"http://127.0.0.1:{BROWSER_DEBUG_PORT}/json/new?{BOOKING_LOGIN_URL}", method="PUT"
            )
            urllib.request.urlopen(req, timeout=3).close()
        except Exception:
            pass
        QMessageBox.information(parent, "登入", "登入用瀏覽器已開啟，請在該視窗完成登入（已登入可略過）。")
        return

    chromium_path, detail = _resolve_chromium_executable()
    if chromium_path is None:
        QMessageBox.critical(parent, "開啟失敗", f"無法取得 Chromium。\n\n{detail}")
        return

    # 先關掉其他占用同一個 profile 的瀏覽器，否則新視窗不會帶上除錯埠。
    ok, close_detail = force_close_profile_browsers(BROWSER_PROFILE_DIR)
    if not ok:
        QMessageBox.warning(parent, "提醒", f"無法預先關閉占用 profile 的瀏覽器程序。\n\n{close_detail}")

    try:
        BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(
            [
                str(chromium_path),
                f"--user-data-dir={BROWSER_PROFILE_DIR}",
                f"--remote-debugging-port={BROWSER_DEBUG_PORT}",
                "--profile-directory=Default",
                "--no-first-run",
                "--new-window",
                BOOKING_LOGIN_URL,
            ],
            cwd=str(BASE_DIR),
        )
    except Exception as exc:
        QMessageBox.critical(parent, "開啟失敗", f"啟動 Chromium 失敗: {exc}")
        return

    QMessageBox.information(
        parent,
        "登入",
        "已開啟瀏覽器，請輸入帳號、密碼、驗證碼並按「登入」。\n\n"
        "登入後請保持此瀏覽器開啟，「掃描課程列表」「執行預約」「執行取消預約」都會沿用這個登入狀態。",
    )


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
        ok, output = _run_python_module(["pip", "install", "playwright"])
        if not ok:
            QMessageBox.critical(
                None,
                "安裝失敗",
                "安裝 playwright 失敗，請先執行 0_run.bat。\n\n"
                f"輸出:\n{output[-1200:] if output else '(無輸出)'}",
            )
            return False

    ok, output = _run_python_module(["playwright", "install", "chromium"])
    if not ok:
        QMessageBox.critical(
            None,
            "安裝失敗",
            "安裝 chromium 失敗，請先執行 0_run.bat。\n\n"
            f"輸出:\n{output[-1200:] if output else '(無輸出)'}",
        )
        return False

    has_playwright, has_chromium, detail = _check_playwright_and_chromium()
    if not (has_playwright and has_chromium):
        QMessageBox.critical(
            None,
            "環境檢查失敗",
            "安裝後仍無法使用 playwright/chromium，請手動執行 0_run.bat。\n\n"
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


def is_registered_status(status: str) -> bool:
    """D=已報名、X=已報名且勾選待取消。"""
    return status.strip().upper() in ("D", "X")


def checked_status(status: str, checked: bool) -> str:
    """勾選框狀態對應的 CSV 狀態：未報名的課勾選=V(待預約)；已報名的課勾選=X(待取消)。"""
    normalized = status.strip().upper()
    if is_registered_status(normalized):
        return "X" if checked else "D"
    if checked:
        return "V"
    return "" if normalized == "V" else status.strip()


def status_display(status: str) -> str:
    normalized = status.strip().upper()
    if normalized == "D":
        return "已報名"
    if normalized == "X":
        return "已報名→待取消"
    if normalized == "V":
        return "待預約"
    return status


# 執行中心各動作要處理的 CSV 狀態：book=V(待預約)、cancel=X(待取消)。
RUN_TARGET_STATUS = {"book": "V", "cancel": "X"}

COURSE_GROUP_ORDER = ["滑冰基礎班", "花式課", "冰球課"]

# 時段依課程開始時間判斷：[開始小時, 結束小時)
PERIOD_RANGES = {"白天(9-14)": (9, 14), "晚上(14-21)": (14, 21)}


def course_group(course_name: str) -> str:
    """篩選用的課程分組；初級/進階班合併，CSV 與預約仍使用原課程名。"""
    if "曲棍球" in course_name or "冰球" in course_name:
        return "冰球課"
    if "花式" in course_name:
        return "花式課"
    return course_name


def time_period(time_slot: str) -> str:
    """由「2026/11/26(四)｜17:00~18:00｜教練」取開始時間，回傳所屬時段名稱；不在任何時段回傳空字串。"""
    match = re.search(r"｜(\d{1,2}):(\d{2})", time_slot)
    if not match:
        return ""
    start = int(match.group(1)) + int(match.group(2)) / 60
    for name, (low, high) in PERIOD_RANGES.items():
        if low <= start < high:
            return name
    return ""


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

        self.course_filter_box = QGroupBox("課程篩選")
        self.course_filter_layout = QHBoxLayout(self.course_filter_box)

        self.weekday_filter_box = QGroupBox("星期篩選")
        self.weekday_filter_layout = QHBoxLayout(self.weekday_filter_box)
        weekday_keywords = ["(一)", "(二)", "(三)", "(四)", "(五)", "(六)", "(日)"]
        for key in weekday_keywords:
            checkbox = QCheckBox(key)
            checkbox.stateChanged.connect(self.on_filter_changed)
            self.weekday_filter_checks[key] = checkbox
            self.weekday_filter_layout.addWidget(checkbox)
        self.weekday_filter_layout.addStretch()

        self.period_filter_checks: dict[str, QCheckBox] = {}
        self.period_filter_box = QGroupBox("時段篩選（依開始時間）")
        self.period_filter_layout = QHBoxLayout(self.period_filter_box)
        for key in PERIOD_RANGES:
            checkbox = QCheckBox(key)
            checkbox.stateChanged.connect(self.on_filter_changed)
            self.period_filter_checks[key] = checkbox
            self.period_filter_layout.addWidget(checkbox)
        self.period_filter_layout.addStretch()

        self.status_label = QLabel("尚未載入")

        self.btn_login = QPushButton("1. 登入")
        self.btn_login.clicked.connect(lambda: open_login_browser(self))

        self.btn_load_from_web = QPushButton("2. 掃描課程列表")
        self.btn_load_from_web.clicked.connect(self.load_list_from_web)

        self.btn_reload = QPushButton("重新載入 CSV")
        self.btn_reload.clicked.connect(self.load_data)

        self.btn_select_all = QPushButton("全選")
        self.btn_select_all.clicked.connect(self.select_all)

        self.btn_clear_all = QPushButton("全不選")
        self.btn_clear_all.clicked.connect(self.clear_all)

        self.btn_save = QPushButton("儲存到 courses_schedule.csv")
        self.btn_save.clicked.connect(self.save_data)

        step_row = QHBoxLayout()
        step_row.addWidget(self.btn_login)
        step_row.addWidget(self.btn_load_from_web)
        step_row.addWidget(QLabel("3. 勾選要上的課（或勾選已報名的課以取消），到「執行中心」按 4 或 5"))
        step_row.addStretch()

        row = QHBoxLayout()
        row.addWidget(self.btn_reload)
        row.addWidget(self.btn_select_all)
        row.addWidget(self.btn_clear_all)
        row.addWidget(self.btn_save)
        row.addStretch()

        layout = QVBoxLayout(self)
        layout.addLayout(step_row)
        layout.addWidget(self.status_label)
        layout.addLayout(row)
        layout.addWidget(self.course_filter_box)
        layout.addWidget(self.weekday_filter_box)
        layout.addWidget(self.period_filter_box)
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
        for checkbox in self.period_filter_checks.values():
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
        groups = {course_group(row.course_name) for row in self.rows if row.course_name}
        group_names = sorted(groups, key=lambda g: (COURSE_GROUP_ORDER.index(g) if g in COURSE_GROUP_ORDER else len(COURSE_GROUP_ORDER), g))
        for name in group_names:
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

    def selected_period_filters(self) -> set[str]:
        return {key for key, check in self.period_filter_checks.items() if check.isChecked()}

    def row_matches_filter(self, row: CourseRow) -> bool:
        selected_courses = self.selected_course_filters()
        if selected_courses and course_group(row.course_name) not in selected_courses:
            return False

        selected_weekdays = self.selected_weekday_filters()
        if selected_weekdays:
            if not any(key in row.time_slot for key in selected_weekdays):
                return False

        selected_periods = self.selected_period_filters()
        if selected_periods:
            if time_period(row.time_slot) not in selected_periods:
                return False

        return True

    def sync_visible_checks_to_rows(self) -> None:
        for table_row, source_idx in enumerate(self.visible_source_indices):
            check_item = self.table.item(table_row, 0)
            if check_item is None or not (check_item.flags() & Qt.ItemIsUserCheckable):
                continue
            checked = check_item.checkState() == Qt.Checked
            self.rows[source_idx].status = checked_status(self.rows[source_idx].status, checked)

    def _style_status_item(self, item: QTableWidgetItem, status: str) -> None:
        item.setText(status_display(status))
        normalized = status.strip().upper()
        if normalized == "X":
            item.setForeground(QColor("#c0392b"))
        elif normalized == "D":
            item.setForeground(QColor("#1b7f3b"))
        else:
            item.setForeground(QColor("#000000"))

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
            normalized = row.status.strip().upper()

            # 未報名的課勾選=待預約；已報名的課勾選=待取消。
            check_item = QTableWidgetItem()
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            check_item.setCheckState(Qt.Checked if normalized in ("V", "X") else Qt.Unchecked)
            self.table.setItem(table_idx, 0, check_item)

            status_item = QTableWidgetItem()
            status_item.setFlags(Qt.ItemIsEnabled)
            self._style_status_item(status_item, row.status)
            self.table.setItem(table_idx, 1, status_item)

            course_item = QTableWidgetItem(row.course_name)
            course_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(table_idx, 2, course_item)

            time_item = QTableWidgetItem(row.time_slot)
            time_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(table_idx, 3, time_item)
        self._updating_table = False

        registered_count = sum(1 for row in self.rows if is_registered_status(row.status))
        self.status_label.setText(
            f"已載入 {len(self.rows)} 筆課程（已報名 {registered_count} 堂），篩選顯示 {len(self.visible_source_indices)} 筆"
        )

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
        self.rows[source_idx].status = checked_status(self.rows[source_idx].status, checked)

        self._updating_table = True
        status_item = self.table.item(row_idx, 1)
        if status_item is not None:
            self._style_status_item(status_item, self.rows[source_idx].status)
        self._updating_table = False

    def load_list_from_web(self) -> None:
        if self.load_process is not None:
            QMessageBox.warning(self, "忙碌中", "目前正在掃描課程列表")
            return

        if not SCAN_SCRIPT.exists():
            QMessageBox.critical(self, "錯誤", f"找不到 {SCAN_SCRIPT.name}")
            return

        if not is_login_browser_running():
            QMessageBox.warning(self, "尚未登入", "請先按「1. 登入」開啟瀏覽器並完成登入。")
            return

        # 先把目前勾選存進 CSV，掃描後會依 CSV 保留勾選狀態。
        if self.rows:
            self.save_data(show_message=False)

        self.load_output_chunks = []
        process = QProcess(self)
        process.setProgram(sys.executable)
        process.setArguments(["-u", str(SCAN_SCRIPT)])
        process.setWorkingDirectory(str(BASE_DIR))
        env = QProcessEnvironment.systemEnvironment()
        env.insert("BOOKING_BASE_DIR", str(BASE_DIR))
        env.insert("PYTHONUNBUFFERED", "1")
        env.insert("PYTHONUTF8", "1")
        env.insert("PYTHONIOENCODING", "utf-8")
        process.setProcessEnvironment(env)
        process.setProcessChannelMode(QProcess.MergedChannels)
        process.readyReadStandardOutput.connect(self.on_load_from_web_output)
        process.finished.connect(self.on_load_from_web_finished)

        self.load_process = process
        self.set_busy_state(True)
        self.status_label.setText("掃描課程列表中...（若登入已失效，請在瀏覽器重新登入）")
        process.start()

    def on_load_from_web_output(self) -> None:
        if self.load_process is None:
            return
        data = self.load_process.readAllStandardOutput().data()
        text = decode_process_output(data)
        if text:
            self.load_output_chunks.append(text)

    def on_load_from_web_finished(self, exit_code: int, _status) -> None:
        if self.load_process is not None:
            data = self.load_process.readAllStandardOutput().data()
            text = decode_process_output(data)
            if text:
                self.load_output_chunks.append(text)

        self.load_process = None
        self.set_busy_state(False)

        if exit_code == 0:
            self.load_data()
            registered = [row for row in self.rows if is_registered_status(row.status)]
            lines = [f"已掃描課程列表，共 {len(self.rows)} 堂課", f"已報名 {len(registered)} 堂："]
            lines += [f"- {row.course_name} {row.time_slot}" for row in registered] or ["（無）"]
            QMessageBox.information(self, "完成", "\n".join(lines))
            return

        logs = "".join(self.load_output_chunks).strip()
        if logs:
            logs = logs[-800:]
            QMessageBox.critical(self, "掃描失敗", f"掃描課程列表失敗 (exit_code={exit_code})\n\n最近輸出:\n{logs}")
        else:
            QMessageBox.critical(self, "掃描失敗", f"掃描課程列表失敗 (exit_code={exit_code})")
        self.status_label.setText("掃描課程列表失敗")

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
        # 全選只勾「待預約」，不會把已報名的課標成待取消。
        for r, source_idx in enumerate(self.visible_source_indices):
            item = self.table.item(r, 0)
            if item is not None and not is_registered_status(self.rows[source_idx].status):
                item.setCheckState(Qt.Checked)

    def clear_all(self) -> None:
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item is not None:
                item.setCheckState(Qt.Unchecked)

    def save_data(self, show_message: bool = True) -> None:
        self.sync_visible_checks_to_rows()

        # 延後匯入：metro_site 依賴 playwright，需等啟動時的環境檢查完成。
        from metro_site import write_csv_atomic

        rows = [["狀態", "課程名", "時間"]]
        rows += [[row.status, row.course_name, row.time_slot] for row in self.rows]
        write_csv_atomic(COURSE_CSV_FILE, rows)

        self.status_label.setText("儲存完成")
        if show_message:
            QMessageBox.information(self, "完成", "courses_schedule.csv 已儲存")


class RunTab(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.process: QProcess | None = None
        self.current_script = ""
        self.sync_course_data_before_run: Callable[[], None] | None = None
        # 執行結束後通知「選課清單」重新讀取 CSV，避免舊勾選狀態覆蓋執行結果。
        self.on_run_finished: Callable[[], None] | None = None
        self.run_action = "book"
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

        self.btn_run_do_table = QPushButton("4. 執行預約")
        self.btn_run_do_table.clicked.connect(lambda: self.run_do_table("book"))

        self.btn_run_cancel = QPushButton("5. 執行取消預約")
        self.btn_run_cancel.clicked.connect(lambda: self.run_do_table("cancel"))

        self.btn_stop = QPushButton("停止執行")
        self.btn_stop.clicked.connect(self.stop_script)
        self.btn_stop.setEnabled(False)

        row = QHBoxLayout()
        row.addWidget(self.btn_run_do_table)
        row.addWidget(self.btn_run_cancel)
        row.addWidget(self.btn_stop)
        row.addStretch()

        help_box = QGroupBox("提示")
        help_layout = QVBoxLayout(help_box)
        help_layout.addWidget(QLabel("1. 請先在「選課清單」按「1. 登入」並保持登入用瀏覽器開啟。"))
        help_layout.addWidget(QLabel("2. 「4. 執行預約」會逐筆預約「待預約」的課，成功後改成「已報名」。"))
        help_layout.addWidget(QLabel("3. 「5. 執行取消預約」會逐筆取消「已報名→待取消」的課，並回到上課列表確認已取消。"))

        layout = QVBoxLayout(self)
        layout.addWidget(self.status_label)
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
        self.btn_run_do_table.setEnabled(not running)
        self.btn_run_cancel.setEnabled(not running)
        self.btn_stop.setEnabled(running)

    def _read_pending_targets(self) -> list[tuple[str, str]]:
        if not COURSE_CSV_FILE.exists():
            return []

        targets: list[tuple[str, str]] = []
        with COURSE_CSV_FILE.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                status = str(row.get("狀態", "")).strip().upper()
                if status != RUN_TARGET_STATUS[self.run_action]:
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
            label = "待預約" if self.run_action == "book" else "待取消"
            self.exec_list_label.setText(f"本次執行清單：目前沒有「{label}」的課程")
            return

        self.exec_table.setRowCount(len(self.execution_targets))
        for idx, (course_name, time_slot) in enumerate(self.execution_targets):
            self.exec_table.setItem(idx, 0, QTableWidgetItem(course_name))
            self.exec_table.setItem(idx, 1, QTableWidgetItem(time_slot))
            self.exec_table.setItem(idx, 2, QTableWidgetItem("待執行"))

        self.exec_list_label.setText(f"本次執行清單：共 {len(self.execution_targets)} 筆")

    def _build_status_queue_from_csv(self) -> dict[tuple[str, str], list[str]] | None:
        """讀取 CSV 狀態；讀取失敗（例如檔案正被替換）時回傳 None，留待下次輪詢。"""
        queue_map: dict[tuple[str, str], list[str]] = {}
        if not COURSE_CSV_FILE.exists():
            return queue_map

        try:
            with COURSE_CSV_FILE.open("r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    course_name = str(row.get("課程名", "")).strip()
                    time_slot = str(row.get("時間", "")).strip()
                    status = str(row.get("狀態", "")).strip().upper()
                    key = (course_name, time_slot)
                    queue_map.setdefault(key, []).append(status)
        except (OSError, UnicodeDecodeError, csv.Error):
            return None
        return queue_map

    def refresh_execution_status_from_csv(self) -> None:
        if not self.execution_targets:
            return

        queue_map = self._build_status_queue_from_csv()
        if queue_map is None:
            return
        completed = 0
        for idx, target in enumerate(self.execution_targets):
            status_list = queue_map.get(target, [])
            csv_status = status_list.pop(0) if status_list else ""

            if csv_status != RUN_TARGET_STATUS[self.run_action]:
                ui_status = "已完成"
                completed += 1
            elif idx == self.current_running_index and self.process is not None:
                ui_status = "執行中"
            else:
                ui_status = "待執行"

            self._set_exec_row_status(idx, ui_status)

        self.exec_list_label.setText(f"本次執行清單：共 {len(self.execution_targets)} 筆，已完成 {completed} 筆")

    def run_do_table(self, action: str = "book") -> None:
        if self.process is not None:
            QMessageBox.warning(self, "忙碌中", "目前已有執行中的程序")
            return

        if not DO_TABLE_SCRIPT.exists():
            QMessageBox.critical(self, "錯誤", "找不到 do_table.py")
            return

        if not is_login_browser_running():
            QMessageBox.warning(self, "尚未登入", "請先到「選課清單」按「1. 登入」開啟瀏覽器並完成登入。")
            return

        if self.sync_course_data_before_run is not None:
            try:
                self.sync_course_data_before_run()
            except Exception as exc:
                QMessageBox.critical(self, "錯誤", f"同步選課清單失敗: {exc}")
                return

        self.run_action = action
        self.build_execution_list()

        self.run_script(
            DO_TABLE_SCRIPT,
            {
                "BOOKING_ACTION": action,
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
        process.setArguments(["-u", str(script_path)])
        process.setWorkingDirectory(str(BASE_DIR))
        env = QProcessEnvironment.systemEnvironment()
        env.insert("BOOKING_BASE_DIR", str(BASE_DIR))
        env.insert("PYTHONUNBUFFERED", "1")
        env.insert("PYTHONUTF8", "1")
        env.insert("PYTHONIOENCODING", "utf-8")
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
        text = decode_process_output(data)
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
        if self.on_run_finished is not None:
            self.on_run_finished()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Taipei Arena Ice Booking Tool")
        self.resize(1100, 760)

        tabs = QTabWidget()
        self.course_tab = CourseTab()
        self.run_tab = RunTab()
        self.run_tab.sync_course_data_before_run = lambda: self.course_tab.save_data(show_message=False)
        self.run_tab.on_run_finished = self.course_tab.load_data

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
