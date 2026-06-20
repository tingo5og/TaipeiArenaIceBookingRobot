import asyncio
import csv
import json
import os
import random
import re
import sys
import time
import traceback
from pathlib import Path

from playwright.async_api import Locator, Page, async_playwright

def get_base_dir() -> Path:
    from_env = os.environ.get("BOOKING_BASE_DIR", "").strip()
    if from_env:
        return Path(from_env).resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()

USER_DATA_DIR = str(BASE_DIR / "google_profile")
FALLBACK_USER_DATA_DIR = str(BASE_DIR / "google_profile_playwright")
FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSchP7dRjEOyofx3V6cu7do8UM_WghZRuB9QnwMwQAvceZ2evg/viewform?pli=1&pli=1"
SELECTIONS_FILE = BASE_DIR / "selections.json"
QUESTIONS_SNAPSHOT_FILE = BASE_DIR / "questions_snapshot.json"
COURSE_CSV_FILE = BASE_DIR / "courses_schedule.csv"
PAUSE_ON_STALL = True
AUTO_UPDATE_SELECTIONS = False
DRY_RUN = False  # True 時會走到提交按鈕但不真正點擊,用於測試流程和檢查填寫內容。

BATCH_KEY = "批次課程"  # selections.json 裡存放批次清單的 key
COURSE_TIME_KEY = "課程時間"
COURSE_COMBO_KEY = "課程與時間"
COURSE_TIME_LABELS = {
    "滑冰基礎班",
    "花式初級班",
    "花式進階班",
    "冰球初級班",
    "冰球進階班",
}

CONFIRM_MODE = os.environ.get("BOOKING_CONFIRM_MODE", "once").strip().lower()
CONFIRM_MODE_ONCE = "once"
CONFIRM_MODE_EVERY = "every"
CONFIRM_MODE_NONE = "none"
_ONCE_CONFIRMED = False
SUBMIT_DELAY_RANGE = os.environ.get("BOOKING_SUBMIT_DELAY_RANGE", "0-1").strip()
ANTIBOT_JITTER_ENABLED = os.environ.get("BOOKING_ANTIBOT_JITTER", "0").strip() == "1"
VERIFY_POLL_INTERVAL_SEC = float(os.environ.get("BOOKING_VERIFY_POLL_INTERVAL_SEC", "1.2"))
VERIFY_CLEAR_STABLE_ROUNDS = int(os.environ.get("BOOKING_VERIFY_CLEAR_STABLE_ROUNDS", "2"))
VERIFY_MAX_WAIT_SEC = int(os.environ.get("BOOKING_VERIFY_MAX_WAIT_SEC", "0"))
PROFILE_STRATEGY = os.environ.get("BOOKING_PROFILE_STRATEGY", "auto").strip().lower()


def configure_stdio_utf8() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is None:
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


configure_stdio_utf8()


def safe_print(message: str) -> None:
    try:
        print(message)
        return
    except UnicodeEncodeError:
        pass

    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    encoded = message.encode(encoding, errors="replace")
    buffer = getattr(sys.stdout, "buffer", None)
    if buffer is not None:
        buffer.write(encoded + b"\n")
        buffer.flush()
        return

    sys.stdout.write(encoded.decode(encoding, errors="replace") + "\n")
    sys.stdout.flush()


def log(step: str, message: str) -> None:
    ts = time.strftime("%H:%M:%S")
    safe_print(f"[{ts}] [{step}] {message}")


def cleanup_profile_lock_files(profile_dir: Path) -> None:
    lock_paths = [
        profile_dir / "SingletonLock",
        profile_dir / "SingletonCookie",
        profile_dir / "SingletonSocket",
        profile_dir / "lockfile",
        profile_dir / "Default" / "LOCK",
    ]
    removed: list[str] = []

    for lock_path in lock_paths:
        try:
            if lock_path.exists():
                lock_path.unlink()
                removed.append(str(lock_path.relative_to(profile_dir)))
        except Exception as exc:
            log("WARN", f"無法移除 lock 檔 {lock_path.name}: {exc}")

    if removed:
        log("BROWSER", f"已清理 lock 檔: {', '.join(removed)}")


async def launch_persistent_context_with_fallback(playwright_instance):
    launch_kwargs = {
        "headless": False,
        "slow_mo": 300,
        "args": ["--disable-blink-features=AutomationControlled"],
    }

    primary_dir = Path(USER_DATA_DIR)
    fallback_dir = Path(FALLBACK_USER_DATA_DIR)

    primary_dir.mkdir(parents=True, exist_ok=True)
    cleanup_profile_lock_files(primary_dir)

    try:
        context = await playwright_instance.chromium.launch_persistent_context(
            str(primary_dir),
            **launch_kwargs,
        )
        log("BROWSER", f"使用主 profile 啟動: {primary_dir}")
        return context
    except Exception as exc:
        log("WARN", f"主 profile 啟動失敗: {exc}")
        if PROFILE_STRATEGY == "primary_only":
            raise RuntimeError("主 profile 啟動失敗，且目前為 primary_only 模式") from exc
        log("WARN", "可能是 profile 正被其他瀏覽器使用或發生鎖檔衝突，改用備援 profile")

    fallback_dir.mkdir(parents=True, exist_ok=True)
    cleanup_profile_lock_files(fallback_dir)

    context = await playwright_instance.chromium.launch_persistent_context(
        str(fallback_dir),
        **launch_kwargs,
    )
    log("BROWSER", f"使用備援 profile 啟動: {fallback_dir}")
    return context


def get_submit_delay_seconds() -> float:
    if SUBMIT_DELAY_RANGE == "5-10":
        return random.uniform(5, 10)
    if SUBMIT_DELAY_RANGE == "30-60":
        return random.uniform(30, 60)
    if SUBMIT_DELAY_RANGE == "300-600":
        return random.uniform(300, 600)
    # 舊值相容,避免舊環境變數仍使用 60-120。
    if SUBMIT_DELAY_RANGE == "60-120":
        return random.uniform(60, 120)
    return random.uniform(0, 1)


async def delay_before_submit_if_needed() -> None:
    seconds = get_submit_delay_seconds()
    log("DELAY", f"Submit 前延遲 {seconds:.2f} 秒")
    await asyncio.sleep(seconds)


def get_antibot_jitter_seconds() -> float:
    # 基準：每頁送出前 1-3 秒。
    low, high = 1.0, 3.0

    # 若 Submit 延遲選到 30 秒以上，對應拉長每頁等待，讓節奏更像人工操作。
    if SUBMIT_DELAY_RANGE == "30-60":
        low, high = 3.0, 6.0
    elif SUBMIT_DELAY_RANGE == "300-600":
        low, high = 12.0, 20.0
    elif SUBMIT_DELAY_RANGE == "60-120":
        low, high = 6.0, 12.0

    return random.uniform(low, high)


async def delay_before_page_action_if_needed() -> None:
    if not ANTIBOT_JITTER_ENABLED:
        return
    seconds = get_antibot_jitter_seconds()
    log("JITTER", f"頁面送出前隨機等待 {seconds:.2f} 秒")
    await asyncio.sleep(seconds)


def should_confirm_before_submit() -> bool:
    if CONFIRM_MODE == CONFIRM_MODE_NONE:
        return False
    if CONFIRM_MODE == CONFIRM_MODE_EVERY:
        return True
    # 預設 once
    return not _ONCE_CONFIRMED


def build_submit_confirmation_text(selections: dict) -> str:
    fields = [
        "請問您是新生還是舊生？",
        "請問您的名字？(須與會員資料相同)",
        "請問您的生日?",
        "已了解以上說明內容",
        "你的電子郵件",
    ]

    lines: list[str] = ["請確認本次送出資料：", "", "基本資料"]
    for key in fields:
        lines.append(f"- {key}: {str(selections.get(key, '')).strip()}")

    lines.append("")
    lines.append("這次執行課程名,時間")

    combo = selections.get(COURSE_COMBO_KEY)
    if isinstance(combo, list) and len(combo) >= 2:
        course_name = str(combo[0]).strip()
        time_slot = str(combo[1]).strip()
        lines.append(f"- {course_name},{time_slot}")
    else:
        lines.append("- (本次無課程批次資訊)")

    return "\n".join(lines)


def ask_submit_confirmation(selections: dict) -> bool:
    text = build_submit_confirmation_text(selections)

    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        confirmed = messagebox.askyesno("執行前確認", text, parent=root)
        root.destroy()
        return confirmed
    except Exception as exc:
        log("WARN", f"無法顯示圖形確認視窗，改用終端機確認: {exc}")
        print(text)
        choice = input("確認送出？輸入 y 繼續，其餘取消: ").strip().lower()
        return choice == "y"




def normalize_label(raw_text: str) -> str:
    return raw_text.split("\n")[0].replace("*", "").strip()


def normalize_key(text: str) -> str:
    return re.sub(r"\s+", "", text).strip().lower()


def normalize_date_answer(text: str) -> str | None:
    raw = text.strip()
    match = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", raw)
    if not match:
        return None
    year, month, day = match.groups()
    return f"{year}-{int(month):02d}-{int(day):02d}"


def split_date_parts(text: str) -> tuple[str, str, str] | None:
    normalized = normalize_date_answer(text)
    if not normalized:
        return None
    year, month, day = normalized.split("-")
    return year, month, day


def is_email_like_label(label: str) -> bool:
    normalized = normalize_key(label)
    return ("電子郵件" in normalized) or ("email" in normalized)


def is_course_time_label(label: str) -> bool:
    return label in COURSE_TIME_LABELS


def is_course_selector_label(label: str) -> bool:
    return bool(re.search(r"想預約|預約.*課程", label))


def fuzzy_match_option(pattern: str, options: list[str]) -> str | None:
    """支援 * 分隔的多段子字串模糊比對"""
    parts = [p for p in pattern.split("*") if p]
    for opt in options:
        if all(p in opt for p in parts):
            return opt
    return None


def resolve_selection_key(selections: dict, label: str) -> str | None:
    # 新格式：課程與時間 array,優先處理
    if COURSE_COMBO_KEY in selections and isinstance(selections.get(COURSE_COMBO_KEY), list):
        if is_course_time_label(label) or is_course_selector_label(label):
            return COURSE_COMBO_KEY

    if is_course_time_label(label):
        if COURSE_TIME_KEY in selections:
            return COURSE_TIME_KEY
        for key in COURSE_TIME_LABELS:
            if key in selections:
                return key

    if label in selections:
        return label

    if is_email_like_label(label):
        non_empty: list[str] = []
        fallback: list[str] = []
        for key, value in selections.items():
            if is_email_like_label(key):
                fallback.append(key)
                if value.strip():
                    non_empty.append(key)
        if non_empty:
            return non_empty[0]
        if fallback:
            return fallback[0]

    return None


def get_answer_for_label(selections: dict, label: str) -> str:
    matched_key = resolve_selection_key(selections, label)
    if not matched_key:
        return ""
    value = selections.get(matched_key)
    if isinstance(value, list):
        # 課程與時間: [課程類型pattern, 時間pattern]
        idx = 1 if is_course_time_label(label) else 0
        return str(value[idx]).strip() if idx < len(value) else ""
    return str(value).strip() if value is not None else ""


def page_signature(page: Page, items: list[dict[str, object]]) -> str:
    labels = [str(x.get("label", "")) for x in items]
    return f"{page.url}::{'|'.join(labels)}"


async def get_robot_verification_signals(page: Page) -> list[str]:
    """蒐集人機驗證跡象,回傳命中的訊號清單。"""
    signals: list[str] = []

    url = page.url.lower()
    if any(x in url for x in ["recaptcha", "captcha", "challenge", "sorry"]):
        signals.append(f"url={page.url}")

    body_text = (await page.locator("body").inner_text()).lower()
    keywords = [
        "i'm not a robot",
        "verify you are human",
        "不是機器人",
        "人機驗證",
        "驗證你不是",
        "異常流量",
        "unusual traffic",
    ]
    for keyword in keywords:
        if keyword in body_text:
            signals.append(f"keyword={keyword}")
            break

    # 一般 Google Form 頁面可能帶隱藏的 reCAPTCHA 資源，只有可見 challenge 元件才算。
    visible_challenge_selectors = [
        'iframe[src*="recaptcha/api2/anchor"]',
        'iframe[src*="recaptcha/api2/bframe"]',
        'iframe[title*="reCAPTCHA" i]',
        'div.g-recaptcha',
        '#recaptcha',
        'span[aria-label*="robot" i]',
        'span[aria-label*="不是機器人"]',
        'div[role="checkbox"][aria-label*="robot" i]',
        'div[role="checkbox"][aria-label*="不是機器人"]',
    ]
    for selector in visible_challenge_selectors:
        locator = page.locator(selector)
        count = await locator.count()
        for idx in range(count):
            try:
                if await locator.nth(idx).is_visible():
                    signals.append(f"visible={selector}")
                    break
            except Exception:
                continue

    return signals


async def is_robot_verification_page(page: Page) -> bool:
    """偵測目前頁面是否為人機驗證流程。"""
    return len(await get_robot_verification_signals(page)) > 0


async def handle_robot_verification_if_needed(page: Page) -> str | None:
    """若偵測到人機驗證則自動輪詢,完成後自動繼續。"""
    signals = await get_robot_verification_signals(page)
    if not signals:
        return None

    log("VERIFY", f"偵測到人機驗證,判斷依據: {'; '.join(signals)}")
    log("VERIFY", "請在瀏覽器完成驗證,系統會自動檢查並繼續")

    start = time.monotonic()
    clear_rounds = 0
    last_progress_log_at = -999.0

    while True:
        try:
            await page.wait_for_load_state("domcontentloaded", timeout=2000)
        except Exception:
            pass
        await page.wait_for_timeout(int(max(0.2, VERIFY_POLL_INTERVAL_SEC) * 1000))

        signals = await get_robot_verification_signals(page)
        elapsed = time.monotonic() - start

        if not signals:
            clear_rounds += 1
            if clear_rounds >= max(1, VERIFY_CLEAR_STABLE_ROUNDS):
                log("VERIFY", "人機驗證已通過，繼續流程")
                return "verified"
        else:
            clear_rounds = 0
            if elapsed - last_progress_log_at >= 5:
                log("VERIFY", f"仍在驗證流程中 ({elapsed:.0f} 秒),目前依據: {'; '.join(signals)}")
                last_progress_log_at = elapsed

        if VERIFY_MAX_WAIT_SEC > 0 and elapsed >= VERIFY_MAX_WAIT_SEC:
            log("VERIFY", f"等待人機驗證逾時 ({VERIFY_MAX_WAIT_SEC} 秒),停止批次")
            return "aborted"


def load_batch_from_csv() -> list[tuple[list[str], int]] | None:
    """從 CSV 讀取狀態=V 的課程,回傳 [([課程名, 時間], 行號), ...]"""
    if not COURSE_CSV_FILE.exists():
        log("CSV", f"{COURSE_CSV_FILE.name} 不存在")
        return None

    try:
        batch: list[tuple[list[str], int]] = []
        with COURSE_CSV_FILE.open("r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if not header or header[0] != "狀態":
                log("CSV", "CSV 格式不正確(缺少標頭或首欄不是狀態)")
                return None

            for row_idx, row in enumerate(reader, start=2):  # 從第2行開始,row_idx=2,3,4...
                if len(row) < 3:
                    continue
                status, course_name, time_slot = row[0].strip(), row[1].strip(), row[2].strip()
                if status.upper() == "V":
                    batch.append(([course_name, time_slot], row_idx))

        if batch:
            log("CSV", f"已從 CSV 讀取 {len(batch)} 筆批次課程")
            return batch
        else:
            log("CSV", "CSV 中找不到狀態=V 的課程")
            return None
    except Exception as exc:
        log("ERROR", f"讀取 CSV 失敗: {exc}")
        return None


def update_csv_row_status(row_num: int, status: str) -> None:
    """更新 CSV 特定行(1-indexed)的狀態"""
    if not COURSE_CSV_FILE.exists():
        return

    try:
        with COURSE_CSV_FILE.open("r", encoding="utf-8-sig") as f:
            lines = f.readlines()
        
        # row_num 是 1-indexed(第幾行),需要轉換為 0-indexed
        line_idx = row_num - 1
        if line_idx < 0 or line_idx >= len(lines):
            log("CSV", f"行號 {row_num} 超出範圍")
            return
        
        # 解析該行
        line = lines[line_idx]
        parts = line.rstrip("\n").rstrip("\r").split(",")
        if parts:
            parts[0] = status
            lines[line_idx] = ",".join(parts) + "\n"
        
        with COURSE_CSV_FILE.open("w", encoding="utf-8-sig") as f:
            f.writelines(lines)
        
        log("CSV", f"已更新第 {row_num} 行狀態為: {status}")
    except Exception as exc:
        log("ERROR", f"更新 CSV 失敗: {exc}")


def load_selections() -> dict[str, str]:
    if not SELECTIONS_FILE.exists():
        log("LOAD", "selections.json 不存在,回傳空設定")
        return {}

    try:
        data = json.loads(SELECTIONS_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        data = {}

    if not isinstance(data, dict):
        data = {}

    normalized: dict = {}
    for key, value in data.items():
        if isinstance(key, str):
            normalized[key] = value if isinstance(value, list) else (str(value) if value is not None else "")

    log("LOAD", f"已載入 selections,共 {len(normalized)} 題")

    return normalized


def save_selections(selections: dict[str, str]) -> None:
    SELECTIONS_FILE.write_text(
        json.dumps(selections, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def save_questions_snapshot(snapshot: dict[str, dict[str, object]]) -> None:
    QUESTIONS_SNAPSHOT_FILE.write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


async def get_options_from_role(question: Locator, role_name: str) -> list[str]:
    role_items = question.locator(f'div[role="{role_name}"]')
    try:
        raw_values = await role_items.evaluate_all(
            "els => els.map(e => (e.getAttribute('data-value') || e.textContent || '').trim())"
        )
    except Exception as exc:
        # 某些頁面切換瞬間元素會 detach,略過該題避免整體中斷。
        log("WARN", f"讀取 {role_name} 選項失敗,已略過: {exc}")
        return []

    options: list[str] = []
    for value in raw_values:
        label = str(value).strip()
        if label and label not in options:
            options.append(label)
    return options


async def inspect_current_page(page: Page) -> tuple[list[dict[str, object]], dict[str, dict[str, object]]]:
    result: list[dict[str, object]] = []
    snapshot: dict[str, dict[str, object]] = {}

    question_items = page.locator('div[role="listitem"]')
    question_count = await question_items.count()
    log("SCAN", f"本頁偵測到 listitem 數量: {question_count}")

    for idx in range(question_count):
        question = question_items.nth(idx)
        heading = question.locator('div[role="heading"]').first
        if await heading.count() == 0:
            continue

        label_text = normalize_label(await heading.inner_text())
        question_type = "unknown"
        options: list[str] = []

        has_date_input = await question.locator('input[type="date"]').count() > 0
        has_text_input = await question.locator('input[type="text"], input[type="email"], textarea').count() > 0
        has_any_input = await question.locator('input:not([type="hidden"]), textarea').count() > 0
        date_like_label = bool(re.search(r"生日|日期|date", label_text, re.I))
        has_multi_input = await question.locator('input:not([type="hidden"])').count() >= 2

        if has_date_input:
            question_type = "date"
        elif date_like_label and has_multi_input:
            question_type = "date_parts"
        elif has_text_input:
            question_type = "text"
        elif has_any_input:
            question_type = "text"
        elif await question.locator('div[role="radio"]').count() > 0:
            question_type = "radio"
            options = await get_options_from_role(question, "radio")
        elif await question.locator('div[role="checkbox"]').count() > 0:
            question_type = "checkbox"
            options = await get_options_from_role(question, "checkbox")
        elif await question.locator('div[role="listbox"]').count() > 0:
            question_type = "dropdown"

        result.append({
            "label": label_text,
            "type": question_type,
            "question": question,
            "options": options,
        })
        snapshot[label_text] = {
            "type": question_type,
            "options": options,
        }

    log("SCAN", f"可處理題目數: {len(result)}")

    return result, snapshot


def sync_selections_with_questions(
    selections: dict[str, str],
    snapshot: dict[str, dict[str, object]],
) -> tuple[dict[str, str], list[str]]:
    added_questions: list[str] = []
    if not AUTO_UPDATE_SELECTIONS:
        for label in snapshot:
            if resolve_selection_key(selections, label) is None:
                added_questions.append(label)
        return selections, added_questions

    for label in snapshot:
        if resolve_selection_key(selections, label) is None:
            selections[label] = ""
            added_questions.append(label)
    return selections, added_questions


async def fill_question(item: dict[str, object], selections: dict[str, str]) -> None:
    label = str(item["label"])
    answer = get_answer_for_label(selections, label)
    if not answer:
        log("FILL", f"略過未填答案: {label}")
        return

    matched_key = resolve_selection_key(selections, label)
    if matched_key and matched_key != label:
        log("FILL", f"同義鍵映射: {label} -> {matched_key}")

    question = item["question"]
    question_type = item["type"]
    log("FILL", f"開始填寫: {label} -> {answer}")

    if question_type == "text":
        input_box = question.locator('input[type="text"], input[type="email"], textarea, input:not([type="hidden"])').first
        if await input_box.count() > 0:
            await input_box.fill(answer)
            log("FILL", f"已填文字欄: {label}")
        else:
            log("FILL", f"找不到文字欄: {label}")
        return

    if question_type == "date":
        date_box = question.locator('input[type="date"]').first
        if await date_box.count() == 0:
            log("FILL", f"找不到日期欄: {label}")
            return

        date_text = normalize_date_answer(answer)
        if not date_text:
            log("FILL", f"日期格式不符,請用 YYYY/MM/DD 或 YYYY-MM-DD: {label} -> {answer}")
            return

        await date_box.fill(date_text)
        current = (await date_box.input_value()).strip()
        if current == date_text:
            log("FILL", f"已填日期欄: {label} -> {date_text}")
        else:
            log("FILL", f"日期欄填寫可能失敗: {label},目前值: {current}")
        return

    if question_type == "date_parts":
        parts = split_date_parts(answer)
        if not parts:
            log("FILL", f"日期格式不符,請用 YYYY/MM/DD 或 YYYY-MM-DD: {label} -> {answer}")
            return
        year, month, day = parts

        inputs = question.locator('input:not([type="hidden"])')
        input_count = await inputs.count()
        if input_count == 0:
            log("FILL", f"找不到日期分欄輸入框: {label}")
            return

        filled = False
        for idx in range(input_count):
            inp = inputs.nth(idx)
            hint = " ".join(
                [
                    (await inp.get_attribute("aria-label") or ""),
                    (await inp.get_attribute("placeholder") or ""),
                    (await inp.get_attribute("name") or ""),
                ]
            ).lower()

            if re.search(r"year|年|yyyy", hint):
                await inp.fill(year)
                filled = True
            elif re.search(r"month|月|mm", hint):
                await inp.fill(month)
                filled = True
            elif re.search(r"day|日|dd", hint):
                await inp.fill(day)
                filled = True

        if not filled:
            # 無法辨識欄位語意時,用常見順序填入：月/日/年
            values = [month, day, year]
            for idx in range(min(input_count, 3)):
                await inputs.nth(idx).fill(values[idx])

        log("FILL", f"已填日期分欄: {label} -> {year}-{month}-{day}")
        return

    if question_type == "radio":
        options_list = [str(o) for o in item.get("options", [])]
        actual_answer = answer
        if options_list and answer not in options_list:
            matched = fuzzy_match_option(answer, options_list)
            if matched:
                log("FILL", f"模糊比對: {answer!r} → {matched!r}")
                actual_answer = matched
        option = question.locator(f'div[role="radio"][data-value="{actual_answer}"]').first
        if await option.count() > 0:
            await option.click()
            log("FILL", f"已選單選: {label}")
        else:
            log("FILL", f"找不到單選值: {label} -> {actual_answer}")
        return

    if question_type == "checkbox":
        targets = [x.strip() for x in answer.split(",") if x.strip()]
        for target in targets:
            option = question.locator(f'div[role="checkbox"][data-value="{target}"]').first
            if await option.count() > 0:
                checked = await option.get_attribute("aria-checked")
                if checked != "true":
                    await option.click()
                    log("FILL", f"已勾選複選: {label} -> {target}")
            else:
                log("FILL", f"找不到複選值: {label} -> {target}")
        return

    if question_type == "dropdown":
        dropdown = question.locator('div[role="listbox"]').first
        if await dropdown.count() == 0:
            log("FILL", f"找不到下拉欄: {label}")
            return
        await dropdown.click()
        await question.page.get_by_role("option", name=answer, exact=True).click()
        log("FILL", f"已選下拉: {label}")


def resolve_actual_answer(item: dict[str, object], answer: str) -> str:
    question_type = str(item.get("type", ""))
    if question_type != "radio":
        return answer

    options_list = [str(o) for o in item.get("options", [])]
    if options_list and answer not in options_list:
        matched = fuzzy_match_option(answer, options_list)
        if matched:
            return matched
    return answer


def build_preferred_radio_indices(
    inspected: list[dict[str, object]],
    selections: dict[str, str],
) -> dict[tuple[str, str, str], int]:
    """同一 dedupe key 有多個單選群組時,優先挑選包含目標選項的群組。"""
    preferred: dict[tuple[str, str, str], int] = {}
    for idx, item in enumerate(inspected):
        label = str(item.get("label", "")).strip()
        question_type = str(item.get("type", "")).strip()
        if question_type != "radio":
            continue

        matched_key = resolve_selection_key(selections, label) or label
        answer = get_answer_for_label(selections, label).strip()
        if not answer:
            continue

        dedupe_key = (matched_key, question_type, answer)
        actual_answer = resolve_actual_answer(item, answer)
        options_list = [str(o) for o in item.get("options", [])]
        if actual_answer in options_list and dedupe_key not in preferred:
            preferred[dedupe_key] = idx
    return preferred


async def fill_questions_for_current_page(inspected: list[dict[str, object]], selections: dict[str, str]) -> None:
    preferred_radio_indices = build_preferred_radio_indices(inspected, selections)
    filled_question_keys: set[tuple[str, str, str]] = set()

    for idx, item in enumerate(inspected):
        label = str(item.get("label", "")).strip()
        question_type = str(item.get("type", "")).strip()
        matched_key = resolve_selection_key(selections, label) or label
        answer = get_answer_for_label(selections, label).strip()
        dedupe_key = (matched_key, question_type, answer)

        if answer:
            if dedupe_key in filled_question_keys:
                log("FILL", f"略過重複題目: {label} -> {matched_key}")
                continue

            if question_type == "radio":
                preferred_idx = preferred_radio_indices.get(dedupe_key)
                if preferred_idx is not None and preferred_idx != idx:
                    log("FILL", f"略過替代單選群組: {label} -> {matched_key}")
                    continue

        await fill_question(item, selections)

        if answer:
            filled_question_keys.add(dedupe_key)


async def verify_question_state(item: dict[str, object], selections: dict[str, str]) -> str | None:
    label = str(item.get("label", "")).strip()
    answer = get_answer_for_label(selections, label).strip()
    if not answer:
        return None

    question = item["question"]
    question_type = str(item.get("type", "")).strip()

    if question_type == "radio":
        actual_answer = resolve_actual_answer(item, answer)
        option = question.locator(f'div[role="radio"][data-value="{actual_answer}"]').first
        if await option.count() == 0:
            return f"{label}: 找不到應選單選值 {actual_answer}"
        checked = await option.get_attribute("aria-checked")
        if checked != "true":
            return f"{label}: 單選目前未選中 {actual_answer}"
        return None

    if question_type == "checkbox":
        targets = [x.strip() for x in answer.split(",") if x.strip()]
        for target in targets:
            option = question.locator(f'div[role="checkbox"][data-value="{target}"]').first
            if await option.count() == 0:
                return f"{label}: 找不到應選複選值 {target}"
            checked = await option.get_attribute("aria-checked")
            if checked != "true":
                return f"{label}: 複選目前未勾選 {target}"
        return None

    if question_type == "dropdown":
        dropdown = question.locator('div[role="listbox"]').first
        if await dropdown.count() == 0:
            return f"{label}: 找不到下拉欄"
        current = normalize_label((await dropdown.inner_text()).strip())
        if answer not in current:
            return f"{label}: 下拉目前值不是 {answer}"
        return None

    return None


async def verify_current_page_answers(inspected: list[dict[str, object]], selections: dict[str, str]) -> list[str]:
    mismatches: list[str] = []
    preferred_radio_indices = build_preferred_radio_indices(inspected, selections)
    seen_keys: set[tuple[str, str, str]] = set()

    for idx, item in enumerate(inspected):
        label = str(item.get("label", "")).strip()
        question_type = str(item.get("type", "")).strip()
        matched_key = resolve_selection_key(selections, label) or label
        answer = get_answer_for_label(selections, label).strip()
        dedupe_key = (matched_key, question_type, answer)

        if answer and dedupe_key in seen_keys:
            continue

        if answer and question_type == "radio":
            preferred_idx = preferred_radio_indices.get(dedupe_key)
            if preferred_idx is not None and preferred_idx != idx:
                continue

        mismatch = await verify_question_state(item, selections)
        if mismatch:
            mismatches.append(mismatch)

        if answer:
            seen_keys.add(dedupe_key)

    return mismatches


async def goto_form_and_wait_login(page: Page) -> None:
    log("NAV", f"前往表單: {FORM_URL}")
    await page.goto(FORM_URL, wait_until="domcontentloaded")

    if "accounts.google.com" in page.url:
        log("LOGIN", "目前在 Google 登入頁,請手動登入")
        await asyncio.to_thread(input, "登入完成後按 Enter 繼續... ")
        await page.goto(FORM_URL, wait_until="domcontentloaded")
    log("NAV", f"目前頁面: {page.url}")


async def click_next_or_submit(page: Page, selections: dict) -> str:
    global _ONCE_CONFIRMED
    submit_patterns = re.compile(r"提交|送出|傳送|Submit", re.I)
    next_patterns = re.compile(r"下一步|下一頁|繼續|Next|Continue", re.I)

    # 先走 ARIA role,失敗再走純 CSS 文字匹配。
    submit_btn = page.get_by_role("button", name=submit_patterns).first
    next_btn = page.get_by_role("button", name=next_patterns).first

    if await submit_btn.count() > 0 and await submit_btn.is_visible():
        if DRY_RUN:
            log("ACTION", "DRY_RUN=True,已到提交前,不送出")
            return "dry_run_ready"

        if should_confirm_before_submit():
            confirmed = await asyncio.to_thread(ask_submit_confirmation, selections)
            if not confirmed:
                log("ACTION", "使用者在提交前取消")
                return "aborted"
            _ONCE_CONFIRMED = True

        await delay_before_page_action_if_needed()
        await delay_before_submit_if_needed()
        log("ACTION", "偵測到提交按鈕,準備點擊")
        try:
            await submit_btn.click(timeout=10000)
            await page.wait_for_timeout(800)
            verify_result = await handle_robot_verification_if_needed(page)
            if verify_result == "aborted":
                return "aborted"
        except Exception as exc:
            log("ERROR", f"點擊提交失敗: {exc}")
            return "none"
        return "submitted"

    if await next_btn.count() > 0 and await next_btn.is_visible():
        log("ACTION", "偵測到下一步按鈕,準備點擊")
        try:
            await delay_before_page_action_if_needed()
            await next_btn.click(timeout=10000)
            await page.wait_for_load_state("networkidle")
        except Exception as exc:
            log("ERROR", f"點擊下一步失敗: {exc}")
            return "none"
        return "next"

    submit_fallback = page.locator('div[role="button"], button').filter(has_text=submit_patterns).first
    next_fallback = page.locator('div[role="button"], button').filter(has_text=next_patterns).first

    if await submit_fallback.count() > 0 and await submit_fallback.is_visible():
        if DRY_RUN:
            log("ACTION", "DRY_RUN=True,已到提交前(fallback),不送出")
            return "dry_run_ready"

        if should_confirm_before_submit():
            confirmed = await asyncio.to_thread(ask_submit_confirmation, selections)
            if not confirmed:
                log("ACTION", "使用者在提交前取消")
                return "aborted"
            _ONCE_CONFIRMED = True

        await delay_before_page_action_if_needed()
        await delay_before_submit_if_needed()
        log("ACTION", "使用 fallback 提交按鈕")
        try:
            await submit_fallback.click(timeout=10000)
            await page.wait_for_timeout(800)
            verify_result = await handle_robot_verification_if_needed(page)
            if verify_result == "aborted":
                return "aborted"
        except Exception as exc:
            log("ERROR", f"fallback 提交失敗: {exc}")
            return "none"
        return "submitted"

    if await next_fallback.count() > 0 and await next_fallback.is_visible():
        log("ACTION", "使用 fallback 下一步按鈕")
        try:
            await delay_before_page_action_if_needed()
            await next_fallback.click(timeout=10000)
            await page.wait_for_load_state("networkidle")
        except Exception as exc:
            log("ERROR", f"fallback 下一步失敗: {exc}")
            return "none"
        return "next"

    # 額外保險：有些 Google Form 的「提交」不是標準 button role。
    submit_text_fallback = page.get_by_text(re.compile(r"^\s*(提交|送出|傳送|Submit)\s*$", re.I)).first
    if await submit_text_fallback.count() > 0 and await submit_text_fallback.is_visible():
        if DRY_RUN:
            log("ACTION", "DRY_RUN=True,已到提交前(文字 fallback),不送出")
            return "dry_run_ready"

        if should_confirm_before_submit():
            confirmed = await asyncio.to_thread(ask_submit_confirmation, selections)
            if not confirmed:
                log("ACTION", "使用者在提交前取消")
                return "aborted"
            _ONCE_CONFIRMED = True

        await delay_before_page_action_if_needed()
        await delay_before_submit_if_needed()
        log("ACTION", "使用文字 fallback 提交")
        try:
            await submit_text_fallback.click(timeout=10000)
            await page.wait_for_timeout(800)
            verify_result = await handle_robot_verification_if_needed(page)
            if verify_result == "aborted":
                return "aborted"
        except Exception as exc:
            log("ERROR", f"文字 fallback 提交失敗: {exc}")
            return "none"
        return "submitted"

    visible_buttons = page.locator('div[role="button"], button')
    button_count = await visible_buttons.count()
    labels: list[str] = []
    hidden_labels: list[str] = []
    for idx in range(button_count):
        btn = visible_buttons.nth(idx)
        text = (await btn.inner_text()).strip()
        is_visible = await btn.is_visible()
        if text:
            if is_visible:
                labels.append(text)
            else:
                hidden_labels.append(text)
    if labels:
        log("DEBUG", f"可見按鈕: {' | '.join(labels[:10])}")
    else:
        log("DEBUG", "無可見按鈕")
    if hidden_labels:
        log("DEBUG", f"隱藏按鈕: {' | '.join(hidden_labels[:10])}")

    return "none"


async def fill_one_form(page: Page, selections: dict) -> str:
    """填寫一張表單,回傳 'submitted' / 'dry_run_ready' / 'aborted'"""
    round_idx = 1
    previous_signature = ""
    stagnant_rounds = 0
    empty_rounds = 0
    no_button_rounds = 0
    while True:
        log("LOOP", f"進入第 {round_idx} 輪")
        inspected, page_snapshot = await inspect_current_page(page)

        if not inspected:
            empty_rounds += 1
            if empty_rounds <= 3:
                log("WAIT", f"偵測到空頁,等待重試 ({empty_rounds}/3)")
                await page.wait_for_timeout(1200)
                round_idx += 1
                continue
        else:
            empty_rounds = 0

        current_signature = page_signature(page, inspected)
        if current_signature == previous_signature:
            stagnant_rounds += 1
        else:
            stagnant_rounds = 0
        previous_signature = current_signature

        selections, added = sync_selections_with_questions(selections, page_snapshot)
        if added:
            if AUTO_UPDATE_SELECTIONS:
                save_selections(selections)
                log("SYNC", "已更新 selections.json,請補齊以下題目答案")
            else:
                log("SYNC", "偵測到未映射欄位(未寫入 selections.json)")
            for label in added:
                log("SYNC", f"未映射欄位: {label}")

        await fill_questions_for_current_page(inspected, selections)

        mismatches = await verify_current_page_answers(inspected, selections)
        if mismatches:
            log("VERIFY", "送出前檢查發現選項狀態不一致，嘗試重新填寫")
            for mismatch in mismatches:
                log("VERIFY", mismatch)

            await fill_questions_for_current_page(inspected, selections)

            mismatches = await verify_current_page_answers(inspected, selections)
            if mismatches:
                log("VERIFY", "重新填寫後仍有未正確選取的欄位，暫停流程")
                for mismatch in mismatches:
                    log("VERIFY", mismatch)
                if PAUSE_ON_STALL:
                    user_choice = await asyncio.to_thread(
                        input,
                        "送出前檢查失敗。按 Enter 重試、輸入 q 結束: ",
                    )
                    if user_choice.strip().lower() == "q":
                        log("DONE", "使用者選擇結束")
                        return "aborted"
                    log("DONE", "使用者選擇重試")
                    round_idx += 1
                    continue
                return "aborted"

        action = await click_next_or_submit(page, selections)
        if action == "aborted":
            log("DONE", "使用者在驗證流程中選擇結束")
            return "aborted"
        if action == "dry_run_ready":
            log("DONE", "DRY_RUN 模式：已停在送出前")
            user_choice = await asyncio.to_thread(
                input,
                "請先手動按提交。完成後按 Enter 繼續下一筆 (q 結束批次): ",
            )
            if user_choice.strip().lower() == "q":
                return "aborted"
            return "dry_run_ready"
        if action == "submitted":
            log("DONE", "表單已送出")
            return "submitted"
        if action == "none":
            no_button_rounds += 1
            if no_button_rounds <= 3:
                log("WAIT", f"尚未偵測到下一步/提交按鈕,自動等待後重試 ({no_button_rounds}/3)")
                try:
                    await page.wait_for_load_state("domcontentloaded", timeout=2500)
                except Exception:
                    pass
                await page.wait_for_timeout(1500)
                round_idx += 1
                continue

            log("DONE", "未找到下一步或提交按鈕(表單可能需要手動填寫或已到達新頁面)")
            if PAUSE_ON_STALL:
                user_choice = await asyncio.to_thread(
                    input,
                    "流程已暫停。按 Enter 重試、輸入 q 結束: ",
                )
                if user_choice.strip().lower() == "q":
                    log("DONE", "使用者選擇結束")
                    return "aborted"
                log("DONE", "使用者選擇重試")
                round_idx += 1
                continue
            log("DONE", "PAUSE_ON_STALL=False,流程自動結束")
            return "aborted"
        else:
            no_button_rounds = 0

        if stagnant_rounds >= 3:
            missing_labels: list[str] = []
            for item in inspected:
                label = str(item.get("label", ""))
                if label and not get_answer_for_label(selections, label):
                    missing_labels.append(label)
            log("DONE", "頁面連續未前進(可能有必填項未完成)")
            if missing_labels:
                log("DONE", f"偵測到未填項目: {', '.join(missing_labels)}")
            if PAUSE_ON_STALL:
                user_choice = await asyncio.to_thread(
                    input,
                    "流程已暫停。按 Enter 重試、輸入 q 結束: ",
                )
                if user_choice.strip().lower() == "q":
                    log("DONE", "使用者選擇結束")
                    return "aborted"
                log("DONE", "使用者選擇重試")
                stagnant_rounds = 0
                round_idx += 1
                continue
            log("DONE", "PAUSE_ON_STALL=False,流程自動結束")
            return "aborted"

        round_idx += 1


async def run_filler() -> None:
    if FORM_URL == "你的GOOGLE表單連結":
        raise ValueError("請先把 FORM_URL 改成實際 Google 表單連結")

    base_selections = load_selections()
    
    # 優先嘗試從 CSV 讀取批次
    batch_with_rows = load_batch_from_csv()
    
    if batch_with_rows:
        # CSV 有結果,使用帶行號的批次
        batch = [entry for entry, _ in batch_with_rows]
        row_nums = [row_num for _, row_num in batch_with_rows]
    else:
        # CSV 無結果,回退到 selections.json 批次
        batch_raw = base_selections.pop(BATCH_KEY, None)
        if isinstance(batch_raw, list) and all(isinstance(x, list) for x in batch_raw):
            batch = batch_raw
            row_nums = [None] * len(batch)
        else:
            batch = [None]
            row_nums = [None]

    async with async_playwright() as p:
        context = await launch_persistent_context_with_fallback(p)
        page = context.pages[0] if context.pages else await context.new_page()
        await goto_form_and_wait_login(page)

        total = len(batch)
        completed = 0
        for idx, (course_entry, row_num) in enumerate(zip(batch, row_nums)):
            selections = dict(base_selections)
            if course_entry is not None:
                selections[COURSE_COMBO_KEY] = course_entry
                log("BATCH", f"第 {idx + 1}/{total} 筆: {course_entry}")
            else:
                log("START", "單筆模式")

            result = await fill_one_form(page, selections)

            if result == "aborted":
                log("BATCH", "使用者中止,停止批次")
                break

            # 表單成功送出或 DRY_RUN 完成後,更新 CSV 狀態
            if result in ("submitted", "dry_run_ready") and row_num is not None:
                update_csv_row_status(row_num, "D")

            completed += 1

            if idx < total - 1:
                await page.goto(FORM_URL, wait_until="domcontentloaded")
                await page.wait_for_timeout(800)

        log("BATCH", f"批次結束,共完成 {completed}/{total} 筆")
        await context.close()


async def main() -> None:
    try:
        await run_filler()
    except KeyboardInterrupt:
        log("EXIT", "使用者中斷執行")
    except Exception as exc:
        log("FATAL", f"未處理例外: {exc}")
        safe_print(traceback.format_exc())


if __name__ == "__main__":
    asyncio.run(main())