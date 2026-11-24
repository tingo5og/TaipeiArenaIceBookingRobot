"""台北小巨蛋冰上樂園學員系統 (TAPICELand) 的共用瀏覽器操作。"""

import asyncio
import csv
import os
import sys
import time
from contextlib import asynccontextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import AsyncIterator

from playwright.async_api import BrowserContext, Page


def get_base_dir() -> Path:
    from_env = os.environ.get("BOOKING_BASE_DIR", "").strip()
    if from_env:
        return Path(from_env).resolve()
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()

BROWSER_PROFILE_DIR = BASE_DIR / "browser_profile"
FALLBACK_PROFILE_DIR = BASE_DIR / "browser_profile_fallback"
COURSE_CSV_FILE = BASE_DIR / "courses_schedule.csv"

SITE_ROOT = "https://ssl.metro.taipei/TAPICELand/Student/"
LOGIN_URL = SITE_ROOT + "Login.aspx"
CLASS_USE_INFO_URL = SITE_ROOT + "ClassUseInfo.aspx"
RESER_GROUP_CARD_URL = SITE_ROOT + "ReserGroupCard.aspx"

# GUI「登入」按鈕開啟的瀏覽器會帶這個遠端除錯埠，掃描/預約腳本透過它沿用同一個登入狀態。
DEBUG_PORT = int(os.environ.get("BOOKING_DEBUG_PORT", "9333"))
CDP_URL = f"http://127.0.0.1:{DEBUG_PORT}"

LOGIN_WAIT_SEC = int(os.environ.get("BOOKING_LOGIN_WAIT_SEC", "600"))
STATUS_BOOKABLE = "可預約"
# 上課列表中已報名的課程，按鈕會顯示「取消預約」。
STATUS_BOOKED = "取消預約"


def configure_stdio_utf8() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


configure_stdio_utf8()


def safe_print(message: str) -> None:
    try:
        print(message, flush=True)
        return
    except UnicodeEncodeError:
        pass
    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    sys.stdout.write(message.encode(encoding, errors="replace").decode(encoding, errors="replace") + "\n")
    sys.stdout.flush()


def log(step: str, message: str) -> None:
    safe_print(f"[{time.strftime('%H:%M:%S')}] [{step}] {message}")


def write_csv_atomic(path: Path, rows: list[list[str]]) -> None:
    """先寫暫存檔再替換，避免 GUI 輪詢時讀到寫一半的 CSV。"""
    tmp_path = path.with_name(path.name + ".tmp")
    with tmp_path.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    # Windows 上若檔案正被讀取，替換會暫時失敗，稍後重試。
    for _ in range(50):
        try:
            os.replace(tmp_path, path)
            return
        except PermissionError:
            time.sleep(0.1)
    # 仍無法替換時直接覆寫，確保結果有寫回（讀取端已能容忍讀到不完整的檔案）。
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    try:
        tmp_path.unlink()
    except OSError:
        pass


@dataclass
class Card:
    label: str
    code: str
    purchase_date: str
    expiry_date: str
    bookable: bool


@dataclass
class CourseSlot:
    date: str
    weekday: str
    time_range: str
    course_name: str
    coach: str
    status: str
    code: str

    @property
    def time_slot(self) -> str:
        """CSV「時間」欄的格式，例如 2026/11/26(四)｜17:00~18:00｜林肇藩。"""
        return f"{self.date}{self.weekday}｜{self.time_range}｜{self.coach}"

    @property
    def is_registered(self) -> bool:
        return self.status == STATUS_BOOKED


def cleanup_profile_lock_files(profile_dir: Path) -> None:
    removed: list[str] = []
    for lock_path in [
        profile_dir / "SingletonLock",
        profile_dir / "SingletonCookie",
        profile_dir / "SingletonSocket",
        profile_dir / "lockfile",
        profile_dir / "Default" / "LOCK",
    ]:
        try:
            if lock_path.exists():
                lock_path.unlink()
                removed.append(str(lock_path.relative_to(profile_dir)))
        except Exception as exc:
            log("WARN", f"無法移除 lock 檔 {lock_path.name}: {exc}")
    if removed:
        log("BROWSER", f"已清理 lock 檔: {', '.join(removed)}")


async def launch_context(playwright_instance, slow_mo: int = 0) -> BrowserContext:
    """找不到登入用瀏覽器時的備援：自行啟動瀏覽器；主 profile 被占用時改用備援 profile。"""
    launch_kwargs = {
        "headless": False,
        "slow_mo": slow_mo,
        "args": ["--disable-blink-features=AutomationControlled"],
    }
    for profile_dir in [BROWSER_PROFILE_DIR, FALLBACK_PROFILE_DIR]:
        profile_dir.mkdir(parents=True, exist_ok=True)
        cleanup_profile_lock_files(profile_dir)
        try:
            context = await playwright_instance.chromium.launch_persistent_context(str(profile_dir), **launch_kwargs)
            log("BROWSER", f"使用 profile 啟動: {profile_dir}")
            return context
        except Exception as exc:
            log("WARN", f"profile 啟動失敗 ({profile_dir}): {exc}")
    raise RuntimeError("無法啟動瀏覽器")


@asynccontextmanager
async def open_site_page(playwright_instance, slow_mo: int = 0) -> AsyncIterator[Page]:
    """優先接上「登入」按鈕開啟的瀏覽器並開新分頁；結束時只關分頁，保留登入狀態。

    找不到該瀏覽器時，改為自行啟動瀏覽器（需在該視窗登入，結束後會關閉）。
    """
    try:
        browser = await playwright_instance.chromium.connect_over_cdp(CDP_URL, timeout=5000, slow_mo=slow_mo)
    except Exception:
        browser = None

    if browser is not None and browser.contexts:
        log("BROWSER", "已連上登入用瀏覽器")
        page = await browser.contexts[0].new_page()
        try:
            yield page
        finally:
            try:
                await page.close()
            except Exception:
                pass
        return

    log("BROWSER", "找不到登入用瀏覽器（請先按「1. 登入」），改為自行開啟瀏覽器")
    context = await launch_context(playwright_instance, slow_mo=slow_mo)
    page = context.pages[0] if context.pages else await context.new_page()
    try:
        yield page
    finally:
        await context.close()


def is_login_page(page: Page) -> bool:
    return "login.aspx" in page.url.lower()


async def has_logout_link(page: Page) -> bool:
    try:
        return await page.locator('a[href*="logoutFunction"]').count() > 0
    except Exception:
        return False


async def ensure_logged_in(page: Page) -> None:
    """確認已登入；未登入時開啟登入頁，等待使用者手動輸入帳密、驗證碼並登入。"""
    await page.goto(CLASS_USE_INFO_URL, wait_until="domcontentloaded", timeout=60000)
    if await has_logout_link(page):
        log("LOGIN", "已是登入狀態")
        return

    if not is_login_page(page):
        await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    log("LOGIN", "尚未登入：請在瀏覽器輸入帳號、密碼、驗證碼並按「登入」，完成後程式會自動繼續")
    deadline = time.monotonic() + LOGIN_WAIT_SEC
    last_notice = time.monotonic()
    while not (not is_login_page(page) and await has_logout_link(page)):
        if time.monotonic() > deadline:
            raise TimeoutError(f"等待登入逾時 ({LOGIN_WAIT_SEC} 秒)")
        if time.monotonic() - last_notice >= 30:
            log("LOGIN", "仍在等待登入...")
            last_notice = time.monotonic()
        await asyncio.sleep(1)
    log("LOGIN", "登入完成")


async def list_cards(page: Page) -> list[Card]:
    """在 ClassUseInfo.aspx 將卡片種類選最上方（全部）並查詢，回傳所有卡片。"""
    await page.goto(CLASS_USE_INFO_URL, wait_until="domcontentloaded", timeout=60000)
    # 未登入時部分頁面會導向錯誤頁而不是登入頁，因此以「登出」連結判斷。
    if not await has_logout_link(page):
        raise RuntimeError("登入已失效，請重新按「1. 登入」")

    card_type = page.locator("#ContentPlaceHolder1_CardType")
    await card_type.select_option(index=0)
    # 剛開啟時表格是空的，按「查詢」後才會局部更新出卡片列（不會換頁）。
    await page.locator("#ContentPlaceHolder1_btnQuery").click()
    try:
        await page.locator(".booking").first.wait_for(state="visible", timeout=15000)
    except Exception:
        log("CARD", "查詢後沒有出現任何「預約」按鈕")
    await page.wait_for_load_state("domcontentloaded")

    raw = await page.evaluate(
        """() => Array.from(document.querySelectorAll('table tbody tr'))
            .map(tr => ({
                cells: Array.from(tr.children).map(td => td.innerText.trim()),
                bookable: !!tr.querySelector('.booking'),
            }))
            .filter(r => r.cells.length >= 3 && r.cells[0].includes('-'))"""
    )
    return [
        Card(
            label=row["cells"][0],
            code=row["cells"][0].split("-")[-1].strip(),
            purchase_date=row["cells"][1],
            expiry_date=row["cells"][2],
            bookable=row["bookable"],
        )
        for row in raw
    ]


def pick_card(cards: list[Card]) -> Card | None:
    """有多張可預約卡片時，選到期日期最晚的（同到期日再比購買日期）。"""
    bookable = [c for c in cards if c.bookable]
    if not bookable:
        return None
    # 日期格式為 YYYY/MM/DD，可直接用字串比較。
    return max(bookable, key=lambda c: (c.expiry_date, c.purchase_date))


async def open_group_booking_page(page: Page) -> Card:
    """選擇卡片並進入團體課程預約頁 (ReserGroupCard.aspx)。"""
    target: Card | None = None
    # 「預約」按鈕在 <form> 內，點擊時 AJAX 跳轉與表單 postback 會互相競爭；沒跳轉成功就重試。
    for attempt in range(1, 4):
        cards = await list_cards(page)
        target = pick_card(cards)
        if target is None:
            raise RuntimeError("找不到可預約的卡片（課程預約頁沒有「預約」按鈕）")
        if attempt == 1:
            others = [c.code for c in cards if c.bookable and c is not target]
            log("CARD", f"使用卡片: {target.label}（到期 {target.expiry_date}）" + (f"，其他可預約卡片: {others}" if others else ""))

        row = page.locator("table tbody tr").filter(has_text=target.code).first
        try:
            async with page.expect_navigation(url="**/ReserGroupCard.aspx", timeout=15000):
                await row.locator(".booking").first.click()
            break
        except Exception:
            if "resergroupcard.aspx" in page.url.lower():
                break
            log("CARD", f"未進入課程預約頁，重試 ({attempt}/3)")
            if attempt == 3:
                raise RuntimeError("無法進入課程預約頁 (ReserGroupCard.aspx)")
    await wait_for_course_table(page)
    assert target is not None
    return target


async def wait_for_course_table(page: Page) -> None:
    await page.wait_for_load_state("domcontentloaded")
    try:
        await page.locator(".btn-booking, .btn-cancel").first.wait_for(state="attached", timeout=10000)
    except Exception:
        log("WARN", "預約頁沒有任何可操作的課程按鈕，可能尚未開放或已額滿")


async def read_course_slots(page: Page) -> list[CourseSlot]:
    """解析 ReserGroupCard.aspx 的課程表。欄位: 日期,星期,時間,課程,教練,按鈕,假日旗標(隱藏),課程代碼(隱藏)。"""
    rows = await page.evaluate(
        """() => Array.from(document.querySelectorAll('table tbody tr'))
            .map(tr => Array.from(tr.children).map(td => td.textContent.trim()))
            .filter(cells => cells.length >= 8 && /^\\d{4}\\/\\d{2}\\/\\d{2}$/.test(cells[0]))"""
    )
    return [
        CourseSlot(
            date=cells[0],
            weekday=cells[1],
            time_range=cells[2],
            course_name=cells[3],
            coach=cells[4],
            status=cells[5],
            code=cells[-1],
        )
        for cells in rows
    ]
