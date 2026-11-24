"""依 courses_schedule.csv 批次預約(狀態=V)或取消預約(狀態=X)小巨蛋團體課程，全自動執行。"""

import asyncio
import csv
import os
import traceback

from playwright.async_api import Page, async_playwright

from metro_site import (
    COURSE_CSV_FILE,
    RESER_GROUP_CARD_URL,
    STATUS_BOOKABLE,
    CourseSlot,
    ensure_logged_in,
    has_logout_link,
    log,
    open_group_booking_page,
    open_site_page,
    read_course_slots,
    safe_print,
    wait_for_course_table,
    write_csv_atomic,
)

DRY_RUN = False  # True 時會開到網站的確認視窗後按「取消」，不會真的預約或取消預約。

RESULT_WAIT_MS = 30000

# book: 預約狀態=V 的課程；cancel: 取消狀態=X 的已報名課程。
ACTION = os.environ.get("BOOKING_ACTION", "book").strip().lower()
ACTION_BOOK = "book"
ACTION_CANCEL = "cancel"
ACTION_LABELS = {ACTION_BOOK: "預約", ACTION_CANCEL: "取消預約"}
ACTION_TARGET_STATUS = {ACTION_BOOK: "V", ACTION_CANCEL: "X"}


def read_csv_rows() -> list[list[str]]:
    with COURSE_CSV_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f))


def load_batch_from_csv(target_status: str = "V", action_label: str = "預約") -> list[tuple[str, str, int]]:
    """回傳 [(課程名, 時間, CSV 列索引), ...]，只包含狀態=target_status 的列。"""
    if not COURSE_CSV_FILE.exists():
        log("CSV", f"{COURSE_CSV_FILE.name} 不存在，請先「掃描課程列表」")
        return []
    rows = read_csv_rows()
    if not rows or rows[0][:1] != ["狀態"]:
        log("CSV", "CSV 格式不正確(缺少標頭或首欄不是狀態)")
        return []
    batch = [
        (row[1].strip(), row[2].strip(), idx)
        for idx, row in enumerate(rows)
        if idx > 0 and len(row) >= 3 and row[0].strip().upper() == target_status
    ]
    log("CSV", f"已從 CSV 讀取 {len(batch)} 筆待{action_label}課程")
    return batch


def update_csv_row_status(row_index: int, status: str) -> None:
    try:
        rows = read_csv_rows()
        if 0 < row_index < len(rows) and rows[row_index]:
            rows[row_index][0] = status
            write_csv_atomic(COURSE_CSV_FILE, rows)
            log("CSV", f"已更新第 {row_index + 1} 行狀態為: {status}")
    except Exception as exc:
        log("ERROR", f"更新 CSV 失敗: {exc}")


async def swal_text(page: Page) -> str:
    container = page.locator("#swal2-html-container")
    if await container.count() and await container.is_visible():
        return (await container.inner_text()).strip()
    return ""


async def open_site_dialog(page: Page, slot: CourseSlot, button_selector: str) -> str:
    """點課程列上的按鈕並等網站確認視窗出現，回傳視窗文字。

    按鈕在 <form> 內，postback 可能讓確認視窗消失；按「確認」前不會送出，重試是安全的。
    """
    confirm_btn = page.locator(".swal2-confirm")
    for attempt in range(1, 4):
        row = page.locator("table tbody tr").filter(has=page.get_by_text(slot.code, exact=True))
        await row.locator(button_selector).first.click()
        try:
            await confirm_btn.wait_for(state="visible", timeout=8000)
            break
        except Exception:
            if attempt == 3:
                raise
            log("SITE", f"未出現確認視窗，重試 ({attempt}/3)")
            await wait_for_course_table(page)
    text = await swal_text(page)
    log("SITE", text)
    return text


async def book_slot(page: Page, slot: CourseSlot) -> tuple[str, str]:
    """點「可預約」→ 網站確認視窗 → 讀取結果。回傳 (結果, 訊息)，結果為 booked/failed/dry_run。"""
    await open_site_dialog(page, slot, ".btn-booking")
    confirm_btn = page.locator(".swal2-confirm")

    if DRY_RUN:
        await page.locator(".swal2-cancel").click()
        return "dry_run", "DRY_RUN 模式，已按取消"

    await confirm_btn.click()
    # 網站確認後會以 AJAX 送出，再以 success/error 圖示的視窗顯示結果。
    await page.wait_for_function(
        """() => document.querySelector(
            '.swal2-popup.swal2-show .swal2-success.swal2-icon-show, .swal2-popup.swal2-show .swal2-error.swal2-icon-show')""",
        timeout=RESULT_WAIT_MS,
    )
    success = await page.locator(".swal2-popup.swal2-show .swal2-success.swal2-icon-show").count() > 0
    message = await swal_text(page)
    await confirm_btn.click()

    if success:
        # 學習卡(MA)會再問「是否繼續預約」，按確認會回到預約頁。
        try:
            await page.locator("#swal2-title", has_text="是否繼續預約").wait_for(state="visible", timeout=3000)
            await confirm_btn.click()
        except Exception:
            pass
    return ("booked" if success else "failed"), message


async def cancel_slot(page: Page, slot: CourseSlot) -> tuple[str, str]:
    """點「取消預約」→ 網站確認視窗 → 讀取結果訊息。回傳 (結果, 訊息)，結果為 sent/dry_run。

    網站的取消結果視窗一律顯示成功圖示，因此是否真的取消，需回到上課列表再確認。
    """
    confirm_text = await open_site_dialog(page, slot, ".btn-cancel")
    confirm_btn = page.locator(".swal2-confirm")

    if DRY_RUN:
        await page.locator(".swal2-cancel").click()
        return "dry_run", "DRY_RUN 模式，已按取消"

    await confirm_btn.click()
    await page.wait_for_function(
        """(confirmText) => {
            const popup = document.querySelector('.swal2-popup.swal2-show');
            const text = document.querySelector('#swal2-html-container');
            return popup && text && text.innerText.trim() && text.innerText.trim() !== confirmText;
        }""",
        arg=confirm_text,
        timeout=RESULT_WAIT_MS,
    )
    message = await swal_text(page)
    # 按確認後網站會跳到「查詢預約」頁。
    await confirm_btn.click()
    await page.wait_for_load_state("domcontentloaded")
    return "sent", message


async def back_to_booking_page(page: Page) -> None:
    await page.goto(RESER_GROUP_CARD_URL, wait_until="domcontentloaded", timeout=60000)
    if not await has_logout_link(page):
        await ensure_logged_in(page)
        await open_group_booking_page(page)
        return
    if not await page.locator("#ContentPlaceHolder1_btnQuery").count():
        await open_group_booking_page(page)
        return
    await wait_for_course_table(page)


async def find_slot(page: Page, course_name: str, time_slot: str) -> CourseSlot | None:
    slots = await read_course_slots(page)
    return next((s for s in slots if s.course_name == course_name and s.time_slot == time_slot), None)


async def process_book(page: Page, slot: CourseSlot, row_index: int) -> str:
    """回傳 done/skip。"""
    if slot.is_registered:
        log("SKIP", "這堂課已報名，標記為 D")
        update_csv_row_status(row_index, "D")
        return "skip"
    if slot.status != STATUS_BOOKABLE:
        log("SKIP", f"目前狀態為「{slot.status}」，無法預約")
        return "skip"
    result, message = await book_slot(page, slot)
    log("RESULT", f"{result}: {message}")
    if result == "booked":
        update_csv_row_status(row_index, "D")
        return "done"
    return "skip"


async def process_cancel(page: Page, slot: CourseSlot, row_index: int) -> str:
    """回傳 done/skip。"""
    if not slot.is_registered:
        log("SKIP", "這堂課目前不是已報名狀態，不需取消，清除標記")
        update_csv_row_status(row_index, "")
        return "skip"
    result, message = await cancel_slot(page, slot)
    log("RESULT", f"{result}: {message}")
    if result == "dry_run":
        return "skip"

    await back_to_booking_page(page)
    after = await find_slot(page, slot.course_name, slot.time_slot)
    if after is not None and after.is_registered:
        log("RESULT", "上課列表仍顯示已報名，取消未成功（可能已超過最晚取消日期）")
        return "skip"
    log("RESULT", "已確認取消成功")
    update_csv_row_status(row_index, "")
    return "done"


async def run_booking() -> None:
    action = ACTION if ACTION in ACTION_LABELS else ACTION_BOOK
    action_label = ACTION_LABELS[action]
    batch = load_batch_from_csv(ACTION_TARGET_STATUS[action], action_label)
    if not batch:
        return

    async with async_playwright() as p:
        async with open_site_page(p, slow_mo=200) as page:
            await ensure_logged_in(page)
            await open_group_booking_page(page)

            total = len(batch)
            done = 0
            for idx, (course_name, time_slot, row_index) in enumerate(batch):
                log("BATCH", f"第 {idx + 1}/{total} 筆: {course_name} | {time_slot}")
                if idx > 0:
                    await back_to_booking_page(page)

                slot = await find_slot(page, course_name, time_slot)
                if slot is None:
                    log("SKIP", "上課列表找不到這堂課（可能尚未開放、已過期或資料已變更）")
                    continue

                try:
                    if action == ACTION_CANCEL:
                        outcome = await process_cancel(page, slot, row_index)
                    else:
                        outcome = await process_book(page, slot, row_index)
                except Exception as exc:
                    log("ERROR", f"{action_label}過程發生錯誤: {exc}")
                    continue

                if outcome == "done":
                    done += 1

            log("BATCH", f"批次結束，共成功{action_label} {done}/{total} 筆")


async def main() -> None:
    try:
        await run_booking()
    except KeyboardInterrupt:
        log("EXIT", "使用者中斷執行")
    except Exception as exc:
        log("FATAL", f"未處理例外: {exc}")
        safe_print(traceback.format_exc())


if __name__ == "__main__":
    asyncio.run(main())
