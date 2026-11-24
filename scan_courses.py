"""掃描小巨蛋學員系統的團體課程表，更新 courses_schedule.csv。"""

import asyncio
import csv

from playwright.async_api import async_playwright

from metro_site import (
    COURSE_CSV_FILE,
    CourseSlot,
    ensure_logged_in,
    log,
    open_group_booking_page,
    open_site_page,
    read_course_slots,
    write_csv_atomic,
)


def read_existing_statuses() -> dict[tuple[str, str], str]:
    statuses: dict[tuple[str, str], str] = {}
    if not COURSE_CSV_FILE.exists():
        return statuses
    with COURSE_CSV_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            key = (str(row.get("課程名", "")).strip(), str(row.get("時間", "")).strip())
            statuses[key] = str(row.get("狀態", "")).strip()
    return statuses


def export_course_csv(slots: list[CourseSlot]) -> int:
    """輸出 CSV: 狀態(空/V/D/X), 課程名, 時間。

    報名狀態以上課列表為準：按鈕為「取消預約」標成 D(已報名)，若原本已勾選待取消(X)則保留 X；
    原本是 D/X 但已不是報名狀態（例如已取消）則清空。其他課程保留舊檔中的勾選狀態(V)。
    """
    old_statuses = read_existing_statuses()
    rows: list[list[str]] = [["狀態", "課程名", "時間"]]
    for slot in slots:
        status = old_statuses.get((slot.course_name, slot.time_slot), "")
        if slot.is_registered:
            status = "X" if status.upper() == "X" else "D"
        elif status.upper() in ("D", "X"):
            status = ""
        rows.append([status, slot.course_name, slot.time_slot])
    write_csv_atomic(COURSE_CSV_FILE, rows)
    return len(slots)


async def scan_courses() -> None:
    async with async_playwright() as p:
        async with open_site_page(p) as page:
            await ensure_logged_in(page)
            await open_group_booking_page(page)

            slots = await read_course_slots(page)
            count = export_course_csv(slots)
            log("DONE", f"已更新 {COURSE_CSV_FILE.name}，共 {count} 筆課程")

            registered = [s for s in slots if s.is_registered]
            log("REGISTERED", f"已報名 {len(registered)} 堂")
            for slot in registered:
                log("REGISTERED", f"{slot.course_name} | {slot.time_slot}")


if __name__ == "__main__":
    asyncio.run(scan_courses())
