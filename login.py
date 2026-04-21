import asyncio
import csv
import json
import re
from pathlib import Path

from playwright.async_api import async_playwright

# 設定資料儲存路徑 (路徑可自訂)
USER_DATA_DIR = "./google_profile"
google_form_url = "https://docs.google.com/forms/d/e/1FAIpQLSchP7dRjEOyofx3V6cu7do8UM_WghZRuB9QnwMwQAvceZ2evg/viewform?pli=1&pli=1"
QUESTIONS_SNAPSHOT_FILE = Path("./questions_snapshot.json")
COURSE_CSV_FILE = Path("./courses_schedule.csv")
COURSE_TIME_KEY = "課程時間"
COURSE_TIME_LABELS = ["滑冰基礎班", "花式初級班", "花式進階班", "冰球初級班", "冰球進階班"]


def normalize_label(raw_text: str) -> str:
    return raw_text.split("\n")[0].replace("*", "").strip()


def is_email_like_label(text: str) -> bool:
    normalized = re.sub(r"\s+", "", text).strip().lower()
    return ("電子郵件" in normalized) or ("email" in normalized)


def load_selections() -> dict[str, str]:
    selections_file = Path("./selections.json")
    if not selections_file.exists():
        return {}
    try:
        data = json.loads(selections_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    if not isinstance(data, dict):
        return {}
    return {str(k): str(v) if v is not None else "" for k, v in data.items()}


def save_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def export_course_csv_from_questions(questions: list[dict[str, object]]) -> int:
    """輸出課程 CSV:狀態(空/V/D), 課程名, 時間。"""
    group_labels = {"滑冰基礎班", "花式課程", "冰球課程"}
    current_group: str | None = None
    rows: list[tuple[str, str, str]] = []
    seen: set[tuple[str, str]] = set()

    for item in questions:
        label = item.get("label")
        if not isinstance(label, str):
            continue

        if label in group_labels:
            current_group = label

        q_type = item.get("type")
        options = item.get("options")
        if q_type != 2 or not isinstance(options, list) or not options:
            continue
        if current_group is None:
            continue

        for opt in options:
            if not isinstance(opt, str):
                continue
            key = (current_group, opt)
            if key in seen:
                continue
            seen.add(key)
            rows.append(("", current_group, opt))

    with COURSE_CSV_FILE.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["狀態", "課程名", "時間"])
        writer.writerows(rows)

    return len(rows)


def normalize_course_time_key(selections: dict[str, str]) -> dict[str, str]:
    chosen = selections.get(COURSE_TIME_KEY, "").strip()
    for key in COURSE_TIME_LABELS:
        value = selections.get(key, "").strip()
        if not chosen and value:
            chosen = value

    for key in COURSE_TIME_LABELS:
        selections.pop(key, None)

    selections[COURSE_TIME_KEY] = chosen
    return selections


def debug_print_questions(questions: list[dict[str, object]]) -> None:
    print("=== 從 FB_PUBLIC_LOAD_DATA_ 解析的題目開始 ===")
    for q in questions:
        print(f"題目: {q.get('label')} | type: {q.get('type')} | entry: {q.get('entry')}")
        options = q.get("options")
        if isinstance(options, list):
            for opt in options:
                print(f"  option: {opt}")
    print("=== 從 FB_PUBLIC_LOAD_DATA_ 解析的題目結束 ===")


def extract_questions_from_fb_data(data: object) -> list[dict[str, object]]:
    if not isinstance(data, list) or len(data) < 2:
        raise ValueError("FB_PUBLIC_LOAD_DATA_ 格式不符")

    section = data[1]
    if not isinstance(section, list) or len(section) < 2 or not isinstance(section[1], list):
        raise ValueError("無法在 FB_PUBLIC_LOAD_DATA_ 找到題目陣列")

    questions: list[dict[str, object]] = []
    for idx, q in enumerate(section[1]):
        if not isinstance(q, list):
            continue

        label = q[1] if len(q) > 1 else None
        q_type = q[3] if len(q) > 3 else None
        entry = None
        options: list[str] = []

        if len(q) > 4 and isinstance(q[4], list) and q[4]:
            first = q[4][0]
            if isinstance(first, list):
                if first:
                    entry = first[0]
                if len(first) > 1 and isinstance(first[1], list):
                    for opt in first[1]:
                        if isinstance(opt, list) and opt:
                            text = opt[0]
                            if text is not None:
                                options.append(str(text))

        normalized_label = normalize_label(str(label)) if label is not None else None
        questions.append({
            "index": idx,
            "label": normalized_label,
            "type": q_type,
            "entry": entry,
            "options": options,
        })

    return questions


async def read_fb_public_load_data(page) -> object:
    # 優先用 JS 直接抓，抓不到再退回 HTML regex。
    data = await page.evaluate("() => window.FB_PUBLIC_LOAD_DATA_ || null")
    if data is not None:
        return data

    html = await page.content()
    match = re.search(r"FB_PUBLIC_LOAD_DATA_\\s*=\\s*(\\[.*?\\]);", html, re.DOTALL)
    if not match:
        raise ValueError("頁面中找不到 FB_PUBLIC_LOAD_DATA_")

    return json.loads(match.group(1))


async def detect_email_field_label(page) -> tuple[str | None, list[str]]:
    email_input = page.locator('input[type="email"], input[autocomplete="email"]')
    if await email_input.count() == 0:
        return None, []

    first = email_input.first
    candidates: list[str | None] = [
        await first.get_attribute("aria-label"),
        await first.get_attribute("placeholder"),
        await first.get_attribute("name"),
    ]

    labelled_text = await first.evaluate(
        """
        (el) => {
            const ids = (el.getAttribute('aria-labelledby') || '').trim();
            if (!ids) return null;
            const nodes = ids.split(/\s+/).map(id => document.getElementById(id)).filter(Boolean);
            const text = nodes.map(n => (n.textContent || '').trim()).filter(Boolean).join(' ').trim();
            return text || null;
        }
        """
    )
    candidates.append(labelled_text)

    nearby_text = await first.evaluate(
        """
        (el) => {
            const clean = (s) => (s || '').replace(/\s+/g, ' ').trim();
            const roots = [el.closest('[role="listitem"]'), el.closest('form'), el.parentElement].filter(Boolean);
            for (const root of roots) {
                const heading = root.querySelector('[role="heading"]');
                if (heading) {
                    const t = clean(heading.textContent);
                    if (/電子郵件|email/i.test(t)) return t;
                }
                const nodes = Array.from(root.querySelectorAll('div, span, label')).slice(0, 120);
                for (const n of nodes) {
                    const t = clean(n.textContent);
                    if (/電子郵件|email/i.test(t)) return t;
                }
            }
            return null;
        }
        """
    )
    candidates.append(nearby_text)

    normalized_candidates: list[str] = []
    for raw in candidates:
        if raw and raw.strip():
            normalized = normalize_label(raw.strip())
            if normalized:
                normalized_candidates.append(normalized)

    for text in normalized_candidates:
        if is_email_like_label(text):
            return text, normalized_candidates

    # 不使用固定預設鍵名，只回傳真實偵測結果。
    return None, normalized_candidates

async def init_browser():
    async with async_playwright() as p:
        # 開啟持久化上下文
        context = await p.chromium.launch_persistent_context(
            USER_DATA_DIR,
            headless=False, # 必須為 False 才能手動登入
            args=["--disable-blink-features=AutomationControlled"] # 隱藏自動化特徵，減少驗證碼
        )
        page = context.pages[0] if context.pages else await context.new_page()
        await page.goto(google_form_url, wait_until="domcontentloaded")

        if "accounts.google.com" in page.url:
            print("請先在瀏覽器完成 Google 登入，完成後按 Enter。")
            await asyncio.to_thread(input, "登入完成後按 Enter 繼續... ")
            await page.goto(google_form_url, wait_until="domcontentloaded")

        await page.wait_for_load_state("networkidle")
        fb_data = await read_fb_public_load_data(page)
        print(f"FB_PUBLIC_LOAD_DATA_ 型別: {type(fb_data).__name__}")
        if isinstance(fb_data, list):
            print(f"FB_PUBLIC_LOAD_DATA_ 第一層長度: {len(fb_data)}")
        questions = extract_questions_from_fb_data(fb_data)
        print(f"總題數: {len(questions)}")
        debug_print_questions(questions)

        selections = load_selections()
        selections = normalize_course_time_key(selections)

        for item in questions:
            label = item.get("label")
            entry = item.get("entry")
            if isinstance(label, str) and label and entry is not None and label not in selections:
                selections[label] = ""

        # Google Form 的「收集電子郵件」常不在 FB_PUBLIC_LOAD_DATA_，改用 DOM 補抓。
        email_label, email_candidates = await detect_email_field_label(page)
        print("Email 候選字串:", email_candidates)
        print("Email 最終偵測結果:", email_label)
        if email_label and email_label not in selections:
            selections[email_label] = ""

        snapshot_data = {
            "source": "FB_PUBLIC_LOAD_DATA_",
            "google_form_url": google_form_url,
            "questions": questions,
            "email_label": email_label,
        }

        save_json(QUESTIONS_SNAPSHOT_FILE, snapshot_data)
        save_json(Path("./selections.json"), selections)
        csv_count = export_course_csv_from_questions(questions)
        print("已建立/更新 questions_snapshot.json 與 selections.json")
        print(f"已建立/更新 {COURSE_CSV_FILE.name}，共 {csv_count} 筆課程時間")
        print("selections 鍵名:", list(selections.keys()))

        await asyncio.to_thread(input, "請檢查 JSON 後按 Enter 結束... ")
        await context.close()

# 第一次執行先執行這個來登入
if __name__ == "__main__":
    asyncio.run(init_browser())