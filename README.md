# Google Form 批次填表工具

此專案用於自動填寫 Google 表單，支援：
- 先讀取表單題目並產生設定檔與課程時段 CSV
- 以 CSV `狀態=V` 進行批次填寫
- 每筆完成後自動把 CSV 狀態更新為 `D`
- 提交後若遇到人機驗證（reCAPTCHA），自動暫停等待手動通過

## 檔案說明

- `login.py`
  - 讀取 Google Form 題目
  - 輸出 `questions_snapshot.json`
  - 建立/更新 `selections.json`
  - 輸出 `courses_schedule.csv`

- `do_table.py`
  - 讀取 `selections.json` 與 `courses_schedule.csv`
  - 自動填寫表單（可批次）
  - 批次時會讀取 CSV 中 `狀態=V` 的列
  - 每筆完成後將對應列狀態改為 `D`

- `selections.json`
  - 一般題目的答案設定（姓名、生日、Email 等）

- `courses_schedule.csv`
  - 欄位：`狀態,課程名,時間`
  - 狀態規則：
    - 空白：未排程
    - `V`：待填寫
    - `D`：已完成

## 環境需求

- Python 3.11+
- Windows（目前使用情境）
- Google 帳號可登入

## 安裝

```bash
pip install playwright
playwright install chromium
```

## 使用流程

### 1) 先抓表單結構

```bash
python login.py
```

執行後：
- 若跳 Google 登入頁，請手動登入
- 完成後回到程式繼續
- 會產生/更新 `selections.json`、`questions_snapshot.json`、`courses_schedule.csv`

### 2) 編輯答案與批次清單

1. 開啟 `selections.json`，填好固定欄位（例如姓名、生日、電子郵件、電話）。
2. 開啟 `courses_schedule.csv`，把想報名的時段狀態改成 `V`。

### 3) 執行批次填寫

```bash
python do_table.py
```

程式行為：
- 逐筆處理 CSV 裡的 `V`
- 成功送出（或 DRY_RUN 手動確認完成）後，該列自動改成 `D`
- 若遇到人機驗證，會停下並提示你先手動通過

## 重要參數（`do_table.py`）

- `DRY_RUN = False`
  - `True`：到提交前停下，不自動按提交
  - `False`：自動提交

- `PAUSE_ON_STALL = True`
  - 找不到下一步/提交或流程停滯時，暫停等待手動決策

- `FORM_URL`
  - 目標 Google Form 網址

- `USER_DATA_DIR = "./google_profile"`
  - 瀏覽器持久化資料夾，用來保留登入狀態

## 人機驗證處理

提交後如果出現 reCAPTCHA 或「不是機器人」頁面：
1. 程式會顯示偵測訊息並暫停。
2. 你在瀏覽器手動完成驗證。
3. 回到終端按 Enter 讓程式繼續。
4. 輸入 `q` 可中止整個批次。

## 常見問題

- 沒讀到批次資料
  - 檢查 `courses_schedule.csv` 是否有 `狀態=V` 的列。

- 一直要求登入
  - 確認 `google_profile` 目錄可寫入，並且沒有被清空。

- 時段對不到
  - `courses_schedule.csv` 的 `時間` 欄位請直接使用表單原字串。

## 建議操作

1. 先用 `DRY_RUN=True` 測一次流程。
2. 確認填寫正確後改回 `DRY_RUN=False` 正式提交。
3. 每次批次前先確認 CSV 的 `V` 只有你要送出的項目。
