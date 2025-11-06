AUTOA-RPA — 以影像辨識與鍵鼠模擬自動化 LINE 桌面版

README（初版 v0.1.0） — 專案說明與上手指南
⚠️ 重要聲明
本專案僅供內部研究與個人自動化用途。請先取得收件者明確同意、遵守 LINE 與作業系統/應用程式之服務條款與各地法規。請避免濫發訊息、詐騙或侵害隱私之用途。你需自行承擔使用風險（封鎖、凍結、法務風險等）。

1. 專案目標

在 不使用官方 Messaging API 的前提下，以 OpenCV 影像辨識 + PyAutoGUI/pywinauto 鍵鼠模擬 自動化 LINE 桌面版的相簿/圖片批量發送與相關操作（新增/刪除/清空、切換相簿、分批發送、節流與排程）。

2. 功能概覽（MVP）

📸 相簿/圖片發送流程自動化：開啟聊天→插入相簿/圖片→輸入標題→送出

🧭 UI 模板定位：多解析度模板比對（載入 templates/ PNG 作為定位錨點）

🖱️ 鍵鼠驅動：點擊、拖放、輸入文字、快捷鍵、滾動

🧱 任務 DSL（YAML）：以 YAML 描述一個「流程腳本」與參數

🛡️ 節流/人類化：隨機延遲、每輪上限、黑名單、退訂名單

♻️ 容錯恢復：逾時重試、失敗截圖、可從指定步驟續跑

🗓️ 排程：Windows 排程器（或內建 APScheduler）

🧾 日誌與證據：logs/、reports/、（可選）螢幕錄影

3. 系統需求

OS：Windows 10/11（桌面版 LINE）

顯示：建議 1920×1080、DPI 縮放 100%（不一致會影響模板比對）

Python：3.10+（建議 3.11）

硬體：建議 RAM 8GB+

權限：建議一般權限即可，若 UAC 阻擋再以系統管理員啟動

依賴套件（會在 requirements.txt 統一安裝）

pyautogui, opencv-python, numpy, pillow

pywinauto（可選，做視窗聚焦/容錯）

pytesseract（可選，OCR 校驗文字）

apscheduler, python-dotenv, pyyaml, loguru

4. 安裝步驟
# 1) 取得專案
git clone https://your.repo/AUTOA-RPA.git
cd AUTOA-RPA

# 2) 建立虛擬環境並安裝依賴
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# 3) 建立設定檔
copy .env.example .env
copy config.example.yaml config.yaml

# 4) 放入 UI 模板
# 將你的 UI 錨點圖片放到 templates/（見第 6 節指引）

5. 專案結構
AUTOA-RPA/
├─ autoa/                   # 核心程式
│  ├─ __init__.py
│  ├─ rpa.py                # 點擊/移動/拖放/鍵盤/熱鍵
│  ├─ vision.py             # OpenCV 模板比對、區域搜尋
│  ├─ flow.py               # YAML 任務 DSL 解析與執行器
│  ├─ guards.py             # 節流、人類化延遲、黑名單
│  ├─ ocr.py                # （可選）Tesseract OCR 介面
│  ├─ win.py                # 視窗聚焦/前景處理（pywinauto）
│  └─ utils.py              # 日誌、截圖、重試、時間工具
├─ tasks/                   # 任務腳本（YAML）
│  └─ album_broadcast.yaml
├─ templates/               # UI 模板（PNG）
│  ├─ btn_album.png
│  ├─ btn_send.png
│  └─ ...
├─ data/                    # 圖片素材與相簿來源
│  └─ album_2025W44/
├─ lists/                   # 受眾、黑名單、退訂名單 CSV
│  ├─ recipients.csv
│  ├─ blacklist.csv
│  └─ unsubscribe.csv
├─ logs/                    # 執行日誌
├─ reports/                 # 報表與截圖證據
├─ .env
├─ config.yaml
├─ requirements.txt
└─ main.py                  # 入口：python main.py run tasks/xxx.yaml

6. UI 模板製作指引（關鍵）

DPI=100%、固定語系：請在「將要運行的環境」抓圖。

抓取清晰圖塊：像「送出」、「相簿按鈕」、「搜尋框圖示」等，儘量含有高對比特徵。

命名：templates/btn_send.png, templates/ico_album.png…

大小：以實際 UI 像素為準，避免縮放後另存。

門檻（threshold）：預設 0.92，可依 UI 差異在 YAML 調整 0.88–0.97。

多樣板支援：同一元素可放多張（淺色/深色模式）；YAML 中可給 templates: [a.png, b.png]。

7. 設定檔 config.yaml（示例）
screen:
  width: 1920
  height: 1080
  dpi_scale: 1.0

vision:
  default_threshold: 0.92
  timeout_sec: 10
  retries: 2
  search_region: null  # 或 [x, y, w, h] 只在局部找

delays:
  click_ms: [60, 120]
  type_ms: [40, 80]
  random_jitter_ms: [100, 600]

line:
  exe_path: "C:\\Users\\You\\AppData\\Local\\LINE\\bin\\LineLauncher.exe"
  lang: "zh-TW"

throttle:
  max_recipients_per_run: 50
  min_interval_sec: [15, 45]   # 取隨機值
  daily_cap: 300

lists:
  recipients: "lists/recipients.csv"
  blacklist: "lists/blacklist.csv"
  unsubscribe: "lists/unsubscribe.csv"

logs:
  screenshot_on_fail: true
  record_screen: false


.env（示例）

AUTOA_ENV=dev
TESSERACT_PATH=C:\Program Files\Tesseract-OCR\tesseract.exe

8. 任務 DSL（YAML）範例：相簿廣播
name: album_broadcast
variables:
  album_title: "本週相簿"
  album_dir: "data/album_2025W44"   # 要發送的圖片資料夾
limits:
  max_recipients: 40
  interval_sec: [18, 36]            # 每位之間隨機延遲
steps:
  - focus_app: { name: "LINE" }     # 聚焦 LINE 視窗
  - ensure_logged_in: {}            # （可選）檢查登入標誌模板

  # 開啟搜尋，定位收件者
  - locate_click: { template: "ico_search.png", threshold: 0.93 }
  - type_text:    { text: "{{recipient.name}}" }
  - press:        { keys: ["enter"] }
  - wait:         { sec: [0.8, 1.6] }

  # 建立相簿/插入圖片
  - locate_click: { template: "btn_album.png" }
  - wait:         { sec: [0.5, 1.0] }
  - click:        { x: "{{anchors.album.add.x}}", y: "{{anchors.album.add.y}}" }
  - upload_dir:   { path: "{{album_dir}}" }      # 系統開啟檔案對話框→貼路徑→Enter
  - wait:         { sec: [1.5, 3.0] }

  # 輸入相簿名稱並送出
  - type_text:    { text: "{{album_title}}" }
  - press:        { keys: ["enter"] }

for_each_recipient:
  source: "lists/recipients.csv"    # 欄位需包含 name 或 uid
  skip_if_blacklisted: true
  steps_ref: steps


可用動作一覽

focus_app, ensure_logged_in（模板校驗）

locate_click（以模板比對定位並點擊；支援 templates: [], threshold, timeout_sec）

click, move, drag_drop, scroll

type_text, press（keys: ["ctrl","v"] 等）

upload_file, upload_dir（系統檔案對話框）

wait（固定或區間隨機）

for_each_recipient（清單回圈，含黑名單跳過、逐筆失敗重試）

assert_present / assert_absent（模板存在性斷言）

screenshot（存證）

goto_step / label（必要時流程跳轉）

9. 執行方式
# 乾跑（不點擊，只畫框示意）
python main.py dryrun tasks/album_broadcast.yaml

# 正式執行
python main.py run tasks/album_broadcast.yaml

# 從特定步驟續跑
python main.py run tasks/album_broadcast.yaml --from "type_text:album_title"

10. 排程（Windows 排程器）
# 每日 10:00 執行一次
schtasks /Create /TN "AUTOA Album W44" /TR "\"C:\Path\to\python.exe\" C:\Path\to\AUTOA-RPA\main.py run C:\Path\to\AUTOA-RPA\tasks\album_broadcast.yaml" /SC DAILY /ST 10:00

11. 名單與節流

lists/recipients.csv：name,uid,tags,...（最少保留 name 或 uid 以便搜尋）

lists/blacklist.csv、lists/unsubscribe.csv：uid 一行一筆

節流策略：

每輪上限 max_recipients_per_run

每位間隔 min_interval_sec（區間隨機）

每日總量 daily_cap

人類化：點擊/輸入/等待都採隨機抖動，避免固定節奏

12. 日誌與存證

logs/autoa.log：關鍵事件、錯誤堆疊

reports/：失敗截圖、（可選）全程錄影片段

每筆收件者形成一筆記錄：recipient, status, retries, timestamp

13. 常見問題（FAQ）

Q1：找不到模板/一直逾時

確認解析度、DPI=100%、深/淺色模式與抓圖時一致

調降 threshold 至 0.90–0.92 試試

指定 search_region 限縮搜尋範圍加速與減少誤判

Q2：多螢幕定位錯位

將 LINE 固定於主螢幕，或在 win.py 內先移動視窗至 (0,0)

Q3：拖放/上傳對話框無反應

試 upload_file（貼路徑 + Enter）替代拖放

檢查是否有系統管理員權限衝突（Admin 與非 Admin 之間）

Q4：鍵盤語系導致快捷鍵錯誤

將系統輸入法切換為 EN-US（英文）再執行

14. 擴充建議

OCR 校驗：對搜尋框或標題進行文字比對，提高成功率

pywinauto 元件路徑：在可能的情況下以控件樹輔助定位

容器化：封裝成 portable 套件或以 PyInstaller 打包

GUI 後台：提供任務管理、名單管理、排程、報表匯出

15. 合規/隱私建議（務必）

僅發送給已同意接收之對象，並提供退訂通道

嚴格控制節流，避免造成騷擾或觸發系統防護

不蒐集無必要的個資，日誌去識別化或最小化保存

依地方法規（如 GDPR）確保資料處理與保存合法

16. 版本策略與變更紀錄

版本號：MAJOR.MINOR.PATCH

v0.1.0：初版 README、任務 DSL、模板製作指引、MVP 結構說明

17. 授權

建議以 MIT 授權（僅供參考，可依實際情況調整）。