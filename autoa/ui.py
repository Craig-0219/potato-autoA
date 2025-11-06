"""Tkinter UI for AUTOA RPA control panel."""
from __future__ import annotations

import csv
import locale
import random
import subprocess
import threading
import time
from collections import deque
import cv2
import numpy as np

from pathlib import Path
from typing import Any, Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from autoa.flow import FlowExecutor
from autoa.guards import GuardRail
from autoa.rpa import RPAController
from autoa.vision import TemplateMatcher
from autoa.line_automation import CycleResult, LineAutomationError, cycle_friend_chats


STATUS_COLORS = {
    "ok": "#d4f4dd",
    "warn": "#fff4d6",
    "fail": "#f8d7da",
    "pending": "#f0f0f0",
}

LOG_CAPACITY = 200
LINE_REFRESH_MS = 5000
TEMPLATE_CONFIDENCE = 0.88
MAX_CHAT_TEST_RECIPIENTS = 8

TEST_CASES_TEXT = """
TC-01 環境相關設定
前置：Windows 10/11；螢幕 1920x1080；DPI = 100%。
步驟：開啟程式，UI 顯示解析度與 DPI 為綠色「合規」。
預期：環境檢查通過；若不合規應警示並禁止正式執行。

TC-02 語系選擇
前置：LINE 桌面版語系為 zh-T（或模板語系）。
步驟：點「驗證模板」。
預期：搜尋圖示、附件、送出模板皆通過；語系錯誤時顯示修正建議。

TC-03 執行步驟：乾跑
前置：收件者「測試」、訊息「Hello」、圖片 test.jpg。
步驟：模式切到「乾跑」、按「開始」。
預期：UI 只畫框提示、不實際點擊；日誌顯示每一步命中結果。

TC-04 新增相片（正式送圖）
前置：TC-03 通過，模式改為「正式」。
步驟：開始 → 聚焦 LINE → 搜尋 → 輸入訊息 → 附件 → 選檔 → 送出。
預期：聊天視窗出現圖片與文字各一；日誌記錄路徑與送出成功。

TC-05 訊息僅文字
前置：圖片欄位留空、訊息「純文字」。
步驟：正式執行。
預期：只送出文字；日誌註記「未選擇圖片略過附件步驟」。

TC-06 執行次數
前置：執行次數固定 1（不可編輯）。
步驟：正式執行。
預期：流程僅跑一次；完成後狀態回到「閒置」。

TC-07 節流與延遲
前置：設定延遲範圍（例如 1-2 秒）。
步驟：正式執行並觀察停頓。
預期：步驟間延遲落在範圍內；日誌紀錄每次延遲。

TC-08 模板未命中
前置：切換 LINE 語系或主題使模板失敗。
步驟：乾跑驗證。
預期：UI 顯示紅色狀態與建議，禁止正式執行。

TC-09 受眾錯誤
前置：輸入不存在的收件者。
步驟：正式執行。
預期：搜尋不到對話 → 顯示「未命中收件者」並終止。

TC-10 權限視窗前景
前置：LINE 被最小化或遮擋。
步驟：正式執行。
預期：先嘗試 bring-to-front；若失敗提示「請將 LINE 置於主螢幕」。
""".strip()


def build_executor() -> FlowExecutor:
    """Create a FlowExecutor with default dependencies."""
    rpa = RPAController()
    matcher = TemplateMatcher()
    guards = GuardRail()
    return FlowExecutor(rpa=rpa, matcher=matcher, guards=guards)


class AutoaApp:
    """Main Tkinter UI."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("AUTOA RPA Control Panel")

        self.executor = build_executor()

        self.recipient_var = tk.StringVar()
        self.recipient_choice_var = tk.StringVar()
        self.message_text: ScrolledText | None = None
        self.image_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="dryrun")
        self.exec_count_var = tk.StringVar(value="1")
        self.throttle_min_var = tk.StringVar(value="1.0")
        self.throttle_max_var = tk.StringVar(value="2.0")
        self.current_step_var = tk.StringVar(value="閒置")
        self.progress_var = tk.DoubleVar(value=0.0)

        self.running = False
        self.paused = False
        self.stop_event = threading.Event()
        self.pause_condition = threading.Condition()
        self.worker_thread: threading.Thread | None = None

        self.log_text: ScrolledText | None = None
        self.progress_bar: ttk.Progressbar | None = None
        self.log_lines: deque[str] = deque(maxlen=LOG_CAPACITY)

        self.recipient_combo: ttk.Combobox | None = None
        self.start_button: ttk.Button | None = None
        self.pause_button: ttk.Button | None = None
        self.stop_button: ttk.Button | None = None
        self.screenshot_button: ttk.Button | None = None
        self.friend_cycle_thread: threading.Thread | None = None

        self.system_status_labels: dict[str, tk.Label] = {}
        self.system_status: dict[str, bool] = {
            "resolution": False,
            "dpi": False,
            "language": False,
            "line": False,
        }

        self.friend_list_template = Path("templates/friend-list.png")
        self.message_cube_template = Path("templates/message_cube.png")
        self.arrow_section_templates: list[tuple[str, Path, str]] = [
            ("收藏", Path("templates/favorite.png"), "hide"),
            ("社群", Path("templates/community.png"), "hide"),
            ("群組", Path("templates/group.png"), "hide"),
            ("好友", Path("templates/friend.png"), "show"),
        ]
        self.hide_arrow_template = Path("templates/hide.png")
        self.show_arrow_template = Path("templates/show.png")

        self.notebook: ttk.Notebook | None = None
        self.main_canvas: tk.Canvas | None = None
        self.main_canvas_content: int | None = None

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self.load_recipients()
        self.run_system_checks()
        self.refresh_line_status()

    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        main_tab = ttk.Frame(self.notebook)
        test_tab = ttk.Frame(self.notebook)
        self.notebook.add(main_tab, text="主控台")
        self.notebook.add(test_tab, text="測試")

        main_tab.columnconfigure(0, weight=1)
        main_tab.rowconfigure(0, weight=1)

        self.main_canvas = tk.Canvas(main_tab, borderwidth=0, highlightthickness=0)
        main_scroll = ttk.Scrollbar(main_tab, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.grid(row=0, column=0, sticky="nsew")
        main_scroll.grid(row=0, column=1, sticky="ns")
        self.main_canvas.configure(yscrollcommand=main_scroll.set)

        container = ttk.Frame(self.main_canvas, padding=12)
        container.columnconfigure(0, weight=1)
        self.main_canvas_content = self.main_canvas.create_window((0, 0), window=container, anchor="nw")

        container.bind(
            "<Configure>",
            lambda event: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")),
        )
        self.main_canvas.bind(
            "<Configure>",
            lambda event: self.main_canvas.itemconfigure(self.main_canvas_content, width=event.width),
        )

        self._bind_scroll_events()

        self._build_recipient_section(container)
        self._build_message_section(container)
        self._build_execution_options(container)
        self._build_system_section(container)
        self._build_execution_section(container)
        self._build_testcase_section(container)

        test_frame = ttk.Frame(test_tab, padding=12)
        test_frame.grid(row=0, column=0, sticky="nsew")
        test_tab.columnconfigure(0, weight=1)
        test_tab.rowconfigure(0, weight=1)
        self._build_test_shortcuts_section(test_frame)

    def _bind_scroll_events(self) -> None:
        self.root.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self._on_mousewheel_linux, add="+")
        self.root.bind_all("<Button-5>", self._on_mousewheel_linux, add="+")

    def _on_mousewheel(self, event: tk.Event) -> None:
        if self.main_canvas is None:
            return
        delta = -int(event.delta / 120) if getattr(event, "delta", 0) else 0
        if delta:
            self.main_canvas.yview_scroll(delta, "units")

    def _on_mousewheel_linux(self, event: tk.Event) -> None:
        if self.main_canvas is None:
            return
        direction = -1 if getattr(event, "num", 0) == 4 else 1
        self.main_canvas.yview_scroll(direction, "units")

    def _build_recipient_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="收件者", padding=10)
        frame.grid(row=0, column=0, sticky="ew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="收件者：").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(frame, textvariable=self.recipient_var)
        entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="清除", command=lambda: self.recipient_var.set("")).grid(row=0, column=2)

        ttk.Label(frame, text="從名單載入：").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.recipient_combo = ttk.Combobox(
            frame,
            textvariable=self.recipient_choice_var,
            state="readonly",
        )
        self.recipient_combo.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        self.recipient_combo.bind("<<ComboboxSelected>>", self.on_recipient_selected)
        ttk.Button(frame, text="重新載入", command=self.load_recipients).grid(row=1, column=2, pady=(8, 0))

    def _build_message_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="訊息內容", padding=10)
        frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        frame.columnconfigure(0, weight=1)

        self.message_text = ScrolledText(frame, wrap="word", height=6)
        self.message_text.grid(row=0, column=0, columnspan=3, sticky="nsew")
        frame.rowconfigure(0, weight=1)

        ttk.Label(frame, text="附件圖片檔案：").grid(row=1, column=0, sticky="w", pady=(8, 0))
        image_entry = ttk.Entry(frame, textvariable=self.image_var)
        image_entry.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        ttk.Button(frame, text="選擇圖片...", command=self.browse_image).grid(row=1, column=2, pady=(8, 0))

    def _build_execution_options(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="執行設定", padding=10)
        frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="模式：").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(frame, text="乾跑", value="dryrun", variable=self.mode_var).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(frame, text="正式", value="live", variable=self.mode_var).grid(row=0, column=2, sticky="w")

        ttk.Label(frame, text="執行次數：").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.exec_count_var, state="disabled", width=6).grid(
            row=1,
            column=1,
            sticky="w",
            pady=(8, 0),
        )

        ttk.Label(frame, text="節流秒數 (min ~ max)：").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.throttle_min_var, width=8).grid(row=2, column=1, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.throttle_max_var, width=8).grid(row=2, column=2, sticky="w", pady=(8, 0))

        ttk.Button(frame, text="驗證模板", command=self.handle_verify_templates).grid(row=0, column=3, sticky="e")
        ttk.Button(frame, text="重新檢查環境", command=self.run_system_checks).grid(row=1, column=3, sticky="e", pady=(8, 0))

    def _build_system_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="系統狀態", padding=10)
        frame.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        status_items = [
            ("resolution", "解析度：未檢查"),
            ("dpi", "DPI：未檢查"),
            ("language", "語系：未檢查"),
            ("line", "LINE 程式：偵測中"),
        ]
        for index, (key, text) in enumerate(status_items):
            label = tk.Label(frame, text=text, anchor="w", padx=6, pady=4, bg=STATUS_COLORS["pending"])
            label.grid(row=index, column=0, columnspan=2, sticky="ew", pady=(0 if index == 0 else 4, 0))
            self.system_status_labels[key] = label

        ttk.Button(frame, text="立即檢查", command=self.run_system_checks).grid(
            row=len(status_items),
            column=0,
            sticky="w",
            pady=(8, 0),
        )
        ttk.Button(frame, text="模板檢查", command=self.handle_verify_templates).grid(
            row=len(status_items),
            column=1,
            sticky="e",
            pady=(8, 0),
        )

    def _build_execution_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="執行區", padding=10)
        frame.grid(row=4, column=0, sticky="nsew", pady=(8, 0))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky="ew")

        ttk.Label(frame, textvariable=self.current_step_var).grid(row=1, column=0, sticky="w", pady=(6, 0))

        self.log_text = ScrolledText(frame, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=3, column=0, sticky="nsew", pady=(8, 0))

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        button_frame.columnconfigure((0, 1, 2, 3), weight=1)

        self.start_button = ttk.Button(button_frame, text="開始", command=self.handle_start)
        self.start_button.grid(row=0, column=0, padx=4)

        self.pause_button = ttk.Button(button_frame, text="暫停", command=self.handle_pause, state=tk.DISABLED)
        self.pause_button.grid(row=0, column=1, padx=4)

        self.stop_button = ttk.Button(button_frame, text="終止", command=self.handle_stop, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=2, padx=4)

        self.screenshot_button = ttk.Button(button_frame, text="截圖", command=self.handle_screenshot)
        self.screenshot_button.grid(row=0, column=3, padx=4)

    def _build_testcase_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="測試項 (Test Cases)", padding=10)
        frame.grid(row=5, column=0, sticky="nsew", pady=(8, 0))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        testcase_text = ScrolledText(frame, height=12, wrap="word")
        testcase_text.grid(row=0, column=0, sticky="nsew")
        testcase_text.insert("1.0", TEST_CASES_TEXT)
        testcase_text.configure(state="disabled")

    def _build_test_shortcuts_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="快捷測試", padding=10)
        frame.grid(row=0, column=0, sticky="ew")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=0)

        ttk.Label(frame, text="測試開啟好友選單").grid(row=0, column=0, sticky="w")
        ttk.Button(frame, text="執行", command=self.handle_test_open_friend_menu, width=12).grid(
            row=0,
            column=1,
            sticky="e",
        )

        ttk.Label(frame, text="測試送出訊息").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Button(frame, text="貼上並送出", command=self.handle_test_send_message, width=12).grid(
            row=1,
            column=1,
            sticky="e",
            pady=(8, 0),
        )

        ttk.Label(frame, text="箭頭校正").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Button(frame, text="開始校正", command=self.handle_align_arrow_sections, width=12).grid(
            row=2,
            column=1,
            sticky="e",
            pady=(8, 0),
        )

        ttk.Label(frame, text="依序開啟聊天窗").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Button(frame, text="開始測試", command=self.handle_cycle_friend_chats, width=12).grid(
            row=3,
            column=1,
            sticky="e",
            pady=(8, 0),
        )
    # ------------------------------------------------------------------
    def append_log(self, message: str) -> None:
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, lambda: self.append_log(message))
            return

        timestamp = time.strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}"
        self.log_lines.append(line)

        if self.log_text is None:
            return

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("end", "\n".join(self.log_lines) + "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    def load_recipients(self, *_args) -> None:
        csv_path = Path("lists/recipients.csv")
        values: list[str] = []
        if csv_path.exists():
            try:
                with csv_path.open("r", encoding="utf-8") as handle:
                    reader = csv.DictReader(handle)
                    for row in reader:
                        identifier = row.get("name") or row.get("uid")
                        if identifier:
                            values.append(identifier.strip())
                if values:
                    self.append_log(f"CSV 名單載入 {len(values)} 筆。")
                else:
                    self.append_log("CSV 名單為空。")
            except Exception as exc:
                self.append_log(f"讀取 CSV 失敗：{exc}")
        else:
            self.append_log("找不到 lists/recipients.csv，請先建立名單。")

        if self.recipient_combo is not None:
            self.recipient_combo["values"] = values

    def on_recipient_selected(self, _event: Any) -> None:
        choice = self.recipient_choice_var.get()
        if choice:
            self.recipient_var.set(choice)
            self.append_log(f"已選擇收件者：{choice}")

    def browse_image(self) -> None:
        selected = filedialog.askopenfilename(
            title="選擇圖片檔案",
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp"), ("All files", "*.*")],
            initialdir=str(Path.cwd() / "data"),
        )
        if selected:
            self.image_var.set(selected)
            self.append_log(f"圖片已選：{selected}")

    def get_throttle_range(self) -> tuple[float, float] | None:
        try:
            min_delay = float(self.throttle_min_var.get())
            max_delay = float(self.throttle_max_var.get())
        except ValueError:
            messagebox.showwarning("節流設定", "請輸入合法的延遲數值。")
            return None
        if min_delay < 0 or max_delay < 0 or min_delay > max_delay:
            messagebox.showwarning("節流設定", "請確認最小值與最大值設定正確。")
            return None
        return min_delay, max_delay

    def run_system_checks(self) -> None:
        width = self.root.winfo_screenwidth()
        height = self.root.winfo_screenheight()
        try:
            dpi = int(self.root.winfo_fpixels("1i"))
        except tk.TclError:
            dpi = 96

        locale_info = locale.getdefaultlocale()
        locale_code = locale_info[0] if locale_info and locale_info[0] else "未知"

        resolution_ok = width == 1920 and height == 1080
        dpi_ok = 94 <= dpi <= 110
        language_ok = locale_code.lower().startswith("zh") if locale_code != "未知" else False

        self.system_status["resolution"] = resolution_ok
        self.system_status["dpi"] = dpi_ok
        self.system_status["language"] = language_ok

        resolution_text = f"解析度：{width}x{height} {'(合規)' if resolution_ok else '(建議調整為 1920x1080)'}"
        dpi_text = f"DPI：{dpi} {'(合規)' if dpi_ok else '(建議調整為 100%)'}"
        language_text = f"語系：{locale_code} {'(合規)' if language_ok else '(建議切換為 zh-T)'}"

        self._update_status_label(self.system_status_labels.get("resolution"), "ok" if resolution_ok else "warn", resolution_text)
        self._update_status_label(self.system_status_labels.get("dpi"), "ok" if dpi_ok else "warn", dpi_text)
        self._update_status_label(
            self.system_status_labels.get("language"),
            "ok" if language_ok else "warn",
            language_text,
        )

        line_running = self._is_line_running()
        self.system_status["line"] = line_running
        line_text = "LINE 程式：已啟動" if line_running else "LINE 程式：未啟動 (請先開啟 LINE)"
        self._update_status_label(self.system_status_labels.get("line"), "ok" if line_running else "fail", line_text)

        self.append_log("系統環境檢查完成。")

    def _update_status_label(self, label: tk.Label | None, state: str, text: str) -> None:
        if label is None:
            return
        label.configure(text=text, bg=STATUS_COLORS.get(state, STATUS_COLORS["pending"]))

    def refresh_line_status(self) -> None:
        running = self._is_line_running()
        self.system_status["line"] = running
        label = self.system_status_labels.get("line")
        if label:
            text = "LINE 程式：已啟動" if running else "LINE 程式：未啟動 (請先開啟 LINE)"
            self._update_status_label(label, "ok" if running else "fail", text)
        self.root.after(LINE_REFRESH_MS, self.refresh_line_status)

    def _is_line_running(self) -> bool:
        try:
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq line.exe"],
                capture_output=True,
                text=True,
                check=False,
            )
        except Exception:
            return False
        return "line.exe" in result.stdout.lower()

    # ------------------------------------------------------------------
    def handle_start(self) -> None:
        if self.running:
            messagebox.showinfo("執行中", "流程已在執行。")
            return

        recipient = self.recipient_var.get().strip()
        if not recipient:
            messagebox.showwarning("收件者", "請輸入收件者。")
            return

        message = ""
        if self.message_text is not None:
            message = self.message_text.get("1.0", "end").strip()

        if not message:
            proceed = messagebox.askyesno("訊息為空", "訊息內容為空，確定要繼續嗎？")
            if not proceed:
                return

        throttle = self.get_throttle_range()
        if throttle is None:
            return

        image_path = self.image_var.get().strip() or None
        dry_run = self.mode_var.get() != "live"

        self.stop_event = threading.Event()
        self.running = True
        self.paused = False

        if self.pause_button is not None:
            self.pause_button.configure(text="暫停")

        self._toggle_buttons(running=True)
        self.append_log(f"開始執行流程{'（乾跑模式）' if dry_run else ''}。")

        self.worker_thread = threading.Thread(
            target=self._run_flow,
            args=(recipient, message, image_path, throttle, dry_run),
            daemon=True,
        )
        self.worker_thread.start()

    def _toggle_buttons(self, running: bool) -> None:
        if self.start_button is not None:
            self.start_button.configure(state=tk.DISABLED if running else tk.NORMAL)
        if self.pause_button is not None:
            self.pause_button.configure(state=tk.NORMAL if running else tk.DISABLED)
        if self.stop_button is not None:
            self.stop_button.configure(state=tk.NORMAL if running else tk.DISABLED)

    def handle_pause(self) -> None:
        if not self.running:
            return
        with self.pause_condition:
            self.paused = not self.paused
            if self.paused:
                if self.pause_button is not None:
                    self.pause_button.configure(text="繼續")
                self.append_log("流程已暫停。")
            else:
                if self.pause_button is not None:
                    self.pause_button.configure(text="暫停")
                self.pause_condition.notify_all()
                self.append_log("流程繼續。")

    def handle_stop(self) -> None:
        if not self.running:
            return
        self.append_log("收到終止指令，準備停止流程。")
        self.stop_event.set()
        with self.pause_condition:
            self.paused = False
            self.pause_condition.notify_all()

    def _run_flow(
        self,
        recipient: str,
        message: str,
        image_path: str | None,
        throttle: tuple[float, float],
        dry_run: bool,
    ) -> None:
        steps = [
            "聚焦 LINE 視窗",
            "搜尋收件者",
            "貼上訊息內容",
            "準備附件圖片" if image_path else None,
            "校正側邊欄箭頭",
            "模擬送出訊息" if dry_run else "送出訊息",
        ]
        sequence = [step for step in steps if step is not None]

        total = len(sequence) if sequence else 1
        success = False

        for index, step in enumerate(sequence, start=1):
            if self.stop_event.is_set():
                break

            if self._wait_if_paused():
                break

            self._set_current_step(step)
            self.append_log(f"執行步驟：{step}")

            progress = index / total * 100
            self._set_progress(progress)

            simulated = random.uniform(*throttle)
            time.sleep(simulated)
            self.append_log(f"步驟完成，耗時 {simulated:.2f} 秒。")

        else:
            if not self.stop_event.is_set():
                success = True

        self.root.after(0, lambda: self._on_worker_finished(success))

    def _wait_if_paused(self) -> bool:
        with self.pause_condition:
            while self.paused and not self.stop_event.is_set():
                self.pause_condition.wait(timeout=0.2)
        return self.stop_event.is_set()

    def _set_current_step(self, text: str) -> None:
        if threading.current_thread() is threading.main_thread():
            self.current_step_var.set(text)
        else:
            self.root.after(0, lambda: self.current_step_var.set(text))

    def _set_progress(self, value: float) -> None:
        if threading.current_thread() is threading.main_thread():
            self.progress_var.set(value)
        else:
            self.root.after(0, lambda: self.progress_var.set(value))

    def _on_worker_finished(self, success: bool) -> None:
        if success:
            self.append_log("流程完成。")
            self.current_step_var.set("完成")
            self.progress_var.set(100.0)
        else:
            if self.stop_event.is_set():
                self.append_log("流程已中止。")
            else:
                self.append_log("流程未完成。")
            self.current_step_var.set("閒置")
            self.progress_var.set(0.0)

        self.running = False
        self.paused = False
        self.worker_thread = None
        self._toggle_buttons(running=False)
        if self.pause_button is not None:
            self.pause_button.configure(text="暫停")
    # ------------------------------------------------------------------
    def handle_screenshot(self) -> None:
        try:
            import pyautogui
        except ImportError as exc:
            messagebox.showerror("截圖失敗", f"無法載入 pyautogui：{exc}")
            return

        reports_dir = Path("reports")
        reports_dir.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        output_path = reports_dir / f"screenshot-{timestamp}.png"

        try:
            screenshot = pyautogui.screenshot()
            screenshot.save(output_path)
        except Exception as exc:
            self.append_log(f"截圖失敗：{exc}")
            messagebox.showerror("截圖失敗", f"無法儲存截圖：{exc}")
            return

        self.append_log(f"已儲存截圖：{output_path}")
        messagebox.showinfo("截圖完成", f"已儲存至 {output_path}")

    def handle_verify_templates(self) -> None:
        missing = [str(path) for path in self._template_paths() if not path.exists()]
        if missing:
            text = "缺少下列模板檔案，請確認 templates 目錄：\n" + "\n".join(missing)
            self.append_log("模板檢查失敗。")
            messagebox.showwarning("模板檢查", text)
        else:
            self.append_log("模板檢查通過。")
            messagebox.showinfo("模板檢查", "所有必要模板檔案皆存在。")

    def handle_test_open_friend_menu(self) -> None:
        if not self.friend_list_template.exists():
            messagebox.showwarning("測試模板", f"找不到模板：{self.friend_list_template}")
            self.append_log("測試模板缺失，無法進行好友選單測試。")
            return

        try:
            import pyautogui
        except ImportError as exc:
            messagebox.showerror("測試失敗", f"無法載入 pyautogui：{exc}")
            return

        if not self._focus_line_window(pyautogui):
            messagebox.showwarning("測試失敗", "未偵測到 LINE 視窗。")
            return

        location = self._try_locate(pyautogui, self.friend_list_template, confidence=0.9)
        if location is None:
            self.append_log("未找到好友選單按鈕。")
            messagebox.showwarning("測試結果", "未偵測到好友選單按鈕，請確認 LINE 介面。")
            return

        try:
            center = pyautogui.center(location)
            pyautogui.click(center.x, center.y)
        except Exception as exc:
            self.append_log(f"點擊好友選單失敗：{exc}")
            messagebox.showerror("測試失敗", f"點擊好友選單按鈕失敗：{exc}")
            return

        self.append_log("好友選單測試完成。")
        messagebox.showinfo("測試完成", "已嘗試點擊好友選單按鈕。")

    def handle_test_send_message(self) -> None:
        if self.message_text is None:
            return

        message = self.message_text.get("1.0", "end").strip()
        if not message:
            messagebox.showwarning("測試訊息", "訊息內容為空。")
            return

        try:
            import pyautogui
        except ImportError as exc:
            messagebox.showerror("測試失敗", f"無法載入 pyautogui：{exc}")
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(message)
        self.root.update()

        if not self._focus_line_window(pyautogui):
            messagebox.showwarning("測試失敗", "未偵測到 LINE 視窗。")
            return

        try:
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.2)
            pyautogui.press("enter")
        except Exception as exc:
            self.append_log(f"貼上訊息失敗：{exc}")
            messagebox.showerror("測試失敗", f"貼上或送出訊息失敗：{exc}")
            return

        self.append_log("測試訊息已貼上並送出。")
        messagebox.showinfo("測試完成", "已嘗試貼上並送出訊息。")

    def handle_cycle_friend_chats(self) -> None:
        if self.friend_cycle_thread and self.friend_cycle_thread.is_alive():
            messagebox.showinfo('聊天測試', '聊天測試正在進行中，請稍候。')
            return

        values: list[str] = []
        if self.recipient_combo is not None:
            try:
                values = [v for v in self.recipient_combo['values'] if v]  # type: ignore[index]
            except Exception:
                values = []
        if not values:
            self.load_recipients()
            if self.recipient_combo is not None:
                try:
                    values = [v for v in self.recipient_combo['values'] if v]  # type: ignore[index]
                except Exception:
                    values = []
        if not values:
            self.append_log('好友名單為空，無法依序開啟聊天窗。')
            messagebox.showwarning('好友名單', '名單為空或尚未載入，請先準備 lists/recipients.csv。')
            return

        limit = min(len(values), MAX_CHAT_TEST_RECIPIENTS)
        if limit <= 0:
            self.append_log('好友名單筆數不足，略過聊天測試。')
            messagebox.showwarning('聊天測試', '名單筆數不足，無法進行聊天測試。')
            return

        message = ''
        if self.message_text is not None:
            message = self.message_text.get('1.0', 'end').strip()
        payload = message or None

        self.append_log(f'聊天測試：準備使用 pywinauto 依序開啟 {limit} 位好友。')
        thread = threading.Thread(
            target=self._cycle_friend_chats_worker,
            args=(limit, payload),
            daemon=True,
        )
        self.friend_cycle_thread = thread
        thread.start()

    def _cycle_friend_chats_worker(self, limit: int, message: str | None) -> None:
        try:
            result = cycle_friend_chats(limit=limit, message=message, log=self.append_log)
        except LineAutomationError as exc:
            self.append_log(f'聊天測試失敗：{exc}')
            self.root.after(0, lambda err=exc: messagebox.showerror('聊天測試', str(err)))
        except Exception as exc:
            self.append_log(f'聊天測試發生未預期錯誤：{exc}')
            self.root.after(0, lambda err=exc: messagebox.showerror('聊天測試', f'執行失敗：{err}'))
        else:
            self.root.after(0, lambda: self._show_cycle_friend_result(result))
        finally:
            self.friend_cycle_thread = None

    def _show_cycle_friend_result(self, result: CycleResult) -> None:
        processed = len(result.processed)
        failed = len(result.failed)
        self.append_log(f'聊天測試完成：成功 {processed} 位，失敗 {failed} 位。')
        if result.failed:
            info = '未成功：' + ', '.join(result.failed)
            messagebox.showwarning('聊天測試', f'部分好友未開啟成功。\n{info}')
        else:
            suffix = '，已達清單尾端。' if result.reached_end else ''
            messagebox.showinfo('聊天測試', f'已依序開啟 {processed} 位好友的聊天視窗{suffix}')

    def handle_align_arrow_sections(self) -> None:
        sequence = [
            ("收藏", self.arrow_section_templates[0][1], "hide"),
            ("社群", self.arrow_section_templates[1][1], "hide"),
            ("群組", self.arrow_section_templates[2][1], "hide"),
            ("好友", self.arrow_section_templates[3][1], "show"),
        ]
        try:
            import pyautogui
        except ImportError as exc:
            messagebox.showerror("箭頭校正", f"無法載入 pyautogui：{exc}")
            return

        self.append_log("開始箭頭校正（線性模式）")

        reports: list[str] = []
        try:
            screen_width, screen_height = pyautogui.size()
        except Exception:
            screen_width = screen_height = None

        for idx, (name, template, expectation) in enumerate(sequence):
            result = self._calibrate_section_once(
                pyautogui,
                name=name,
                template=template,
                expectation=expectation,
                screen_size=(screen_width, screen_height),
            )
            reports.append(result)

            # Add delay between sections to allow UI to stabilize
            # Skip delay after the last section
            if idx < len(sequence) - 1:
                self.append_log(f"等待 UI 穩定...")
                time.sleep(0.6)

        messagebox.showinfo("箭頭校正結果", "\n".join(reports))

    def _scroll_left_panel_to_top(
        self,
        pyautogui_module: Any,
    ) -> None:
        try:
            friend_box = self._try_locate(pyautogui_module, self.friend_list_template, confidence=0.88)
        except Exception:
            friend_box = None
        coords = self._box_to_tuple(friend_box)
        if coords is None:
            self.append_log("無法定位好友區塊，略過自動捲動。")
            return

        x = coords[0] + coords[2] / 2
        y = coords[1] + coords[3] / 2

        try:
            pyautogui_module.moveTo(x, y, duration=0.15)
        except Exception:
            pass

        for _ in range(6):
            try:
                pyautogui_module.scroll(320)
            except Exception:
                break
            time.sleep(0.05)

    def _calibrate_section_once(
        self,
        pyautogui_module: Any,
        *,
        name: str,
        template: Path,
        expectation: str,
        screen_size: tuple[int | None, int | None],
    ) -> str:
        self.append_log(f"校正 {name}，預期狀態 {expectation}")

        screen_width, screen_height = screen_size
        target_region = (0, 0, screen_width, screen_height) if screen_width and screen_height else None

        # 使用 locateAllOnScreen 查找所有匹配的位置
        all_locations = self._try_locate_all(pyautogui_module, template, region=target_region, confidence=0.88)

        if not all_locations:
            self.append_log(f"{name}：模板未命中")
            return f"{name}: 未命中模板"

        self.append_log(f"{name}：找到 {len(all_locations)} 個匹配項")

        processed_count = 0
        skipped_count = 0

        failed_detection_count = 0

        for idx, location in enumerate(all_locations, start=1):
            location_tuple = self._box_to_tuple(location)
            if location_tuple is None:
                self.append_log(f"{name} 第 {idx} 項：模板定位解析失敗")
                continue

            # 計算點擊位置（區塊標題右側，用於切換箭頭）
            # 使用區塊右側而不是中心，因為箭頭在右側
            click_x = location_tuple[0] + location_tuple[2] - 30  # 右側往左30像素
            click_y = location_tuple[1] + location_tuple[3] / 2

            # 檢測當前箭頭狀態
            current_state = self.detect_arrow_state(pyautogui_module, location_tuple)
            self.append_log(f"{name} 第 {idx} 項判定：{current_state}，位置 ({int(click_x)}, {int(click_y)})")

            # 如果無法檢測到箭頭狀態，記錄並嘗試點擊
            if current_state is None:
                self.append_log(f"{name} 第 {idx} 項：⚠️ 無法檢測箭頭狀態，將嘗試點擊切換")
                failed_detection_count += 1
                # 繼續執行點擊邏輯，因為我們需要確保狀態符合預期
            elif current_state == expectation:
                # 狀態已符合預期，跳過
                self.append_log(f"{name} 第 {idx} 項已符合預期，跳過")
                skipped_count += 1
                continue

            # 狀態不符合預期或無法檢測，需要點擊切換
            try:
                pyautogui_module.moveTo(click_x, click_y, duration=0.15)
                pyautogui_module.click(click_x, click_y)
                processed_count += 1

                # Wait for UI animation to complete
                time.sleep(0.8)

                # Re-detect state after clicking
                new_state = self.detect_arrow_state(pyautogui_module, location_tuple)
                self.append_log(f"{name} 第 {idx} 項點擊後狀態：{new_state}")

                # 如果點擊後仍無法檢測，再試一次
                if new_state is None:
                    self.append_log(f"{name} 第 {idx} 項點擊後仍無法檢測狀態，等待後重試...")
                    time.sleep(0.5)
                    new_state = self.detect_arrow_state(pyautogui_module, location_tuple)
                    self.append_log(f"{name} 第 {idx} 項重試後狀態：{new_state}")

                    # 如果還是無法檢測，可能需要再點一次
                    if new_state is None:
                        self.append_log(f"{name} 第 {idx} 項：⚠️ 仍無法檢測狀態，嘗試再次點擊")
                        pyautogui_module.click(click_x, click_y)
                        time.sleep(0.8)
                        new_state = self.detect_arrow_state(pyautogui_module, location_tuple)
                        self.append_log(f"{name} 第 {idx} 項第二次點擊後狀態：{new_state}")

                # 檢查狀態是否符合預期
                elif new_state != expectation:
                    self.append_log(f"{name} 第 {idx} 項狀態未符合預期（{new_state} != {expectation}），等待後重試檢測...")
                    time.sleep(0.5)
                    new_state = self.detect_arrow_state(pyautogui_module, location_tuple)
                    self.append_log(f"{name} 第 {idx} 項重試後狀態：{new_state}")

            except Exception as exc:
                self.append_log(f"{name} 第 {idx} 項切換失敗：{exc}")
                continue

            # Wait for UI to stabilize before processing next item
            if idx < len(all_locations):
                time.sleep(0.4)

        # 如果檢測失敗次數過多，給出警告
        if failed_detection_count > 0:
            self.append_log(f"{name}：⚠️ 有 {failed_detection_count} 項無法檢測箭頭狀態，可能需要調整模板或檢測區域")

        # 生成摘要
        total = len(all_locations)
        if processed_count == 0 and skipped_count == 0:
            summary = f"{name}: 未找到可處理的項目"
        elif processed_count == 0:
            summary = f"{name}: {total} 項已符合預期"
        else:
            summary = f"{name}: 已處理 {processed_count} 項，跳過 {skipped_count} 項（共 {total} 項）"

        self.append_log(summary)
        return summary

    def _ensure_section_state(
        self,
        pyautogui_module: Any,
        name: str,
        template: Path,
        expectation: str,
        state_text: dict[str, str],
        *,
        region_hint: tuple[int, int, int, int] | None = None,
    ) -> tuple[str | None, str | None, bool, tuple[int, int, int, int] | None]:
        expected_label = state_text.get(expectation, expectation)
        screen_width, screen_height = pyautogui_module.size()

        def expand(region: tuple[int, int, int, int] | None, px: int, py: int) -> tuple[int, int, int, int] | None:
            if region is None:
                return None
            x, y, w, h = region
            new_x = max(x - px, 0)
            new_y = max(y - py, 0)
            max_w = screen_width - new_x
            max_h = screen_height - new_y
            new_w = min(max_w, w + px * 2)
            new_h = min(max_h, h + py * 2)
            return (new_x, new_y, max(new_w, w), max(new_h, h))

        attempts: list[tuple[str, tuple[int, int, int, int] | None, float]] = [
            ("primary", region_hint, 0.90),
            ("expanded", expand(region_hint, 80, 60), 0.86),
            ("wider", expand(region_hint, 140, 100), 0.82),
            ("screen", (0, 0, screen_width, screen_height), 0.78),
        ]

        section_location = None
        section_coords: tuple[int, int, int, int] | None = None

        for label, region, conf in attempts:
            if region is None and label != "screen":
                continue
            location = self._try_locate(
                pyautogui_module,
                template,
                region=region,
                confidence=conf,
            )
            if location is not None:
                section_location = location
                section_coords = self._box_to_tuple(location)
                if label != "primary":
                    self.append_log(f"區塊 {name} 模板命中來源：{label} (conf={conf:.2f})")
                break

        if section_location is None:
            self.append_log(f"區塊 {name} 未命中模板。")
            summary = f"{name}: 未命中"
            return summary, f"{name}: 未命中", False, None

        toggled = False

        for _ in range(3):
            state, arrow_location = self._determine_section_state(pyautogui_module, section_location, expectation)
            actual_label = state_text.get(state, "未知")
            self.append_log(f"區塊 {name} 狀態：{actual_label}，預期：{expected_label}。")

            if state == expectation:
                summary = f"{name}: {actual_label} (符合)"
                return summary, None, toggled, section_coords

            coords = self._box_to_tuple(arrow_location)
            if coords is None and section_coords is not None:
                approx_x = section_coords[0] + section_coords[2] - 40
                approx_y = section_coords[1] + section_coords[3] // 2
                coords = (int(approx_x) - 14, int(approx_y) - 14, 28, 28)
                self.append_log(f"{name}: 使用估計箭頭位置 {coords}。")

            if coords is not None and section_coords is not None:
                section_bottom = section_coords[1] + section_coords[3]
                if coords[1] > section_bottom + 24:
                    self.append_log(f"{name}: 偵測到位於下方的箭頭 {coords}，忽略。")
                    arrow_location = None
                    coords = None
                    continue

            if coords is None:
                issue = f"{name}: 無法判定箭頭"
                return f"{name}: {actual_label} -> 未切換", issue, toggled, section_coords

            x = coords[0] + coords[2] / 2
            y = coords[1] + coords[3] / 2
            try:
                pyautogui_module.moveTo(x, y, duration=0.15)
            except Exception:
                pass
            try:
                pyautogui_module.click(x, y)
            except Exception as exc:
                issue = f"{name}: 切換失敗 {exc}"
                self.append_log(f"{name} 切換失敗：{exc}")
                return f"{name}: {actual_label} -> 切換失敗", issue, toggled, section_coords

            toggled = True
            time.sleep(0.5)
            section_location = (
                self._try_locate(pyautogui_module, template, region=section_coords, confidence=0.82) or section_location
            )
            section_coords = self._box_to_tuple(section_location)

        state, _ = self._determine_section_state(pyautogui_module, section_location, expectation)
        actual_label = state_text.get(state, "未知")
        issue = None
        summary = f"{name}: {actual_label} (符合)" if state == expectation else f"{name}: 切換後 {actual_label}"
        if state != expectation:
            issue = f"{name}: 切換後仍為 {actual_label}"
        return summary, issue, toggled, section_coords

    def _determine_section_state(
        self,
        pyautogui_module: Any,
        section_box: Any,
        expectation: str | None = None,
    ) -> tuple[str | None, Any]:
        region = self._section_arrow_region(section_box)
        order: list[tuple[str, Path]]
        if expectation == 'hide':
            order = [('hide', self.hide_arrow_template), ('show', self.show_arrow_template)]
        elif expectation == 'show':
            order = [('show', self.show_arrow_template), ('hide', self.hide_arrow_template)]
        else:
            order = [('show', self.show_arrow_template), ('hide', self.hide_arrow_template)]

        for state, template in order:
            arrow_box = self._locate_arrow(pyautogui_module, template, region, section_box)
            if arrow_box is not None:
                return state, arrow_box

        if expectation in {'hide', 'show'}:
            self.append_log(f"Arrow template missing; unable to confirm {expectation}.")
        return None, None

    def _section_arrow_region(self, section_box: Any) -> tuple[int, int, int, int] | None:
        coords = self._box_to_tuple(section_box)
        if coords is None:
            return None
        left, top, width, height = coords
        arrow_left = max(int(left + width) - 190, 0)
        arrow_top = max(int(top) - 15, 0)
        arrow_width = max(190, int(width) + 60)
        arrow_height = min(160, int(height) + 70)
        return (arrow_left, arrow_top, arrow_width, arrow_height)

    def detect_arrow_state(
        self,
        pyautogui_module: Any,
        anchor_box: Any,
    ) -> str | None:
        region = self._arrow_region(anchor_box)
        if region is None:
            self.append_log("無法計算箭頭搜索區域")
            return None

        # 輸出搜索區域以便調試
        box = self._box_to_tuple(anchor_box)
        if box:
            self.append_log(f"箭頭檢測區域: 錨點=({box[0]}, {box[1]}, {box[2]}x{box[3]}), 搜索=({region[0]}, {region[1]}, {region[2]}x{region[3]})")

        # Try to detect hide arrow (收合) first
        hide_box = self._locate_arrow(pyautogui_module, self.hide_arrow_template, region, anchor_box, save_debug_screenshot=False)
        if hide_box is not None:
            hide_coords = self._box_to_tuple(hide_box)
            self.append_log(f"✓ 檢測到收合箭頭 (hide) 於 ({hide_coords[0]}, {hide_coords[1]})")
            return 'hide'

        # Try to detect show arrow (展開)
        show_box = self._locate_arrow(pyautogui_module, self.show_arrow_template, region, anchor_box, save_debug_screenshot=False)
        if show_box is not None:
            show_coords = self._box_to_tuple(show_box)
            self.append_log(f"✓ 檢測到展開箭頭 (show) 於 ({show_coords[0]}, {show_coords[1]})")
            return 'show'

        # 兩種箭頭都找不到，保存截圖供調試
        self.append_log("✗ 未檢測到任何箭頭狀態（hide 和 show 模板皆未匹配）")
        self._save_debug_screenshot(pyautogui_module, region, "both_arrows")
        return None

    def _save_debug_screenshot(
        self,
        pyautogui_module: Any,
        region: tuple[int, int, int, int] | None,
        template_name: str,
    ) -> None:
        """保存搜索區域截圖供調試"""
        if region is None:
            return

        try:
            import time
            from pathlib import Path as PathLib
            screenshot = pyautogui_module.screenshot(region=region)
            debug_dir = PathLib("reports/arrow_debug")
            debug_dir.mkdir(parents=True, exist_ok=True)
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            debug_path = debug_dir / f"arrow_search_{template_name}_{timestamp}.png"
            screenshot.save(debug_path)
            self.append_log(f"已保存搜索區域截圖至 {debug_path} 供調試")
        except Exception as e:
            self.append_log(f"保存調試截圖失敗: {e}")

    def _arrow_region(
        self,
        anchor_box: Any,
    ) -> tuple[int, int, int, int] | None:
        """計算箭頭搜索區域，箭頭應該在區塊標題的右側"""
        box = self._box_to_tuple(anchor_box)
        if box is None:
            return None
        left, top, width, height = box

        # 箭頭通常在區塊標題右側
        # 從標題左側開始（給一點緩衝）到右側延伸100像素
        region_left = max(int(left) - 50, 0)
        # 上下給一些緩衝空間
        region_top = max(int(top) - 10, 0)
        # 寬度：覆蓋整個標題寬度 + 右側延伸100像素
        region_width = int(width) + 150
        # 高度：標題高度 + 上下緩衝20像素
        region_height = int(height) + 20

        return (region_left, region_top, region_width, region_height)

    def _locate_arrow(
        self,
        pyautogui_module: Any,
        template_path: Path,
        region: tuple[int, int, int, int] | None,
        anchor_box: Any,
        save_debug_screenshot: bool = True,
    ) -> Any:
        anchor_tuple = self._box_to_tuple(anchor_box)
        screen_width, screen_height = pyautogui_module.size()

        def clip(bounds: tuple[int, int, int, int] | None) -> tuple[int, int, int, int] | None:
            if bounds is None:
                return None
            x, y, w, h = bounds
            x = max(x, 0)
            y = max(y, 0)
            w = min(w, screen_width - x)
            h = min(h, screen_height - y)
            if w <= 0 or h <= 0:
                return None
            return (x, y, w, h)

        def is_valid_arrow_position(arrow_loc: Any, anchor: tuple[int, int, int, int]) -> bool:
            """驗證箭頭位置是否在錨點附近（合理範圍內）"""
            arrow_coords = self._box_to_tuple(arrow_loc)
            if arrow_coords is None:
                return False

            arrow_x, arrow_y, arrow_w, arrow_h = arrow_coords
            anchor_x, anchor_y, anchor_w, anchor_h = anchor

            # 箭頭應該在錨點的水平方向附近（±200px）和垂直方向附近（±100px）
            horizontal_distance = abs(arrow_x - anchor_x)
            vertical_distance = abs(arrow_y - anchor_y)

            # 允許的最大距離
            max_horizontal = 350  # 箭頭可能在標題右側
            max_vertical = 80     # 箭頭應該與標題在同一高度

            is_valid = horizontal_distance <= max_horizontal and vertical_distance <= max_vertical

            if not is_valid:
                self.append_log(
                    f"✗ 箭頭位置驗證失敗: 箭頭({arrow_x}, {arrow_y}) 距離錨點({anchor_x}, {anchor_y}) "
                    f"水平 {horizontal_distance}px (限制 {max_horizontal}px), "
                    f"垂直 {vertical_distance}px (限制 {max_vertical}px)"
                )

            return is_valid

        # 只使用原始搜索區域，不進行過度擴張或全螢幕搜索
        regions: list[tuple[int, int, int, int] | None] = [clip(region)]

        # 降低信心度閾值，提高檢測成功率
        confidence_levels = [0.85, 0.80, 0.75, 0.70, 0.65, 0.60]
        seen: set[tuple[int, int, int, int]] = set()

        for search_region in regions:
            if search_region is None or search_region in seen:
                continue
            seen.add(search_region)
            for conf in confidence_levels:
                loc = self._try_locate(
                    pyautogui_module,
                    template_path,
                    region=search_region,
                    confidence=conf,
                )
                if loc is not None:
                    # 驗證箭頭位置是否在錨點附近
                    if anchor_tuple is not None and not is_valid_arrow_position(loc, anchor_tuple):
                        continue  # 位置不合理，繼續搜索

                    if conf < 0.85:
                        self.append_log(f"模板 {template_path.name} 使用降級信心 {conf:.2f} 命中。")
                    return loc

        # OpenCV 備用方案，進一步降低閾值
        for search_region in regions:
            if search_region is None:
                continue
            loc = self._match_template_cv(pyautogui_module, template_path, search_region, threshold=0.55)
            if loc is not None:
                # 驗證箭頭位置是否在錨點附近
                if anchor_tuple is not None and not is_valid_arrow_position(loc, anchor_tuple):
                    continue  # 位置不合理，繼續搜索

                self.append_log(f"模板 {template_path.name} 以 OpenCV 灰階比對命中（閾值 0.55）。")
                return loc

        # 如果仍然失敗，返回 None（不保存截圖，由上層決定）
        return None

    def _match_template_cv(
        self,
        pyautogui_module: Any,
        template_path: Path,
        region: tuple[int, int, int, int],
        threshold: float = 0.78,
    ) -> tuple[int, int, int, int] | None:
        try:
            screenshot = pyautogui_module.screenshot(region=region)
        except Exception as exc:
            self.append_log(f"截圖區域 {region} 失敗：{exc}")
            return None

        shot = np.array(screenshot)
        if shot.ndim == 2:
            shot_gray = shot
        else:
            shot_gray = cv2.cvtColor(shot, cv2.COLOR_RGB2GRAY)

        tpl = cv2.imread(str(template_path), cv2.IMREAD_UNCHANGED)
        if tpl is None:
            self.append_log(f"讀取模板 {template_path.name} 失敗。")
            return None
        if tpl.ndim == 3:
            tpl_gray = cv2.cvtColor(tpl, cv2.COLOR_BGR2GRAY)
        else:
            tpl_gray = tpl

        if shot_gray.shape[0] < tpl_gray.shape[0] or shot_gray.shape[1] < tpl_gray.shape[1]:
            return None

        result = cv2.matchTemplate(shot_gray, tpl_gray, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            x, y = max_loc
            w, h = tpl_gray.shape[1], tpl_gray.shape[0]
            return (region[0] + x, region[1] + y, w, h)
        return None

    def _box_to_tuple(self, box: Any) -> tuple[int, int, int, int] | None:
        if box is None:
            return None
        if hasattr(box, "left"):
            return int(box.left), int(box.top), int(box.width), int(box.height)
        if isinstance(box, (tuple, list)) and len(box) >= 4:
            return int(box[0]), int(box[1]), int(box[2]), int(box[3])
        return None

    def _try_locate(
        self,
        pyautogui_module: Any,
        template_path: Path,
        *,
        region: tuple[int, int, int, int] | None = None,
        confidence: float = TEMPLATE_CONFIDENCE,
    ) -> Any:
        if not template_path.exists():
            return None

        base_kwargs: dict[str, Any] = {}
        if region is not None:
            base_kwargs['region'] = region

        confidence_list: list[float | None] = []
        if confidence is not None:
            confidence_list.append(confidence)
        confidence_list.append(None)

        for conf in confidence_list:
            for use_grayscale in (False, True):
                kwargs = dict(base_kwargs)
                if conf is not None:
                    kwargs['confidence'] = max(min(conf, 0.98), 0.55)
                if use_grayscale:
                    kwargs['grayscale'] = True

                try:
                    return pyautogui_module.locateOnScreen(str(template_path), **kwargs)
                except TypeError:
                    kwargs.pop('confidence', None)
                    try:
                        return pyautogui_module.locateOnScreen(str(template_path), **kwargs)
                    except Exception:
                        # 靜默處理，這是正常的搜索過程
                        pass
                except Exception:
                    # 靜默處理，這是正常的搜索過程
                    pass

        return None

    def _try_locate_all(
        self,
        pyautogui_module: Any,
        template_path: Path,
        *,
        region: tuple[int, int, int, int] | None = None,
        confidence: float = TEMPLATE_CONFIDENCE,
    ) -> list[Any]:
        """查找所有匹配的模板位置"""
        if not template_path.exists():
            return []

        base_kwargs: dict[str, Any] = {}
        if region is not None:
            base_kwargs['region'] = region

        results: list[Any] = []

        # 嘗試使用 locateAllOnScreen
        confidence_list: list[float | None] = []
        if confidence is not None:
            confidence_list.append(confidence)
        confidence_list.append(None)

        for conf in confidence_list:
            for use_grayscale in (False, True):
                kwargs = dict(base_kwargs)
                if conf is not None:
                    kwargs['confidence'] = max(min(conf, 0.98), 0.55)
                if use_grayscale:
                    kwargs['grayscale'] = True

                try:
                    locations = pyautogui_module.locateAllOnScreen(str(template_path), **kwargs)
                    # 將生成器轉換為列表
                    results = list(locations)
                    if results:
                        self.append_log(f"使用 locateAllOnScreen 找到 {len(results)} 個匹配項")
                        return results
                except TypeError:
                    # locateAllOnScreen 可能不支持 confidence 參數
                    kwargs.pop('confidence', None)
                    try:
                        locations = pyautogui_module.locateAllOnScreen(str(template_path), **kwargs)
                        results = list(locations)
                        if results:
                            self.append_log(f"使用 locateAllOnScreen（無 confidence）找到 {len(results)} 個匹配項")
                            return results
                    except Exception:
                        pass
                except Exception:
                    pass

        # 如果 locateAllOnScreen 失敗，回退到單次查找
        single_location = self._try_locate(
            pyautogui_module,
            template_path,
            region=region,
            confidence=confidence,
        )
        if single_location is not None:
            self.append_log(f"locateAllOnScreen 失敗，使用單次查找找到 1 個匹配項")
            return [single_location]

        return []

    def _template_paths(self) -> Iterable[Path]:
        yield self.friend_list_template
        yield self.message_cube_template
        for _, path, _ in self.arrow_section_templates:
            yield path
        yield self.hide_arrow_template
        yield self.show_arrow_template

    def _focus_line_window(self, pyautogui_module: Any) -> bool:
        try:
            windows = pyautogui_module.getWindowsWithTitle("LINE")
        except Exception as exc:
            self.append_log(f"取得 LINE 視窗失敗：{exc}")
            return False

        if not windows:
            self.append_log("未找到 LINE 視窗。")
            return False

        window = windows[0]
        try:
            if getattr(window, "isMinimized", False):
                window.restore()
            window.activate()
            time.sleep(0.3)
        except Exception as exc:
            self.append_log(f"聚焦 LINE 失敗：{exc}")
        return True

    def _on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("關閉程式", "流程仍在執行，確定要終止並關閉嗎？"):
                return
            self.handle_stop()
            if self.worker_thread is not None and self.worker_thread.is_alive():
                self.worker_thread.join(timeout=2.0)
        self.root.destroy()

def launch_ui() -> None:
    root = tk.Tk()
    app = AutoaApp(root)
    root.geometry("820x900")
    root.mainloop()


if __name__ == "__main__":
    launch_ui()
