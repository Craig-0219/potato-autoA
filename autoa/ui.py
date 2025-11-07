"""Tkinter UI for AUTOA RPA control panel."""
from __future__ import annotations

import csv
import locale
import os
import random
import subprocess
import sys
import threading
import time
from collections import deque
import cv2
import numpy as np

from pathlib import Path
from typing import Any, Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from autoa.line_automation import CycleResult, LineAutomationError, cycle_friend_chats


def get_resource_path(relative_path: str) -> Path:
    """獲取資源檔案的絕對路徑（支援 PyInstaller 打包）"""
    try:
        # PyInstaller 創建臨時資料夾，將路徑存儲在 _MEIPASS 中
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        # 開發環境中使用當前工作目錄
        base_path = Path.cwd()
    return base_path / relative_path


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


class AutoaApp:
    """Main Tkinter UI."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("AUTOA RPA Control Panel")

        self.friend_count_var = tk.StringVar(value="10")  # 發送好友數量
        self.delay_var = tk.StringVar(value="2")  # 每個好友延遲秒數
        self.message_text: tk.Text | None = None
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

        self.log_text: tk.Text | None = None
        self.progress_bar: ttk.Progressbar | None = None
        self.log_lines: deque[str] = deque(maxlen=LOG_CAPACITY)

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

        self.friend_list_template = get_resource_path("templates/friend-list.png")
        self.message_cube_template = get_resource_path("templates/message_cube.png")
        self.greenchat_template = get_resource_path("templates/greenchat.png")  # 新增：綠色聊天框模板
        self.hide_arrow_template = get_resource_path("templates/hide.png")
        self.show_arrow_template = get_resource_path("templates/show.png")

        self.notebook: ttk.Notebook | None = None

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
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

        container = ttk.Frame(main_tab, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)

        self._build_recipient_section(container)
        self._build_message_section(container)
        self._build_system_section(container)
        self._build_execution_section(container)

        test_frame = ttk.Frame(test_tab, padding=12)
        test_frame.grid(row=0, column=0, sticky="nsew")
        test_tab.columnconfigure(0, weight=1)
        test_tab.rowconfigure(0, weight=1)
        self._build_test_shortcuts_section(test_frame)

    def _build_recipient_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="好友設定", padding=10)
        frame.grid(row=0, column=0, sticky="ew")
        frame.columnconfigure(1, weight=1)

        # 發送好友數量設定
        ttk.Label(frame, text="發送好友數量：").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.friend_count_var, width=10).grid(row=0, column=1, sticky="w")
        ttk.Label(frame, text="（每次處理的好友數量）").grid(row=0, column=2, sticky="w", padx=(8, 0))

        # 延遲設置
        ttk.Label(frame, text="延遲設置：").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=self.delay_var, width=10).grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Label(frame, text="秒（每個好友之間的延遲）").grid(row=1, column=2, sticky="w", pady=(8, 0), padx=(8, 0))

        # 執行模式設置
        ttk.Label(frame, text="執行模式：").grid(row=2, column=0, sticky="w", pady=(8, 0))
        mode_frame = ttk.Frame(frame)
        mode_frame.grid(row=2, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Radiobutton(mode_frame, text="乾跑（不實際發送）", value="dryrun", variable=self.mode_var).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(mode_frame, text="正式（實際發送）", value="live", variable=self.mode_var).pack(side="left")

    def _build_message_section(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="訊息內容", padding=10)
        frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        frame.columnconfigure(0, weight=1)

        self.message_text = tk.Text(frame, wrap="word", height=6)
        self.message_text.grid(row=0, column=0, columnspan=3, sticky="nsew")
        frame.rowconfigure(0, weight=1)

        ttk.Label(frame, text="附件圖片檔案：").grid(row=1, column=0, sticky="w", pady=(8, 0))
        image_entry = ttk.Entry(frame, textvariable=self.image_var)
        image_entry.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        ttk.Button(frame, text="選擇圖片...", command=self.browse_image).grid(row=1, column=2, pady=(8, 0))

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

        self.log_text = tk.Text(frame, height=10, wrap="word", state="disabled")
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
            # 使用 CREATE_NO_WINDOW 避免 CMD 視窗閃動
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq line.exe"],
                capture_output=True,
                text=True,
                check=False,
                creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0,
            )
        except Exception:
            return False
        return "line.exe" in result.stdout.lower()

    # ------------------------------------------------------------------
    def handle_start(self) -> None:
        if self.running:
            messagebox.showinfo("執行中", "流程已在執行。")
            return

        # 不再需要收件者驗證，直接使用好友列表
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

        # 呼叫 LINE 視窗到最上層
        try:
            import pyautogui
        except ImportError as exc:
            messagebox.showerror("測試失敗", f"無法載入 pyautogui：{exc}")
            return

        if not self._focus_line_window(pyautogui):
            messagebox.showwarning("執行失敗", "未偵測到 LINE 視窗，請確認 LINE 已開啟。")
            return

        self.stop_event = threading.Event()
        self.running = True
        self.paused = False

        if self.pause_button is not None:
            self.pause_button.configure(text="暫停")

        self._toggle_buttons(running=True)
        self.append_log(f"開始執行流程{'（乾跑模式）' if dry_run else ''}。")

        # 不再使用收件者，使用空字符串
        recipient = ""
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
        """主流程：批量發送訊息給好友列表"""
        try:
            import pyautogui
        except ImportError as exc:
            self.append_log(f"無法載入 pyautogui：{exc}")
            self.root.after(0, lambda: messagebox.showerror('執行失敗', f'無法載入 pyautogui：{exc}'))
            self.root.after(0, lambda: self._on_worker_finished(False))
            return

        try:
            # 獲取發送好友數量和延遲設置
            try:
                friend_count = int(self.friend_count_var.get())
                delay = float(self.delay_var.get())
            except ValueError:
                self.append_log("參數錯誤：好友數量或延遲設置無效")
                self.root.after(0, lambda: self._on_worker_finished(False))
                return

            self._set_current_step("準備階段")
            self._set_progress(5.0)

            # 1. 聚焦 LINE 視窗
            self.append_log("步驟 1/3：聚焦 LINE 視窗")
            if not self._focus_line_window(pyautogui):
                self.append_log("未偵測到 LINE 視窗")
                self.root.after(0, lambda: messagebox.showwarning('執行失敗', '未偵測到 LINE 視窗。'))
                self.root.after(0, lambda: self._on_worker_finished(False))
                return

            time.sleep(0.5)

            # 2. 執行箭頭校正（只有好友列表展開）
            self.append_log("步驟 2/3：校正側邊欄（只展開好友列表）")
            self._set_progress(15.0)
            if not self._calibrate_arrows_for_friend_only(pyautogui):
                self.append_log("箭頭校正失敗")
                self.root.after(0, lambda: messagebox.showwarning('執行失敗', '箭頭校正失敗，請確認側邊欄可見。'))
                self.root.after(0, lambda: self._on_worker_finished(False))
                return

            time.sleep(0.5)

            # 3. 找到好友標題並點擊第一個好友
            self.append_log("步驟 3/3：定位第一個好友")
            self._set_progress(25.0)
            friend_template = get_resource_path("templates/friend.png")
            friend_location = self._try_locate(pyautogui, friend_template, confidence=0.88)
            if friend_location is None:
                self.append_log("未找到好友標題")
                self.root.after(0, lambda: messagebox.showwarning('執行失敗', '未找到好友標題，請確認好友區塊已展開。'))
                self.root.after(0, lambda: self._on_worker_finished(False))
                return

            friend_coords = self._box_to_tuple(friend_location)
            if friend_coords is None:
                self.append_log("好友標題位置解析失敗")
                self.root.after(0, lambda: self._on_worker_finished(False))
                return

            first_friend_x = friend_coords[0] + friend_coords[2] // 2 + 40  # 往右移動 40 像素
            first_friend_y = friend_coords[1] + friend_coords[3] + 20

            pyautogui.click(first_friend_x, first_friend_y)
            self.append_log(f"已選中第一個好友於 ({first_friend_x}, {first_friend_y})")
            time.sleep(0.5)

            # 確保 LINE 視窗保持焦點
            self._ensure_line_focus(pyautogui)
            time.sleep(0.3)

            # 4. 開始循環處理每個好友
            self._set_current_step("發送訊息中")
            sent_count = 0
            clicked_count = 0

            for idx in range(friend_count):
                if self.stop_event.is_set():
                    self.append_log("用戶中止流程")
                    break

                if self._wait_if_paused():
                    break

                current_num = idx + 1
                progress = 25.0 + (70.0 * current_num / friend_count)
                self._set_progress(progress)

                self.append_log(f"\n處理第 {current_num}/{friend_count} 位好友...")

                # 確保 LINE 視窗有焦點
                self._ensure_line_focus(pyautogui)
                time.sleep(0.3)

                # 檢測 greenchat.png
                greenchat_location = self._try_locate(pyautogui, self.greenchat_template, confidence=0.95)

                if greenchat_location:
                    # 有綠色按鈕，需要點擊打開聊天窗口
                    greenchat_coords = self._box_to_tuple(greenchat_location)
                    if greenchat_coords:
                        # 點擊前再次確保焦點
                        self._ensure_line_focus(pyautogui)
                        time.sleep(0.2)

                        click_x = greenchat_coords[0] + greenchat_coords[2] // 2
                        click_y = greenchat_coords[1] + greenchat_coords[3] // 2

                        self.append_log(f"  → 檢測到綠色按鈕，點擊開啟聊天窗口")
                        pyautogui.click(click_x, click_y)
                        time.sleep(1.2)

                        # 點擊後再次確保焦點
                        self._ensure_line_focus(pyautogui)
                        time.sleep(0.2)

                        # 驗證按鈕消失
                        check_location = self._try_locate(pyautogui, self.greenchat_template, confidence=0.90)
                        if not check_location:
                            clicked_count += 1
                            self.append_log(f"  ✓ 聊天窗口已開啟")
                        else:
                            self.append_log(f"  ⚠ 開啟可能失敗")

                # 發送訊息
                if self._send_message_to_current_chat(pyautogui, message, image_path, dry_run):
                    sent_count += 1
                    self.append_log(f"  ✓ 訊息已發送（累計 {sent_count} 次）")
                else:
                    self.append_log(f"  ✗ 訊息發送失敗")

                # 延遲後切換到下一個好友
                if current_num < friend_count:
                    self.append_log(f"  → 延遲 {delay} 秒後切換到下一位好友")
                    time.sleep(delay)

                    # 重要：點擊好友列表區域以恢復焦點
                    # 使用第一個好友的位置作為參考點
                    focus_x = first_friend_x
                    focus_y = first_friend_y + (current_num - 1) * 50  # 估算當前好友的Y位置
                    pyautogui.click(focus_x, focus_y)
                    self.append_log(f"  → 重新聚焦到好友列表")
                    time.sleep(0.3)

                    # 按 DOWN 鍵切換到下一個好友
                    pyautogui.press('down')
                    time.sleep(0.3)

            # 5. 完成
            self._set_progress(100.0)
            self._set_current_step("完成")
            self.append_log(f"\n流程完成！處理了 {friend_count} 位好友，發送了 {sent_count} 次訊息")
            self.root.after(0, lambda: messagebox.showinfo('流程完成',
                f'已處理 {friend_count} 位好友\n成功發送 {sent_count} 次訊息\n點擊綠色按鈕 {clicked_count} 次'))
            self.root.after(0, lambda: self._on_worker_finished(True))

        except Exception as exc:
            self.append_log(f'流程發生錯誤：{exc}')
            import traceback
            self.append_log(traceback.format_exc())
            self.root.after(0, lambda err=exc: messagebox.showerror('執行失敗', f'執行失敗：{err}'))
            self.root.after(0, lambda: self._on_worker_finished(False))

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
        """依序開啟聊天窗測試（新實現）"""
        if self.friend_cycle_thread and self.friend_cycle_thread.is_alive():
            messagebox.showinfo('聊天測試', '聊天測試正在進行中，請稍候。')
            return

        # 獲取要發送的好友數量
        try:
            friend_count = int(self.friend_count_var.get())
            if friend_count <= 0:
                messagebox.showwarning('好友數量', '請輸入大於 0 的數量。')
                return
        except ValueError:
            messagebox.showwarning('好友數量', '請輸入有效的數字。')
            return

        # 獲取延遲設置
        try:
            delay = float(self.delay_var.get())
            if delay < 0:
                messagebox.showwarning('延遲設置', '延遲不能為負數。')
                return
        except ValueError:
            messagebox.showwarning('延遲設置', '請輸入有效的數字。')
            return

        self.append_log(f'開始依序開啟聊天窗測試：目標處理 {friend_count} 位好友，每個好友延遲 {delay} 秒')

        thread = threading.Thread(
            target=self._cycle_friend_chats_worker_new,
            args=(friend_count, delay),
            daemon=True,
        )
        self.friend_cycle_thread = thread
        thread.start()

    def _cycle_friend_chats_worker_new(self, friend_count: int, delay: float) -> None:
        """依序開啟聊天窗的 Worker 方法"""
        try:
            import pyautogui
        except ImportError as exc:
            self.append_log(f"無法載入 pyautogui：{exc}")
            self.root.after(0, lambda: messagebox.showerror('測試失敗', f'無法載入 pyautogui：{exc}'))
            return

        try:
            # 1. 聚焦 LINE 視窗
            if not self._focus_line_window(pyautogui):
                self.append_log("未偵測到 LINE 視窗")
                self.root.after(0, lambda: messagebox.showwarning('測試失敗', '未偵測到 LINE 視窗。'))
                return

            # 2. 找到好友標題位置
            friend_template = get_resource_path("templates/friend.png")
            friend_location = self._try_locate(pyautogui, friend_template, confidence=0.88)
            if friend_location is None:
                self.append_log("未找到好友標題")
                self.root.after(0, lambda: messagebox.showwarning('測試失敗', '未找到好友標題，請確認好友區塊已展開。'))
                return

            friend_coords = self._box_to_tuple(friend_location)
            if friend_coords is None:
                self.append_log("好友標題位置解析失敗")
                return

            # 3. 計算第一個好友的點擊位置（標題下方一點點）
            first_friend_x = friend_coords[0] + friend_coords[2] // 2 + 40  # 往右移動 40 像素
            first_friend_y = friend_coords[1] + friend_coords[3] + 20  # 標題下方 20 像素

            self.append_log(f"準備點擊第一個好友位置：({first_friend_x}, {first_friend_y})")

            # 4. 點擊第一個好友選中它（不會打開聊天窗）
            pyautogui.click(first_friend_x, first_friend_y)
            time.sleep(0.5)

            # 確保 LINE 視窗保持焦點
            self._ensure_line_focus(pyautogui)
            time.sleep(0.3)

            opened_count = 0
            clicked_count = 0  # 記錄點擊 greenchat 的次數
            max_attempts = friend_count * 3  # 最多嘗試 3 倍數量，防止無限循環

            # 5. 開始循環檢測和導航
            for attempt in range(max_attempts):
                # 每個好友都算進度
                opened_count += 1
                self.append_log(f"檢測第 {opened_count} 位好友（共 {friend_count} 位）...")

                # 確保 LINE 視窗有焦點
                self._ensure_line_focus(pyautogui)

                # 短暫等待讓 UI 穩定
                time.sleep(0.3)

                # 檢測是否有 greenchat.png（使用極高信心度避免誤判）
                greenchat_location = self._try_locate(pyautogui, self.greenchat_template, confidence=0.95)

                if greenchat_location:
                    # 保存匹配區域的截圖進行驗證
                    greenchat_coords = self._box_to_tuple(greenchat_location)
                    if greenchat_coords:
                        # 計算 greenchat.png 按鈕的中心點
                        click_x = greenchat_coords[0] + greenchat_coords[2] // 2
                        click_y = greenchat_coords[1] + greenchat_coords[3] // 2

                        try:
                            # 截取並保存匹配區域
                            screenshot = pyautogui.screenshot()
                            matched_region = screenshot.crop((
                                greenchat_coords[0],
                                greenchat_coords[1],
                                greenchat_coords[0] + greenchat_coords[2],
                                greenchat_coords[1] + greenchat_coords[3]
                            ))
                            matched_path = Path(f'debug_matched_region_{opened_count}.png')
                            matched_region.save(str(matched_path))

                            self.append_log(
                                f"第 {opened_count} 位好友：檢測到 greenchat.png "
                                f"位置=({greenchat_coords[0]}, {greenchat_coords[1]}), "
                                f"大小=({greenchat_coords[2]}x{greenchat_coords[3]})"
                            )
                            self.append_log(f"  → 匹配區域: {matched_path}")

                        except Exception as e:
                            self.append_log(f"  → 保存調試截圖失敗: {e}")

                        # 點擊前確保焦點
                        self._ensure_line_focus(pyautogui)
                        time.sleep(0.2)

                        # 點擊按鈕
                        pyautogui.click(click_x, click_y)
                        self.append_log(f"  → 點擊於 ({click_x}, {click_y})")
                        time.sleep(1.2)  # 增加延遲，等待聊天窗口打開

                        # 點擊後確保焦點
                        self._ensure_line_focus(pyautogui)
                        time.sleep(0.2)

                        # 檢查是否成功（按鈕消失）
                        check_location = self._try_locate(pyautogui, self.greenchat_template, confidence=0.90)
                        if not check_location:
                            clicked_count += 1
                            self.append_log(f"  ✓ 按鈕已消失，點擊成功！（累計 {clicked_count} 次）")
                        else:
                            self.append_log(f"  ⚠ 按鈕仍在，可能是誤判（請檢查 {matched_path}）")
                else:
                    # 沒有綠色聊天框（可能已經開啟過）
                    self.append_log(f"第 {opened_count} 位好友：無綠色聊天框")

                # 檢查是否已達目標數量
                if opened_count >= friend_count:
                    self.append_log(f"已完成 {opened_count} 位好友的處理（點擊 greenchat {clicked_count} 次）")
                    break

                # 使用方向鍵下移到下一個好友
                pyautogui.press('down')
                time.sleep(delay)

            # 6. 完成報告
            if opened_count >= friend_count:
                self.root.after(0, lambda: messagebox.showinfo('測試完成', f'已處理 {opened_count} 位好友，點擊 greenchat {clicked_count} 次。'))
            else:
                self.root.after(0, lambda: messagebox.showwarning('測試完成', f'僅處理 {opened_count} 位好友（目標：{friend_count}），點擊 greenchat {clicked_count} 次。'))

        except Exception as exc:
            self.append_log(f'聊天測試發生錯誤：{exc}')
            self.root.after(0, lambda err=exc: messagebox.showerror('測試失敗', f'執行失敗：{err}'))
        finally:
            self.friend_cycle_thread = None

    def _calibrate_arrows_for_friend_only(self, pyautogui_module: Any) -> bool:
        """執行箭頭校正：先收合所有箭頭，再展開好友（最下面的一個）"""
        try:
            screen_width, screen_height = pyautogui_module.size()
        except Exception:
            screen_width = screen_height = None

        self.append_log("開始箭頭校正：先收合所有箭頭，再展開最下面的一個（好友區塊）")

        # 獲取 LINE 視窗範圍
        line_region = self._get_line_window_region(pyautogui_module)
        if line_region:
            left, top, width, height = line_region
            self.append_log(f"  → LINE 視窗範圍：X={left}, Y={top}, 寬={width}, 高={height}")
        else:
            self.append_log("  → 無法獲取 LINE 視窗範圍")
            left, top, width, height = 0, 0, screen_width if screen_width else 1920, screen_height if screen_height else 1080

        # 先滾動到最上方，確保所有箭頭都在可見範圍內
        self.append_log("  → 滾動左側面板到頂部")
        self._scroll_left_panel_to_top(pyautogui_module)
        time.sleep(0.5)  # 等待滾動完成

        # 動態偵測左側面板寬度（使用 friend-list 按鈕位置）
        friend_list_location = self._try_locate(pyautogui_module, self.friend_list_template, confidence=0.88)
        if friend_list_location:
            coords = self._box_to_tuple(friend_list_location)
            if coords:
                # 左側面板寬度 = friend-list 按鈕的右邊界 + 1000 像素餘量（擴大搜尋範圍）
                left_panel_width = int(coords[0] + coords[2] + 1000) - left
                self.append_log(f"  → 動態偵測左側面板寬度：{left_panel_width} 像素")
            else:
                left_panel_width = 1100
                self.append_log(f"  → 無法解析 friend-list 位置，使用預設寬度：{left_panel_width} 像素")
        else:
            left_panel_width = 1100
            self.append_log(f"  → 未偵測到 friend-list 按鈕，使用預設寬度：{left_panel_width} 像素")

        # 限制搜尋範圍：只搜尋 LINE 視窗的左側面板（排除底部 80 像素避免抓到工作列）
        bottom_margin = 80  # 預留底部餘量，避免抓到工作列
        search_height = max(height - bottom_margin, 400)  # 至少保留 400 像素高度
        target_region = (left, top, left_panel_width, search_height)
        self.append_log(f"  → 箭頭搜尋範圍：X={left}, Y={top}, 寬={left_panel_width}, 高={search_height}（已排除底部 {bottom_margin} 像素）")

        # === 初始診斷：檢查當前箭頭狀態 ===
        self.append_log("  → 初始診斷：檢查當前箭頭狀態")

        initial_show = self._try_locate_all(
            pyautogui_module,
            self.show_arrow_template,
            region=target_region,
            confidence=0.85
        )
        initial_hide = self._try_locate_all(
            pyautogui_module,
            self.hide_arrow_template,
            region=target_region,
            confidence=0.85
        )

        initial_show_count = len(initial_show) if initial_show else 0
        initial_hide_count = len(initial_hide) if initial_hide else 0

        self.append_log(f"    初始狀態：{initial_show_count} 個展開箭頭、{initial_hide_count} 個收合箭頭")

        if initial_show and initial_show_count > 0:
            self.append_log(f"    展開箭頭位置：")
            for i, loc in enumerate(initial_show, 1):
                coords = self._box_to_tuple(loc)
                if coords:
                    self.append_log(f"      {i}. Y={int(coords[1])}, X={int(coords[0])}")

        if initial_hide and initial_hide_count > 0:
            self.append_log(f"    收合箭頭位置：")
            for i, loc in enumerate(initial_hide, 1):
                coords = self._box_to_tuple(loc)
                if coords:
                    self.append_log(f"      {i}. Y={int(coords[1])}, X={int(coords[0])}")

        # === 第一階段：收合所有展開的箭頭 ===
        self.append_log("  階段 1：收合所有展開的箭頭")
        max_collapse_attempts = 8  # 最多嘗試 8 次收合

        for attempt in range(1, max_collapse_attempts + 1):
            # 尋找所有展開箭頭
            show_locations = self._try_locate_all(
                pyautogui_module,
                self.show_arrow_template,
                region=target_region,
                confidence=0.85
            )

            show_count = len(show_locations) if show_locations else 0

            if show_count == 0:
                self.append_log(f"  ✓ 所有箭頭已收合")
                break

            self.append_log(f"    第 {attempt} 次：偵測到 {show_count} 個展開箭頭")

            # 顯示位置
            for i, loc in enumerate(show_locations, 1):
                coords = self._box_to_tuple(loc)
                if coords:
                    self.append_log(f"      展開箭頭 {i} 位於 Y={int(coords[1])}")

            # 收合第一個（最上面的）展開箭頭
            show_locations_sorted = sorted(
                show_locations,
                key=lambda loc: self._box_to_tuple(loc)[1] if self._box_to_tuple(loc) else float('inf')
            )

            first_show = show_locations_sorted[0]
            location_tuple = self._box_to_tuple(first_show)

            if location_tuple:
                click_x = location_tuple[0] + location_tuple[2] / 2
                click_y = location_tuple[1] + location_tuple[3] / 2

                self.append_log(f"    → 點擊收合箭頭於 ({int(click_x)}, {int(click_y)})")
                pyautogui_module.click(click_x, click_y)
                time.sleep(0.8)  # 等待動畫完成

        # === 第二階段：確認有足夠的收合箭頭 ===
        self.append_log("  階段 2：確認收合箭頭數量")

        # 檢查模板檔案是否存在
        if not self.hide_arrow_template.exists():
            self.append_log(f"  ✗ hide.png 模板檔案不存在：{self.hide_arrow_template}")
            return False

        # 先嘗試標準信心值 0.85
        hide_locations = self._try_locate_all(
            pyautogui_module,
            self.hide_arrow_template,
            region=target_region,
            confidence=0.85
        )

        hide_count = len(hide_locations) if hide_locations else 0
        self.append_log(f"    偵測到 {hide_count} 個收合箭頭（信心值 0.85）")

        # 如果找不到，嘗試降低信心值
        if hide_count == 0:
            self.append_log(f"    → 嘗試降低信心值到 0.7 重新搜尋")
            hide_locations = self._try_locate_all(
                pyautogui_module,
                self.hide_arrow_template,
                region=target_region,
                confidence=0.7
            )
            hide_count = len(hide_locations) if hide_locations else 0
            self.append_log(f"    偵測到 {hide_count} 個收合箭頭（信心值 0.7）")

        # 如果還是找不到，再嘗試更低的信心值
        if hide_count == 0:
            self.append_log(f"    → 嘗試降低信心值到 0.6 重新搜尋")
            hide_locations = self._try_locate_all(
                pyautogui_module,
                self.hide_arrow_template,
                region=target_region,
                confidence=0.6
            )
            hide_count = len(hide_locations) if hide_locations else 0
            self.append_log(f"    偵測到 {hide_count} 個收合箭頭（信心值 0.6）")

        # 顯示位置
        if hide_locations:
            for i, loc in enumerate(hide_locations, 1):
                coords = self._box_to_tuple(loc)
                if coords:
                    self.append_log(f"      收合箭頭 {i} 位於 Y={int(coords[1])}")

        if hide_count == 0:
            self.append_log(f"  ✗ 沒有偵測到任何收合箭頭")
            self.append_log(f"  → 可能原因：")
            self.append_log(f"     1. hide.png 模板無法匹配 LINE 中的收合箭頭樣式")
            self.append_log(f"     2. LINE 視窗被其他視窗遮擋")
            self.append_log(f"     3. 螢幕縮放比例不是 100%")
            self.append_log(f"  → 解決方法：")
            self.append_log(f"     1. 重新截取 hide.png：只截取收合箭頭（→）本身，約 20x20 像素")
            self.append_log(f"     2. 確保 LINE 視窗完全可見且沒有被遮擋")
            self.append_log(f"     3. 檢查 Windows 顯示設定，確認縮放比例為 100%")

            # 截圖診斷
            try:
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_path = f"debug_line_window_{timestamp}.png"
                if target_region:
                    # 截取 LINE 視窗範圍
                    screenshot = pyautogui_module.screenshot(region=target_region)
                else:
                    # 截取全螢幕
                    screenshot = pyautogui_module.screenshot()
                screenshot.save(screenshot_path)
                self.append_log(f"  → 已保存診斷截圖：{screenshot_path}")
                self.append_log(f"     請檢查截圖，確認 LINE 側邊欄的箭頭是否可見")
            except Exception as exc:
                self.append_log(f"  → 保存診斷截圖失敗：{exc}")

            return False

        self.append_log(f"  ✓ 找到 {hide_count} 個收合箭頭，準備展開最下面的一個")

        # === 第三階段：展開最下面的箭頭（好友）===
        self.append_log("  階段 3：展開好友區塊（最下面的箭頭）")

        # 按 Y 座標排序，找到最下面的
        hide_locations_sorted = sorted(
            hide_locations,
            key=lambda loc: self._box_to_tuple(loc)[1] if self._box_to_tuple(loc) else float('inf')
        )

        # 展開最後一個（最下面的）
        last_hide = hide_locations_sorted[-1]
        location_tuple = self._box_to_tuple(last_hide)

        if location_tuple:
            click_x = location_tuple[0] + location_tuple[2] / 2
            click_y = location_tuple[1] + location_tuple[3] / 2

            self.append_log(f"    → 點擊展開最下面的箭頭於 ({int(click_x)}, {int(click_y)})")
            pyautogui_module.click(click_x, click_y)
            time.sleep(1.0)  # 等待動畫完成

        # === 第四階段：驗證最終狀態 ===
        self.append_log("  階段 4：驗證最終狀態")

        show_locations = self._try_locate_all(
            pyautogui_module,
            self.show_arrow_template,
            region=target_region,
            confidence=0.85
        )

        hide_locations = self._try_locate_all(
            pyautogui_module,
            self.hide_arrow_template,
            region=target_region,
            confidence=0.85
        )

        show_count = len(show_locations) if show_locations else 0
        hide_count = len(hide_locations) if hide_locations else 0

        self.append_log(f"    最終狀態：{show_count} 個展開箭頭、{hide_count} 個收合箭頭")

        if show_count >= 1:
            self.append_log(f"  ✓ 箭頭校正成功：至少 1 個展開箭頭（好友區塊）")
            return True
        else:
            self.append_log(f"  ⚠ 未偵測到展開箭頭，好友區塊可能未展開")
            return False

    def _send_message_to_current_chat(
        self,
        pyautogui_module: Any,
        message: str,
        image_path: str | None,
        dry_run: bool,
    ) -> bool:
        """發送訊息到當前打開的聊天窗口"""
        try:
            # 確保 LINE 視窗有焦點
            self._ensure_line_focus(pyautogui_module)

            # 1. 找到訊息輸入框（使用 message_cube.png 模板）
            message_cube_location = self._try_locate(pyautogui_module, self.message_cube_template, confidence=0.85)

            if not message_cube_location:
                self.append_log("  ⚠ 未找到訊息輸入框")
                return False

            cube_coords = self._box_to_tuple(message_cube_location)
            if not cube_coords:
                return False

            # 計算輸入框的點擊位置（在 message_cube 右側）
            input_x = cube_coords[0] + cube_coords[2] + 50
            input_y = cube_coords[1] + cube_coords[3] // 2

            # 2. 點擊輸入框獲得焦點
            pyautogui_module.click(input_x, input_y)
            time.sleep(0.2)

            # 確保焦點在點擊後仍然在 LINE
            self._ensure_line_focus(pyautogui_module)
            time.sleep(0.1)

            # 3. 貼上訊息文字（使用剪貼簿）
            if message:
                try:
                    import pyperclip

                    self.append_log("  → 準備複製訊息到剪貼簿")
                    pyperclip.copy(message)

                    # 剪貼簿操作後立即確保焦點（防止剪貼簿工具搶焦點）
                    self._ensure_line_focus(pyautogui_module)
                    time.sleep(0.1)

                    self.append_log("  → 貼上訊息")
                    pyautogui_module.hotkey('ctrl', 'v')
                    time.sleep(0.3)

                    # 貼上後再次確保焦點
                    self._ensure_line_focus(pyautogui_module)

                except ImportError:
                    self.append_log("  ⚠ pyperclip 未安裝，無法貼上訊息")
                    return False

            # 4. 如果有圖片，附加圖片
            if image_path and Path(image_path).exists():
                self.append_log(f"  → 附加圖片：{image_path}")

                # 將圖片路徑轉換為絕對路徑
                abs_image_path = str(Path(image_path).absolute())

                try:
                    import pyperclip
                    from PIL import Image

                    # 方法1：嘗試使用剪貼簿直接複製圖片
                    try:
                        # 使用 PIL 打開圖片
                        img = Image.open(abs_image_path)

                        # Windows: 使用 win32clipboard 複製圖片到剪貼簿
                        import io
                        import win32clipboard
                        from PIL import ImageGrab

                        output = io.BytesIO()
                        img.convert("RGB").save(output, "BMP")
                        data = output.getvalue()[14:]  # BMP 文件頭是 14 字節
                        output.close()

                        win32clipboard.OpenClipboard()
                        win32clipboard.EmptyClipboard()
                        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                        win32clipboard.CloseClipboard()

                        # 剪貼簿操作後確保焦點
                        self._ensure_line_focus(pyautogui_module)
                        time.sleep(0.2)

                        # 貼上圖片
                        self.append_log(f"  → 貼上圖片")
                        pyautogui_module.hotkey('ctrl', 'v')
                        time.sleep(1.0)

                        # 貼上後確保焦點
                        self._ensure_line_focus(pyautogui_module)

                        self.append_log(f"  ✓ 圖片已附加（方法1：剪貼簿）")

                    except Exception as e1:
                        # 方法2：使用拖放功能
                        self.append_log(f"  → 方法1失敗（{e1}），嘗試方法2：文件拖放")

                        # 保存當前剪貼簿內容
                        saved_clipboard = None
                        try:
                            saved_clipboard = pyperclip.paste()
                        except:
                            pass

                        # 複製文件路徑並模擬拖放
                        pyperclip.copy(abs_image_path)

                        # 剪貼簿操作後確保焦點
                        self._ensure_line_focus(pyautogui_module)
                        time.sleep(0.1)

                        # 使用 Ctrl+V 直接貼上（某些版本的 LINE 支持貼上文件路徑）
                        pyautogui_module.hotkey('ctrl', 'v')
                        time.sleep(1.5)

                        # 貼上後確保焦點
                        self._ensure_line_focus(pyautogui_module)

                        # 恢復剪貼簿
                        if saved_clipboard:
                            try:
                                pyperclip.copy(saved_clipboard)
                            except:
                                pass

                        self.append_log(f"  ✓ 圖片已附加（方法2：路徑貼上）")

                except Exception as e:
                    self.append_log(f"  ✗ 圖片附加失敗：{e}")
                    self.append_log(f"  → 建議：請手動測試在 LINE 聊天窗口中如何附加圖片")
                    # 即使附加失敗，仍然繼續發送文字訊息

            # 5. 發送訊息
            # 發送前最後一次確保焦點
            self._ensure_line_focus(pyautogui_module)
            time.sleep(0.1)

            if dry_run:
                self.append_log("  → 乾跑模式：不實際發送")
                # 選取全部內容（Ctrl+A）然後刪除（Delete）
                pyautogui_module.hotkey('ctrl', 'a')
                time.sleep(0.2)
                pyautogui_module.press('delete')
                time.sleep(0.5)
                return True
            else:
                self.append_log("  → 發送訊息")
                # 按 Enter 發送
                pyautogui_module.press('enter')
                time.sleep(0.5)
                return True

        except Exception as exc:
            self.append_log(f"  ✗ 發送訊息時發生錯誤：{exc}")
            return False

    def _detect_greenchat(self, pyautogui_module: Any) -> bool:
        """檢測是否存在綠色聊天框"""
        if not self.greenchat_template.exists():
            return False

        location = self._try_locate(pyautogui_module, self.greenchat_template, confidence=0.85)
        return location is not None

    def handle_align_arrow_sections(self) -> None:
        """測試頁面的箭頭校正功能"""
        try:
            import pyautogui
        except ImportError as exc:
            messagebox.showerror("箭頭校正", f"無法載入 pyautogui：{exc}")
            return

        if not self._focus_line_window(pyautogui):
            messagebox.showwarning("箭頭校正", "未偵測到 LINE 視窗。")
            return

        self.append_log("開始箭頭校正")
        success = self._calibrate_arrows_for_friend_only(pyautogui)

        if success:
            messagebox.showinfo("箭頭校正結果", "✓ 箭頭校正成功：已展開好友區塊")
        else:
            messagebox.showwarning("箭頭校正結果", "✗ 箭頭校正失敗，請查看日誌")

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
        yield self.greenchat_template
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
            # 如果視窗最小化，先還原
            if getattr(window, "isMinimized", False):
                window.restore()
                time.sleep(0.3)

            # 使用 activate() 將視窗置頂
            window.activate()
            time.sleep(0.5)

            # 額外點擊視窗中心確保獲得焦點
            try:
                # 獲取視窗位置和大小
                left = getattr(window, "left", 0)
                top = getattr(window, "top", 0)
                width = getattr(window, "width", 800)
                height = getattr(window, "height", 600)

                # 點擊視窗中心
                center_x = left + width // 2
                center_y = top + height // 2
                pyautogui_module.click(center_x, center_y)
                time.sleep(0.3)
                self.append_log("已將 LINE 視窗置頂並獲得焦點。")
            except Exception as click_exc:
                self.append_log(f"點擊視窗中心失敗：{click_exc}")

        except Exception as exc:
            self.append_log(f"聚焦 LINE 失敗：{exc}")
            return False

        return True

    def _ensure_line_focus(self, pyautogui_module: Any) -> bool:
        """輕量級的 LINE 視窗焦點檢查和恢復"""
        try:
            windows = pyautogui_module.getWindowsWithTitle("LINE")
            if not windows:
                return False

            window = windows[0]

            # 檢查視窗是否最小化
            if getattr(window, "isMinimized", False):
                window.restore()
                time.sleep(0.3)

            # 主動激活視窗（不檢查是否已活動，直接激活）
            window.activate()
            time.sleep(0.2)

            return True
        except Exception:
            return False

    def _get_line_window_region(self, pyautogui_module: Any) -> tuple[int, int, int, int] | None:
        """獲取 LINE 視窗的區域範圍 (left, top, width, height)"""
        try:
            windows = pyautogui_module.getWindowsWithTitle("LINE")
        except Exception as exc:
            self.append_log(f"取得 LINE 視窗失敗：{exc}")
            return None

        if not windows:
            return None

        window = windows[0]
        try:
            left = getattr(window, "left", 0)
            top = getattr(window, "top", 0)
            width = getattr(window, "width", 800)
            height = getattr(window, "height", 600)

            # 確保座標有效
            if width <= 0 or height <= 0:
                return None

            return (left, top, width, height)
        except Exception as exc:
            self.append_log(f"獲取 LINE 視窗範圍失敗：{exc}")
            return None

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
