import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import platform
import queue
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .archive import Archive
from .detect import Detector
from .settings import load_config, save_config, set_startup, startup_enabled, state_directory


def make_icon():
    from PIL import Image, ImageDraw
    image = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((5, 5, 59, 59), radius=14, fill="#176b5b")
    draw.rectangle((19, 15, 44, 47), fill="white")
    draw.rectangle((24, 22, 39, 26), fill="#176b5b")
    draw.line((25, 35, 30, 40, 41, 30), fill="#176b5b", width=5)
    return image


class Application:
    def __init__(self, background=False):
        import pystray
        self.directory = state_directory()
        self.directory.mkdir(parents=True, exist_ok=True)
        handler = RotatingFileHandler(self.directory / "app.log", maxBytes=512_000, backupCount=2, encoding="utf-8")
        logging.basicConfig(level=logging.INFO, handlers=[handler], format="%(asctime)s %(levelname)s %(message)s")
        self.config_path = self.directory / "config.json"
        self.config_error = ""
        try:
            self.config = load_config(self.config_path)
        except (OSError, ValueError) as exc:
            self.config = {"destination": ""}
            self.config_error = f"原设置无法读取，请重新选择目录：{exc}"
        self.lock = threading.Lock()
        self.events = queue.Queue()
        self.stop = threading.Event()
        self.wake = threading.Event()
        self.paused = False
        self.snapshot = {}
        self.root = tk.Tk()
        self.root.title(f"知庭 · PPT 自动留存  {__version__}")
        self.root.geometry("660x490")
        self.root.minsize(600, 460)
        self.root.protocol("WM_DELETE_WINDOW", self.root.withdraw)
        self.destination = tk.StringVar(value=self.config["destination"] or "尚未选择")
        self.status = tk.StringVar(value="请选择保存目录，随后开始后台监测。")
        self.details = tk.StringVar(value="打开 PPT 后，通常等待 5–15 秒即可查看副本。")
        self.startup = tk.BooleanVar(value=startup_enabled())
        self.build_window()
        self.icon = pystray.Icon("ZhitingPPTSave", make_icon(), "知庭 · PPT 自动留存", menu=pystray.Menu(
            pystray.MenuItem("设置与状态", lambda: self.events.put(("show", None)), default=True),
            pystray.MenuItem("打开保存目录", lambda: self.events.put(("open", None))),
            pystray.MenuItem("暂停 / 继续", lambda: self.events.put(("pause", None))),
            pystray.MenuItem("退出", lambda: self.events.put(("quit", None)))))
        self.icon.run_detached()
        self.thread = threading.Thread(target=self.monitor, daemon=True, name="archive-monitor")
        self.thread.start()
        self.root.after(100, self.pump)
        if background and self.config["destination"]:
            self.root.withdraw()
        if self.config_error:
            self.root.after(200, lambda: messagebox.showerror("设置需要恢复", self.config_error))

    def build_window(self):
        panel = ttk.Frame(self.root, padding=24)
        panel.pack(fill="both", expand=True)
        ttk.Label(panel, text="PPT 自动留存", font=("Microsoft YaHei UI", 20, "bold")).pack(anchor="w")
        ttk.Label(panel, text="WPS / PowerPoint 打开本地或 U 盘演示文件后，自动保留副本。", wraplength=590).pack(anchor="w", pady=(8, 20))
        ttk.Label(panel, text="保存目录", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
        ttk.Entry(panel, textvariable=self.destination, state="readonly").pack(fill="x", pady=6)
        row = ttk.Frame(panel)
        row.pack(fill="x")
        ttk.Button(row, text="选择 / 修改目录", command=self.choose_directory).pack(side="left")
        ttk.Button(row, text="打开目录", command=self.open_directory).pack(side="left", padx=8)
        ttk.Checkbutton(panel, text="登录 Windows 后自动运行", variable=self.startup, command=self.change_startup).pack(anchor="w", pady=16)
        ttk.Separator(panel).pack(fill="x", pady=(0, 12))
        ttk.Label(panel, textvariable=self.status, font=("Microsoft YaHei UI", 11, "bold"), wraplength=590).pack(anchor="w")
        ttk.Label(panel, textvariable=self.details, wraplength=590, justify="left").pack(anchor="w", pady=10)
        controls = ttk.Frame(panel)
        controls.pack(side="bottom", fill="x", pady=(12, 0))
        self.pause_button = ttk.Button(controls, text="暂停", command=self.toggle_pause)
        self.pause_button.pack(side="left")
        ttk.Button(controls, text="导出诊断", command=self.export_diagnostics).pack(side="left", padx=8)
        ttk.Button(controls, text="收起到托盘", command=self.root.withdraw).pack(side="right")
        ttk.Label(panel, text="只读取已经保存的文件；同内容去重，保存修改后留新版本。\n关闭此窗口后仍会运行；完全停止请右键托盘图标 → 退出。", wraplength=590, foreground="#666666").pack(side="bottom", anchor="w")

    def choose_directory(self):
        chosen = filedialog.askdirectory(title="选择 PPT 副本保存目录")
        if not chosen:
            return
        try:
            chosen = str(Path(chosen).resolve())
            with tempfile.TemporaryFile(dir=chosen):
                pass
            new_config = {"destination": chosen}
            save_config(self.config_path, new_config)
            with self.lock:
                self.config = new_config
            self.destination.set(chosen)
            self.wake.set()
        except OSError as exc:
            messagebox.showerror("目录不可用", str(exc))

    def open_directory(self):
        try:
            with self.lock:
                directory = self.config["destination"]
            if directory:
                os.startfile(directory)
            else:
                self.choose_directory()
        except OSError as exc:
            messagebox.showerror("无法打开目录", str(exc))

    def change_startup(self):
        try:
            set_startup(self.startup.get())
        except OSError as exc:
            self.startup.set(startup_enabled())
            messagebox.showerror("开机启动设置失败", str(exc))

    def toggle_pause(self):
        with self.lock:
            self.paused = not self.paused
            paused = self.paused
        self.pause_button.config(text="继续" if paused else "暂停")
        self.status.set("已暂停" if paused else "正在恢复监测…")
        self.icon.title = "知庭 · 已暂停" if paused else "知庭 · PPT 自动留存"
        self.wake.set()

    def cancelled(self, expected_directory):
        with self.lock:
            return self.stop.is_set() or self.paused or self.config["destination"] != expected_directory

    def monitor(self):
        detector = Detector()
        archive = None
        directory = None
        saved = 0
        last_copy = ""
        try:
            while not self.stop.is_set():
                self.wake.clear()
                with self.lock:
                    chosen = self.config["destination"]
                    paused = self.paused
                if chosen and not paused:
                    try:
                        if chosen != directory or archive is None:
                            if archive:
                                archive.close()
                            archive = None
                            archive = Archive(chosen)
                            directory = chosen
                        result = detector.scan(lambda: self.cancelled(chosen))
                        errors = []
                        for path in result["paths"]:
                            if self.cancelled(chosen):
                                break
                            try:
                                copied = archive.archive(path, lambda: self.cancelled(chosen))
                                if copied.status == "saved":
                                    saved += 1
                                    last_copy = copied.destination
                                    logging.info("已保存 %s", copied.destination)
                            except (OSError, ValueError, RuntimeError) as exc:
                                errors.append(f"{Path(path).name}：{exc}")
                        if not self.cancelled(chosen):
                            result.update(saved=saved, last_copy=last_copy, errors=errors)
                            self.events.put(("status", result))
                    except Exception as exc:
                        logging.exception("本次扫描失败，下次重试")
                        self.events.put(("error", str(exc)))
                        if archive:
                            archive.close()
                            archive = None
                self.wake.wait(5)
        finally:
            detector.close()
            if archive:
                archive.close()

    def pump(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "show":
                    self.root.deiconify()
                    self.root.lift()
                elif kind == "open":
                    self.open_directory()
                elif kind == "pause":
                    self.toggle_pause()
                elif kind == "quit":
                    self.quit()
                    return
                elif kind == "error":
                    self.status.set("暂时无法完成复制，稍后重试")
                    self.details.set(payload[:280])
                    self.snapshot["error"] = payload
                elif kind == "status":
                    self.snapshot = payload
                    if not self.paused:
                        count = len(payload["paths"])
                        self.status.set(f"监测中 · 本次识别 {count} 个文件 · 本次运行新增 {payload['saved']} 份")
                        if payload["errors"]:
                            detail = "等待重试：" + "；".join(payload["errors"])
                        elif not count and payload["processes"]:
                            detail = "检测到 WPS / PowerPoint，暂未取得文件路径。请打开已保存的 PPT；若仍无法识别，请导出诊断。"
                        elif payload["last_copy"]:
                            detail = "最近保存：" + payload["last_copy"]
                        else:
                            detail = "等待你在 WPS / PowerPoint 中打开 PPT。"
                        self.details.set(detail[:300])
        except queue.Empty:
            pass
        if not self.stop.is_set():
            self.root.after(200, self.pump)

    def export_diagnostics(self):
        path = filedialog.asksaveasfilename(title="保存诊断（包含本机文件路径）", defaultextension=".json", initialfile="zhiting-diagnostics.json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with self.lock:
                config = dict(self.config)
            data = {"version": __version__, "os": platform.platform(), "config": config, "scan": self.snapshot,
                    "note": "仅保存在本机，包含文件路径；不包含 PPT 内容。"}
            Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            messagebox.showinfo("诊断已保存", path)
        except OSError as exc:
            messagebox.showerror("无法保存诊断", str(exc))

    def quit(self):
        self.stop.set()
        self.wake.set()
        self.icon.stop()
        self.root.withdraw()
        # Keep the UI responsive while the worker finishes/cleans a partial copy.
        self.wait_for_exit()

    def wait_for_exit(self):
        if self.thread.is_alive():
            self.root.after(100, self.wait_for_exit)
        else:
            self.root.destroy()


def run(background=False):
    import win32api
    import win32event
    import winerror
    mutex = win32event.CreateMutex(None, False, r"Local\ZhitingPPTSave")
    try:
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            if not background:
                win32api.MessageBox(0, "程序已在运行。请在右下角托盘（含隐藏图标）中找到绿色图标，双击打开设置。", "知庭 PPT 自动留存", 0)
            return
        application = Application(background)
        application.root.mainloop()
    finally:
        win32api.CloseHandle(mutex)
