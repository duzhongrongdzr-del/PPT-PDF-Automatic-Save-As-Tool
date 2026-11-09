"""Best-effort discovery of already open local presentations on Windows.

No Dispatch(progID), launch, document Save, or recursive disk search is used.
An isolated process bounds unresponsive COM and file handle inspection.
"""
import multiprocessing
import os
from pathlib import Path
import time
from .archive import supported

PROCESS_NAMES = {"wpp.exe", "wps.exe", "wpspresentation.exe", "powerpnt.exe"}
PROG_IDS = ("kwpp.Application", "KWPP.Application.9", "PowerPoint.Application")


def _existing_path(value):
    if not isinstance(value, str) or not value:
        return None
    p = Path(value.strip('"'))
    if p.is_absolute() and supported(p) and p.is_file():
        return str(p.resolve())
    return None


def _worker(connection):
    import psutil
    import pythoncom
    import win32com.client.dynamic
    pythoncom.CoInitialize()
    try:
        while connection.recv():
            paths = {}
            notes = []
            processes = []
            for p in psutil.process_iter(["pid", "name"]):
                if (p.info["name"] or "").lower() in PROCESS_NAMES:
                    processes.append(p)
            names = [{"pid": p.pid, "name": p.info["name"]} for p in processes]
            connection.send({"paths": paths, "notes": notes, "processes": names, "done": not processes})
            if not processes:
                continue
            # WPS editions without an exposed automation API may still hold the file open.
            for p in processes:
                try:
                    for item in p.open_files():
                        path = _existing_path(item.path)
                        if path:
                            paths[path] = "文件句柄"
                except (psutil.Error, OSError) as exc:
                    notes.append(f"进程 {p.pid} 句柄读取：{type(exc).__name__}")
                connection.send({"paths": paths, "notes": notes, "processes": names, "done": False})
            for prog_id in PROG_IDS:
                app = collection = document = None
                try:
                    # Wrapping an existing IDispatch cannot start a new Office instance.
                    active = pythoncom.GetActiveObject(prog_id)
                    app = win32com.client.dynamic.Dispatch(active.QueryInterface(pythoncom.IID_IDispatch))
                    collection = app.Presentations
                    for index in range(1, int(collection.Count) + 1):
                        document = collection.Item(index)
                        path = _existing_path(str(document.FullName))
                        if path:
                            paths[path] = prog_id
                        document = None
                except Exception as exc:
                    notes.append(f"{prog_id} 未提供可读文稿：{type(exc).__name__}")
                finally:
                    document = collection = app = active = None
                connection.send({"paths": paths, "notes": notes, "processes": names, "done": False})
            # Read ROT names only, never bind a moniker (which could open a document).
            try:
                table = pythoncom.GetRunningObjectTable()
                context = pythoncom.CreateBindCtx(0)
                for moniker in table.EnumRunning():
                    try:
                        path = _existing_path(moniker.GetDisplayName(context, None))
                        if path:
                            paths.setdefault(path, "运行对象表")
                    except Exception:
                        continue
                table = context = moniker = None
            except Exception as exc:
                notes.append(f"运行对象表：{type(exc).__name__}")
            connection.send({"paths": paths, "notes": notes, "processes": names, "done": True})
    except (EOFError, BrokenPipeError, OSError):
        pass
    finally:
        pythoncom.CoUninitialize()
        connection.close()


class Detector:
    def __init__(self):
        self.process = None
        self.connection = None
        self.cooldown = 0

    def close(self):
        if self.connection:
            self.connection.close()
            self.connection = None
        if self.process:
            if self.process.is_alive():
                self.process.terminate()
            self.process.join(timeout=1)
            self.process.close()
            self.process = None

    def scan(self, cancelled=lambda: False, timeout=8):
        result = {"paths": {}, "notes": [], "processes": [], "done": False}
        if time.monotonic() < self.cooldown:
            result["notes"] = ["识别接口暂时无响应，稍后自动重试"]
            return result
        if not self.process or not self.process.is_alive():
            self.close()
            context = multiprocessing.get_context("spawn")
            self.connection, child = context.Pipe()
            self.process = context.Process(target=_worker, args=(child,), daemon=True)
            self.process.start()
            child.close()
        try:
            self.connection.send(True)
            deadline = time.monotonic() + timeout
            while time.monotonic() < deadline and not cancelled():
                if self.connection.poll(0.1):
                    result = self.connection.recv()
                    if result["done"]:
                        return result
                if not self.process.is_alive():
                    raise EOFError("识别进程已退出")
            result["notes"].append("识别超时或已暂停；保留已识别的路径，30 秒后重试")
        except (EOFError, BrokenPipeError, OSError) as exc:
            result["notes"].append(f"识别进程异常：{exc}")
        self.close()
        self.cooldown = time.monotonic() + 30
        return result
