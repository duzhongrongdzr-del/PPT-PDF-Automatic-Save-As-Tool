"""Best-effort discovery of already open local presentations on Windows.

No Dispatch(progID), launch, document Save, or recursive disk search is used.
An isolated process bounds unresponsive COM and file handle inspection.
"""
import multiprocessing
import os
from pathlib import Path
import time
from .archive import supported, temporary_name
from .originals import prefer_originals, title_matches_name

PROCESS_NAMES = {"wpp.exe", "wps.exe", "wpspresentation.exe", "powerpnt.exe"}
PROG_IDS = ("kwpp.Application", "KWPP.Application.9", "PowerPoint.Application")


def _existing_path(value):
    if not isinstance(value, str) or not value:
        return None
    p = Path(value.strip('"'))
    if p.is_absolute() and supported(p) and p.is_file():
        return str(p.resolve())
    return None


def _window_titles(processes):
    import win32gui
    import win32process
    pids = {p.pid for p in processes}
    titles = []

    def collect(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32process.GetWindowThreadProcessId(hwnd)[1] in pids:
            title = win32gui.GetWindowText(hwnd)
            if title:
                titles.append(title)
        return True

    win32gui.EnumWindows(collect, None)
    return titles


def _original_candidates(processes, titles):
    """Read process arguments and title-matched Recent shortcuts, never crawl disks."""
    import pythoncom
    from win32com.shell import shell
    candidates = []
    for process in processes:
        try:
            for arg in process.cmdline()[1:]:
                path = _existing_path(arg)
                if path and any(title_matches_name(title, path) for title in titles):
                    candidates.append(path)
        except Exception:
            continue
    appdata = os.environ.get("APPDATA")
    if not appdata or not titles:
        return candidates
    recent = Path(appdata) / "Microsoft" / "Windows" / "Recent"
    try:
        # Only shortcut names which identify a presentation in a current window.
        links = [p for p in recent.glob("*.lnk")
                 if any(title_matches_name(t, p.stem) for t in titles)]
        links.sort(key=lambda p: p.stat().st_mtime_ns, reverse=True)
        for shortcut in links[:32]:
            try:
                link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
                link.QueryInterface(pythoncom.IID_IPersistFile).Load(str(shortcut))
                path = _existing_path(link.GetPath(shell.SLGP_RAWPATH)[0])
                if path and any(title_matches_name(t, path) for t in titles):
                    candidates.append(path)
                link = None
            except Exception:
                continue
    except OSError:
        pass
    return candidates


def _worker(connection):
    import psutil
    import pythoncom
    import win32com.client.dynamic
    pythoncom.CoInitialize()
    digest_cache = {}
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
                        try:
                            original = _existing_path(str(Path(str(document.Path)) / str(document.Name)))
                            if original:
                                paths[original] = prog_id
                        except Exception:
                            pass
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
            titles = []
            unresolved = []
            recovered = []
            if any(temporary_name(p) for p in paths):
                try:
                    titles = _window_titles(processes)
                    candidates = _original_candidates(processes, titles)
                    paths, unresolved, recovered = prefer_originals(paths, candidates, digest_cache)
                except Exception as exc:
                    notes.append(f"原文件名恢复：{type(exc).__name__}")
                    unresolved = [p for p in paths if temporary_name(p)]
            connection.send({"paths": paths, "notes": notes, "processes": names, "done": True,
                             "window_titles": titles, "unresolved_names": unresolved, "recovered_names": recovered})
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
