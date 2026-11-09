import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile


def state_directory():
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ZhitingPPTSave"


def load_config(path):
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("destination", ""), str):
            raise ValueError("设置格式无效")
        return {"destination": data.get("destination", "")}
    except FileNotFoundError:
        return {"destination": ""}


def save_config(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(dir=path.parent, prefix="config-", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as output:
            json.dump(data, output, ensure_ascii=False, indent=2)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "ZhitingPPTSave"


def startup_enabled():
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            winreg.QueryValueEx(key, RUN_NAME)
        return True
    except FileNotFoundError:
        return False


def set_startup(enabled):
    import winreg
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
        if enabled:
            if getattr(sys, "frozen", False):
                command = [sys.executable, "--background"]
            else:
                command = [str(Path(sys.executable).with_name("pythonw.exe")), str(Path(sys.argv[0]).resolve()), "--background"]
            winreg.SetValueEx(key, RUN_NAME, 0, winreg.REG_SZ, subprocess.list2cmdline(command))
        else:
            try:
                winreg.DeleteValue(key, RUN_NAME)
            except FileNotFoundError:
                pass
