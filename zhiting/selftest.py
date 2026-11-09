"""Smoke test executed against the packaged Windows executable in CI."""
import json
from pathlib import Path
import tempfile
from .archive import Archive


def run(output):
    import tkinter as tk
    import pystray
    import win32api
    import pythoncom
    from .app import make_icon
    from .detect import Detector
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        source = root / "验证.pptx"
        source.write_bytes(b"smoke-test")
        archive = Archive(root / "archive")
        assert archive.archive(source).status == "saved"
        assert archive.archive(source).status == "unchanged"
        archive.close()
    window = tk.Tk()
    window.withdraw()
    window.update()
    window.destroy()
    icon = pystray.Icon("smoke-test", make_icon())
    assert icon.icon.size == (64, 64)
    detector = Detector()
    try:
        result = detector.scan(timeout=20)
        assert result["done"], result
    finally:
        detector.close()
    Path(output).write_text(json.dumps({"ok": True, "scan": result}), encoding="utf-8")
