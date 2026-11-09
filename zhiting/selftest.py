"""Smoke test executed against the packaged Windows executable in CI."""
import json
from pathlib import Path
import tempfile
from .archive import Archive


def run(output):
    import os
    from unittest.mock import patch
    import tkinter as tk
    import pystray
    import win32api
    import pythoncom
    from .app import make_icon
    from .detect import Detector
    from .detect import _original_candidates
    from win32com.shell import shell
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        source = root / "验证.pptx"
        source.write_bytes(b"smoke-test")
        archive = Archive(root / "archive")
        saved = archive.archive(source)
        assert saved.status == "saved"
        assert Path(saved.destination).name == source.name
        assert archive.archive(source).status == "unchanged"
        archive.close()
        # Verify Windows shortcut reading in the frozen app, using isolated fixtures.
        roaming = root / "Roaming"
        recent = roaming / "Microsoft" / "Windows" / "Recent"
        recent.mkdir(parents=True)
        pythoncom.CoInitialize()
        try:
            link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
            link.SetPath(str(source))
            link.QueryInterface(pythoncom.IID_IPersistFile).Save(str(recent / "验证.lnk"), 0)
            with patch.dict(os.environ, {"APPDATA": str(roaming)}):
                assert _original_candidates([], ["验证.pptx - WPS Office"]) == [str(source.resolve())]
            link = None
        finally:
            pythoncom.CoUninitialize()
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
