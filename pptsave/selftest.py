"""Smoke test executed against the packaged Windows executable in CI."""
import json
import logging
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
    from .app import Application, AUTHOR_URL, make_icon
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
        pdf = root / "PDF展示.pdf"
        pdf.write_bytes(b"%PDF-1.4\nsmoke-test")
        pdf_saved = archive.archive(pdf)
        assert pdf_saved.status == "saved"
        assert Path(pdf_saved.destination).name == pdf.name
        assert Path(pdf_saved.destination).read_bytes() == pdf.read_bytes()
        assert archive.archive(pdf).status == "unchanged"
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
    # Exercise the real settings window without starting monitoring or changing
    # the account's startup entry. Verify the new labels fit at minimum size.
    with tempfile.TemporaryDirectory() as temp:
        with patch("pptsave.app.state_directory", return_value=Path(temp)), \
                patch("pptsave.app.migrate_startup", return_value=False), \
                patch("pptsave.app.RotatingFileHandler", return_value=logging.NullHandler()), \
                patch("pptsave.app.threading.Thread"), patch("pystray.Icon"):
            application = Application()
        window = application.root
        try:
            assert window.title().startswith("PPTSave ")
            application.status.set("监测中 · 本次识别 2 个文件 · 本次运行新增 2 份")
            application.details.set("最近保存：D:/报告留存/2026-11/会议报告.pdf")
            def descendants(widget):
                for child in widget.winfo_children():
                    yield child
                    yield from descendants(child)
            for size in ("660x520+30+30", "680x540+30+30"):
                window.geometry(size)
                window.update()
                widgets = list(descendants(window))
                labels = {str(w.cget("text")): w for w in widgets if w.winfo_class() == "TLabel"}
                title = labels["PPTSave"]
                disclaimer = labels["仅供学习交流，请在获得授权后使用。"]
                author = labels["作者：duzhongrongdzr-del · GitHub"]
                assert title.winfo_rootx() + title.winfo_width() <= disclaimer.winfo_rootx()
                assert author.winfo_rooty() > application.pause_button.winfo_rooty()
                for widget in widgets:
                    assert widget.winfo_ismapped(), (size, str(widget))
                    assert widget.winfo_width() >= widget.winfo_reqwidth(), (size, str(widget), "clipped width")
                    assert widget.winfo_height() >= widget.winfo_reqheight(), (size, str(widget), "clipped height")
                with patch("pptsave.app.webbrowser.open") as open_profile:
                    author.event_generate("<Button-1>", x=5, y=5)
                    window.update()
                    open_profile.assert_called_once_with(AUTHOR_URL)
            # A source-build preview lets the release author inspect native Windows
            # rendering; the frozen app follows the same layout assertions above.
            if Path(output).name == "source-smoke.json":
                from PIL import ImageGrab
                window.lift()
                window.update()
                x, y = window.winfo_rootx(), window.winfo_rooty()
                ImageGrab.grab(bbox=(x, y, x + window.winfo_width(), y + window.winfo_height())).save("PPTSave-preview.png")
        finally:
            application.stop.set()
            window.destroy()
    icon = pystray.Icon("smoke-test", make_icon())
    assert icon.icon.size == (64, 64)
    detector = Detector()
    try:
        result = detector.scan(timeout=20)
        assert result["done"], result
    finally:
        detector.close()
    Path(output).write_text(json.dumps({"ok": True, "ui_layout": True, "scan": result}), encoding="utf-8")
