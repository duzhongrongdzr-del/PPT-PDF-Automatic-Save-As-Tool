"""CI-only end-to-end check: a real Edge window opens a valid local PDF."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from zhiting.archive import Archive
from zhiting.detect import Detector


def write_pdf(path):
    stream = b"BT /F1 18 Tf 40 140 Td (Zhiting PDF integration test) Tj ET"
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    data = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for index, value in enumerate(objects, 1):
        offsets.append(len(data))
        data.extend(f"{index} 0 obj\n".encode() + value + b"\nendobj\n")
    xref = len(data)
    data.extend(f"xref\n0 {len(offsets)}\n0000000000 65535 f \n".encode())
    for offset in offsets[1:]:
        data.extend(f"{offset:010d} 00000 n \n".encode())
    data.extend(f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode())
    path.write_bytes(data)


def main():
    import psutil
    candidates = [Path(os.environ.get(name, "C:/Program Files")) / "Microsoft/Edge/Application/msedge.exe"
                  for name in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA")]
    edge = next((p for p in candidates if p.is_file()), None)
    if edge is None:
        raise RuntimeError("CI requires the installed Microsoft Edge browser")
    with tempfile.TemporaryDirectory(prefix="zhiting-pdf-ci-", ignore_cleanup_errors=True) as temp:
        root = Path(temp)
        source = root / "PDF展示测试.pdf"
        write_pdf(source)
        profile = str(root / "edge-test-profile")
        browser = subprocess.Popen([str(edge), "--no-first-run", "--no-default-browser-check", "--user-data-dir=" + profile,
                                    "--new-window", source.as_uri()], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        detector = Detector()
        archive = Archive(root / "copies")
        last = {}
        try:
            deadline = time.monotonic() + 45
            while time.monotonic() < deadline:
                last = detector.scan(timeout=8)
                if str(source.resolve()) in last["paths"]:
                    result = archive.archive(source)
                    assert result.status == "saved"
                    assert Path(result.destination).name == source.name
                    assert Path(result.destination).read_bytes() == source.read_bytes()
                    assert archive.archive(source).status == "unchanged"
                    print(json.dumps({"ok": True, "reader": "Microsoft Edge", "detected_by": last["paths"][str(source.resolve())],
                                      "original_name_preserved": True, "bytes_equal": True}, ensure_ascii=True))
                    return
                time.sleep(0.5)
            raise AssertionError("Local PDF was not detected: " + json.dumps(last, ensure_ascii=True))
        finally:
            detector.close()
            archive.close()
            # Only our isolated test profile, never the default browser profile.
            for process in psutil.process_iter(["name", "cmdline"]):
                if (process.info["name"] or "").lower() == "msedge.exe" and any(profile in arg for arg in (process.info["cmdline"] or [])):
                    try:
                        children = process.children(recursive=True)
                        for child in children:
                            child.terminate()
                        process.terminate()
                        psutil.wait_procs([process, *children], timeout=3)
                    except psutil.Error:
                        pass
            if browser.poll() is None:
                browser.terminate()


if __name__ == "__main__":
    main()
