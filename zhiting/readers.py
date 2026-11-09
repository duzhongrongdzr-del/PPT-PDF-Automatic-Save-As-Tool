"""Reader process selection and local-file references; no web downloads."""
import re
from urllib.parse import unquote, urlsplit

PPT_PROCESSES = {"wpp.exe", "wps.exe", "wpspresentation.exe", "powerpnt.exe"}
PDF_PROCESSES = {
    "wpspdf.exe", "wpdf.exe", "wpsufd.exe", "kpdf.exe",
    "acrord32.exe", "acrobat.exe", "foxitpdfreader.exe", "foxitreader.exe",
    "foxitpdfeditor.exe", "pdfxedit.exe", "sumatrapdf.exe",
}
BROWSER_PROCESSES = {"msedge.exe", "chrome.exe", "firefox.exe"}
PROCESS_NAMES = PPT_PROCESSES | PDF_PROCESSES | BROWSER_PROCESSES


def local_reference(value):
    if not isinstance(value, str) or not value.strip():
        return None
    value = value.strip().strip('"')
    if value.lower().startswith("file://"):
        parsed = urlsplit(value)
        path = unquote(parsed.path)
        if parsed.netloc and parsed.netloc.lower() != "localhost":
            path = "//" + parsed.netloc + path
        elif re.match(r"^/[A-Za-z]:/", path):
            path = path[1:]
        return path
    if re.match(r"^[A-Za-z][A-Za-z0-9+.-]*:", value) and not re.match(r"^[A-Za-z]:[\\/]", value):
        return None
    if value.startswith("--"):
        return None
    return value


def is_pdf_reference(value):
    path = local_reference(value)
    return bool(path and path.lower().endswith(".pdf"))


def select_processes(processes, windows):
    """Include readers and PDF-related browser process trees only."""
    names = {p.pid: (p.info.get("name") or "").lower() for p in processes}
    selected = {pid for pid, name in names.items() if name in PPT_PROCESSES | PDF_PROCESSES}
    for window in windows:
        pid = window["pid"]
        if names.get(pid) not in BROWSER_PROCESSES:
            continue
        process = next((p for p in processes if p.pid == pid), None)
        if ".pdf" in window["title"].lower():
            selected.add(pid)
        elif process:
            try:
                if any(is_pdf_reference(arg) for arg in process.cmdline()[1:]):
                    selected.add(pid)
            except Exception:
                pass
    changed = True
    while changed:
        changed = False
        for process in processes:
            if names[process.pid] in BROWSER_PROCESSES and process.info.get("ppid") in selected and process.pid not in selected:
                selected.add(process.pid)
                changed = True
    # Desktop readers first, browser roots before their child processes.
    return sorted((p for p in processes if p.pid in selected),
                  key=lambda p: (names[p.pid] in BROWSER_PROCESSES, p.info.get("ppid") in selected))
