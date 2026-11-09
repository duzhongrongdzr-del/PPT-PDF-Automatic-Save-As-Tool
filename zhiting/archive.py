"""Streaming, content-addressed local archive. No Office or GUI dependency."""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import hashlib
import os
import sqlite3
import tempfile
import time

EXTENSIONS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm"}
CHUNK = 1024 * 1024


def supported(path):
    p = Path(path)
    return p.suffix.lower() in EXTENSIONS and not p.name.startswith("~$")


def signature(path):
    s = Path(path).stat()
    return s.st_size, s.st_mtime_ns, s.st_ctime_ns, s.st_ino


def digest_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(CHUNK), b""):
            h.update(chunk)
    return h.hexdigest()


@dataclass
class Result:
    status: str
    destination: str = ""


class Archive:
    def __init__(self, directory):
        self.root = Path(directory).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.meta = self.root / ".zhiting"
        self.meta.mkdir(exist_ok=True)
        self.db = sqlite3.connect(self.meta / "index.sqlite3", timeout=10)
        self.db.execute("CREATE TABLE IF NOT EXISTS copies (digest TEXT PRIMARY KEY, relative_path TEXT NOT NULL, source TEXT, created TEXT)")
        self.db.commit()
        self.cache = {}

    def close(self):
        self.db.close()

    def archive(self, source, cancel=lambda: False):
        source = Path(source).resolve()
        if not supported(source) or source.is_relative_to(self.root):
            return Result("ignored")
        before = signature(source)
        if before[0] == 0:
            raise OSError("文件仍为空，等待下一次扫描")
        key = str(source)
        cached = self.cache.get(key)
        if cached:
            old_sig, dest, dest_sig, checked_at = cached
            # Re-check content periodically even when timestamp/size are preserved.
            if before == old_sig and time.monotonic() - checked_at < 60:
                try:
                    if signature(dest) == dest_sig:
                        return Result("unchanged", str(dest))
                except OSError:
                    pass
        fd, temporary = tempfile.mkstemp(prefix="copy-", suffix=".part", dir=self.meta)
        temporary = Path(temporary)
        try:
            h = hashlib.sha256()
            with os.fdopen(fd, "wb") as output, source.open("rb") as original:
                for chunk in iter(lambda: original.read(CHUNK), b""):
                    if cancel():
                        raise InterruptedError("已暂停或退出")
                    output.write(chunk)
                    h.update(chunk)
                output.flush()
                os.fsync(output.fileno())
            if signature(source) != before:
                raise OSError("源文件在复制期间改变，等待重新复制")
            if cancel():
                raise InterruptedError("已暂停或退出")
            digest = h.hexdigest()
            row = self.db.execute("SELECT relative_path FROM copies WHERE digest=?", (digest,)).fetchone()
            destination = None
            if row:
                candidate = (self.root / row[0]).resolve()
                if candidate.is_relative_to(self.root) and candidate.is_file() and digest_file(candidate) == digest:
                    destination = candidate
            status = "duplicate"
            if destination is None:
                month = self.root / datetime.now().strftime("%Y-%m")
                month.mkdir(exist_ok=True)
                stem = source.stem[:60].rstrip(" .") or "presentation"
                filename = f"{stem}__{digest[:16]}"
                destination = month / f"{filename}{source.suffix.lower()}"
                number = 1
                while destination.exists():
                    if destination.is_file() and digest_file(destination) == digest:
                        break
                    destination = month / f"{filename}_{number}{source.suffix.lower()}"
                    number += 1
                else:
                    # os.rename never overwrites on Windows. Handle an external race safely.
                    if os.name == "nt":
                        os.rename(temporary, destination)
                    else:
                        os.link(temporary, destination)
                        temporary.unlink()
                    status = "saved"
                with self.db:
                    self.db.execute("INSERT OR REPLACE INTO copies VALUES (?, ?, ?, ?)",
                                    (digest, str(destination.relative_to(self.root)), key, datetime.now().isoformat()))
            self.cache[key] = (before, destination, signature(destination), time.monotonic())
            if len(self.cache) > 512:
                self.cache.pop(next(iter(self.cache)))
            return Result(status, str(destination))
        finally:
            temporary.unlink(missing_ok=True)
