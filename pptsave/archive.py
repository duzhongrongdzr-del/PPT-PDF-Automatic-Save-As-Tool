"""Streaming, content-addressed local archive. No Office or GUI dependency."""
from dataclasses import dataclass
from contextlib import closing
from datetime import datetime
from pathlib import Path
import hashlib
import os
import re
import sqlite3
import tempfile
import time

EXTENSIONS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pdf"}
CHUNK = 1024 * 1024


def supported(path):
    p = Path(path)
    return p.suffix.lower() in EXTENSIONS and not p.name.startswith("~$")


def temporary_name(path):
    return Path(path).name.lower().startswith(("~tmp", "~wpp", "~ppt", "~pdf"))


def move_without_overwrite(source, destination):
    if os.name == "nt":
        os.rename(source, destination)
    else:
        os.link(source, destination)
        Path(source).unlink()


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
        self.meta = self.root / ".pptsave"
        legacy = self.root / ".zhiting"
        if legacy.is_dir() and not self.meta.exists():
            legacy.rename(self.meta)
        self.meta.mkdir(exist_ok=True)
        self.db = sqlite3.connect(self.meta / "index.sqlite3", timeout=10)
        self.db.execute("CREATE TABLE IF NOT EXISTS copies (digest TEXT PRIMARY KEY, relative_path TEXT NOT NULL, source TEXT, created TEXT)")
        self.db.commit()
        # Preserve both histories if a partial upgrade left two metadata folders.
        legacy_index = legacy / "index.sqlite3"
        if legacy_index.is_file():
            try:
                with closing(sqlite3.connect(legacy_index.as_uri() + "?mode=ro", uri=True)) as old:
                    rows = old.execute("SELECT digest, relative_path, source, created FROM copies").fetchall()
                with self.db:
                    self.db.executemany("INSERT OR IGNORE INTO copies VALUES (?, ?, ?, ?)", rows)
            except Exception:
                self.db.close()
                raise
        self.cache = {}

    def close(self):
        self.db.close()

    def destination_for(self, source, digest):
        """Keep the full original basename; collisions change the folder only."""
        month = self.root / datetime.now().strftime("%Y-%m")
        if temporary_name(source):
            month /= "待确认原名"
        candidate = month / source.name
        number = 0
        while candidate.exists():
            if candidate.is_file() and digest_file(candidate) == digest:
                return candidate
            number += 1
            version = digest[:16] if number == 1 else f"{digest[:16]}_{number}"
            candidate = month / "其他版本" / version / source.name
        candidate.parent.mkdir(parents=True, exist_ok=True)
        return candidate

    def should_restore_name(self, candidate, recorded_source, current_source, digest):
        """Only migrate this app's old generated names or a verified temp copy."""
        if temporary_name(current_source):
            return False
        old = Path(recorded_source or "")
        stem = old.stem[:60].rstrip(" .") or "presentation"
        legacy = re.fullmatch(re.escape(f"{stem}__{digest[:16]}") + r"(?:_\d+)?" + re.escape(old.suffix.lower()), candidate.name)
        return bool(legacy) or (temporary_name(old) and temporary_name(candidate))

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
            row = self.db.execute("SELECT relative_path, source FROM copies WHERE digest=?", (digest,)).fetchone()
            destination = None
            status = "duplicate"
            if row:
                candidate = (self.root / row[0]).resolve()
                if candidate.is_relative_to(self.root) and candidate.is_file() and digest_file(candidate) == digest:
                    destination = candidate
                    if self.should_restore_name(candidate, row[1], source, digest):
                        destination = self.destination_for(source, digest)
                        if destination != candidate:
                            if not destination.exists():
                                move_without_overwrite(candidate, destination)
                            # If an identical target already exists, retain the old file.
                            # It may have been deliberately copied by the user.
                            status = "renamed"
                        with self.db:
                            self.db.execute("UPDATE copies SET relative_path=?, source=? WHERE digest=?",
                                            (str(destination.relative_to(self.root)), key, digest))
            if destination is None:
                destination = self.destination_for(source, digest)
                if not destination.exists():
                    move_without_overwrite(temporary, destination)
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
