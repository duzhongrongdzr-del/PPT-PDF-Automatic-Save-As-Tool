from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from zhiting.archive import Archive, signature, supported
from zhiting.settings import load_config, save_config


class ArchiveTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.destination = self.root / "副本"
        self.archive = Archive(self.destination)

    def tearDown(self):
        self.archive.close()
        self.temp.cleanup()

    def source(self, name="报告.pptx", content=b"original"):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)
        return path

    def test_copy_preserves_bytes_and_source(self):
        source = self.source(content=b"a" * (3 * 1024 * 1024 + 17))
        before = signature(source)
        result = self.archive.archive(source)
        self.assertEqual(result.status, "saved")
        self.assertEqual(Path(result.destination).read_bytes(), source.read_bytes())
        self.assertEqual(signature(source), before)

    def test_same_content_different_name_and_directory_deduplicates(self):
        first = self.archive.archive(self.source())
        second = self.archive.archive(self.source("U盘/改名.PPTX"))
        self.assertEqual(second.status, "duplicate")
        self.assertEqual(first.destination, second.destination)

    def test_content_change_creates_version(self):
        first = self.archive.archive(self.source())
        second = self.archive.archive(self.source(content=b"new version"))
        self.assertEqual(second.status, "saved")
        self.assertNotEqual(first.destination, second.destination)
        self.assertEqual(Path(first.destination).read_bytes(), b"original")

    def test_same_name_different_content_never_overwrites(self):
        first = self.archive.archive(self.source("one/report.ppt", b"one"))
        second = self.archive.archive(self.source("two/report.ppt", b"two"))
        self.assertNotEqual(first.destination, second.destination)
        self.assertEqual(Path(first.destination).read_bytes(), b"one")

    def test_dedup_survives_restart(self):
        source = self.source()
        first = self.archive.archive(source)
        self.archive.close()
        self.archive = Archive(self.destination)
        second = self.archive.archive(source)
        self.assertEqual(second.status, "duplicate")
        self.assertEqual(first.destination, second.destination)

    def test_deleted_copy_is_recreated(self):
        source = self.source()
        first = self.archive.archive(source)
        Path(first.destination).unlink()
        self.assertEqual(self.archive.archive(source).status, "saved")
        self.assertEqual(Path(first.destination).read_bytes(), b"original")

    def test_modified_copy_is_preserved_and_repaired_as_new_file(self):
        source = self.source()
        first = self.archive.archive(source)
        Path(first.destination).write_bytes(b"edited archive")
        second = self.archive.archive(source)
        self.assertEqual(second.status, "saved")
        self.assertEqual(Path(first.destination).read_bytes(), b"edited archive")
        self.assertEqual(Path(second.destination).read_bytes(), b"original")

    def test_unchanged_copy_skips_disk_write(self):
        source = self.source()
        self.archive.archive(source)
        with patch("zhiting.archive.tempfile.mkstemp", side_effect=AssertionError("unexpected write")):
            self.assertEqual(self.archive.archive(source).status, "unchanged")

    def test_source_changing_during_copy_is_retried(self):
        source = self.source()
        sig = signature(source)
        with patch("zhiting.archive.signature", side_effect=[sig, (999, 0, 0, 0)]):
            with self.assertRaises(OSError):
                self.archive.archive(source)
        self.assertFalse(list(self.destination.rglob("*.pptx")))
        self.assertFalse(list(self.destination.rglob("*.part")))
        self.assertEqual(self.archive.archive(source).status, "saved")

    def test_pause_cleans_partial_copy(self):
        source = self.source()
        with self.assertRaises(InterruptedError):
            self.archive.archive(source, cancel=lambda: True)
        self.assertFalse(list(self.destination.rglob("*.part")))
        self.assertFalse(list(self.destination.rglob("*.pptx")))

    def test_output_folder_and_lock_files_are_ignored(self):
        first = self.archive.archive(self.source())
        self.assertEqual(self.archive.archive(first.destination).status, "ignored")
        self.assertEqual(self.archive.archive(self.source("~$report.pptx")).status, "ignored")
        self.assertEqual(self.archive.archive(self.source("report.txt")).status, "ignored")

    def test_missing_and_empty_input_do_not_leave_files(self):
        with self.assertRaises(FileNotFoundError):
            self.archive.archive(self.root / "missing.pptx")
        with self.assertRaises(OSError):
            self.archive.archive(self.source(content=b""))
        self.assertFalse(list(self.destination.rglob("*.part")))

    def test_new_destination_gets_its_own_copy(self):
        source = self.source()
        self.archive.archive(source)
        other = Archive(self.root / "new-output")
        try:
            self.assertEqual(other.archive(source).status, "saved")
        finally:
            other.close()

    def test_config_roundtrip_and_invalid_config(self):
        path = self.root / "config.json"
        self.assertEqual(load_config(path), {"destination": ""})
        save_config(path, {"destination": "D:/会议/副本"})
        self.assertEqual(load_config(path)["destination"], "D:/会议/副本")
        path.write_text("[]", encoding="utf-8")
        with self.assertRaises(ValueError):
            load_config(path)

    def test_supported_extensions(self):
        for extension in ("ppt", "pptx", "pptm", "pps", "ppsx", "ppsm"):
            self.assertTrue(supported("test." + extension))
        self.assertFalse(supported("test.pdf"))


if __name__ == "__main__":
    unittest.main()
