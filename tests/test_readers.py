from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from pptsave.detect import _commandline_pdfs, _existing_path
from pptsave.originals import title_matches_name
from pptsave.readers import local_reference, select_processes


def process(pid, name, parent=0, args=()):
    return SimpleNamespace(pid=pid, info={"pid": pid, "name": name, "ppid": parent},
                           cmdline=lambda: [name, *args])


class PDFReaderTests(unittest.TestCase):
    def test_local_file_url_decodes_once_and_removes_viewer_fragment(self):
        self.assertEqual(local_reference("file:///D:/%E6%8A%A5%E5%91%8A/a%20b%23c.pdf#page=3"), "D:/报告/a b#c.pdf")
        self.assertEqual(local_reference("file://localhost/C:/a.pdf"), "C:/a.pdf")
        self.assertEqual(local_reference("file://server/share/a.pdf"), "//server/share/a.pdf")
        self.assertEqual(local_reference(r"D:\PDF\a#b.pdf"), r"D:\PDF\a#b.pdf")

    def test_online_urls_and_browser_options_are_rejected(self):
        for value in ("https://example.com/report.pdf", "http://example.com/a.pdf", "blob:https://example.com/1", "--profile-directory=Default"):
            self.assertIsNone(local_reference(value))

    def test_desktop_pdf_readers_and_wps_are_selected(self):
        apps = [process(1, "wpsufd.exe"), process(2, "AcroRd32.exe"), process(3, "SumatraPDF.exe"), process(4, "POWERPNT.EXE")]
        self.assertEqual({p.pid for p in select_processes(apps, [])}, {1, 2, 3, 4})

    def test_general_browser_windows_are_not_scanned(self):
        apps = [process(1, "msedge.exe"), process(2, "msedge.exe", parent=1), process(3, "chrome.exe")]
        windows = [{"pid": 1, "title": "新闻 - Microsoft Edge"}, {"pid": 3, "title": "Home - Google Chrome"}]
        self.assertEqual(select_processes(apps, windows), [])

    def test_pdf_browser_root_and_its_children_are_selected(self):
        apps = [process(1, "msedge.exe"), process(2, "msedge.exe", parent=1), process(3, "msedge.exe", parent=2), process(4, "chrome.exe")]
        chosen = select_processes(apps, [{"pid": 1, "title": "报告.pdf - Microsoft Edge"}])
        self.assertEqual({p.pid for p in chosen}, {1, 2, 3})

    def test_browser_pdf_launch_argument_can_enable_handle_detection(self):
        app = process(1, "msedge.exe", args=("file:///D:/报告.pdf",))
        self.assertEqual(select_processes([app], [{"pid": 1, "title": "演示标题 - Microsoft Edge"}]), [app])

    def test_pdf_titles_from_multiple_readers_match_original(self):
        for title in ("报告.pdf - WPS PDF", "报告.pdf - Adobe Acrobat Reader", "报告.pdf - Microsoft Edge", "报告.pdf - 工作 - Microsoft Edge", "报告.pdf - SumatraPDF", "报告.pdf - Foxit PDF Reader"):
            self.assertTrue(title_matches_name(title, "报告.pdf"), title)

    def test_launch_path_requires_current_matching_window(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "报告.pdf"
            path.write_bytes(b"%PDF-1.4")
            app = process(1, "msedge.exe", args=(path.as_uri(),))
            self.assertEqual(_existing_path(path.as_uri()), str(path.resolve()))
            self.assertEqual(_commandline_pdfs(app, ["报告.pdf - Microsoft Edge"]), [str(path.resolve())])
            self.assertEqual(_commandline_pdfs(app, ["新闻 - Microsoft Edge"]), [])

    def test_edge_formatting_character_does_not_hide_pdf_name(self):
        self.assertTrue(title_matches_name("PDF展示测试.pdf - Profile 1 - Microsoft\u200b Edge", "PDF展示测试.pdf"))
        self.assertFalse(title_matches_name("PDF展示测试2.pdf - Profile 1 - Microsoft\u200b Edge", "PDF展示测试.pdf"))
