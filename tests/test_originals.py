from pathlib import Path
import tempfile
import unittest
from zhiting.originals import prefer_originals, title_matches_name


class OriginalNameTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def file(self, name, content=b"same"):
        path = self.root / name
        path.write_bytes(content)
        return str(path)

    def test_open_original_precedes_temp_and_removes_duplicate(self):
        temp = self.file("~tmp123.pptx")
        original = self.file("报告.pptx")
        paths, pending, recovered = prefer_originals({temp: "handle", original: "COM"})
        self.assertEqual(list(paths), [original])
        self.assertEqual(pending, [])
        self.assertEqual(recovered[0]["original"], original)

    def test_history_requires_exact_content_match(self):
        temp = self.file("~tmp123.pptx")
        different = self.file("错误名称.pptx", b"else")
        original = self.file("正确名称.pptx")
        paths, pending, recovered = prefer_originals({temp: "handle"}, [different, original])
        self.assertEqual(list(paths), [original])
        self.assertFalse(pending)

    def test_ambiguous_names_are_not_guessed(self):
        temp = self.file("~tmp123.pptx")
        one = self.file("one.pptx")
        two = self.file("two.pptx")
        paths, pending, recovered = prefer_originals({temp: "handle"}, [one, two])
        self.assertEqual(list(paths), [temp])
        self.assertEqual(pending, [temp])
        self.assertFalse(recovered)

    def test_missing_original_keeps_temporary_copy(self):
        temp = self.file("~tmp123.pptx")
        paths, pending, _ = prefer_originals({temp: "handle"}, [str(self.root / "missing.pptx")])
        self.assertEqual(pending, [temp])
        self.assertIn(temp, paths)

    def test_title_match_is_exact_not_prefix(self):
        for title in ("报告.pptx - WPS Office", "报告 - WPS 演示", "报告 [只读] - PowerPoint", "报告.pptx"):
            self.assertTrue(title_matches_name(title, "报告.pptx"), title)
        for title in ("报告2.pptx - WPS Office", "旧报告.pptx", "报告.docx - WPS Office"):
            self.assertFalse(title_matches_name(title, "报告.pptx"), title)
