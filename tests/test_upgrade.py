"""Upgrade regressions: existing choices and saved-copy histories must survive."""
from contextlib import nullcontext
from pathlib import Path
import os
import sys
import tempfile
import types
import unittest
from unittest.mock import patch

from pptsave.archive import Archive
from pptsave.settings import load_config, save_config, state_directory, migrate_startup, set_startup


class UpgradeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_old_settings_and_logs_move_to_new_directory(self):
        legacy = self.root / 'ZhitingPPTSave'
        save_config(legacy / 'config.json', {'destination': 'D:/会议/副本'})
        (legacy / 'app.log').write_text('old log')
        with patch.dict(os.environ, {'LOCALAPPDATA': str(self.root)}):
            current = state_directory()
            self.assertEqual(state_directory(), current)
        self.assertEqual(current.name, 'PPTSave')
        self.assertEqual(load_config(current / 'config.json')['destination'], 'D:/会议/副本')
        self.assertEqual((current / 'app.log').read_text(), 'old log')
        self.assertFalse(legacy.exists())

    def test_existing_new_settings_take_priority(self):
        save_config(self.root / 'ZhitingPPTSave/config.json', {'destination': 'old'})
        save_config(self.root / 'PPTSave/config.json', {'destination': 'new'})
        with patch.dict(os.environ, {'LOCALAPPDATA': str(self.root)}):
            self.assertEqual(load_config(state_directory() / 'config.json')['destination'], 'new')

    def test_partial_upgrade_recovers_missing_config(self):
        save_config(self.root / 'ZhitingPPTSave/config.json', {'destination': 'old'})
        (self.root / 'PPTSave').mkdir()
        with patch.dict(os.environ, {'LOCALAPPDATA': str(self.root)}):
            self.assertEqual(load_config(state_directory() / 'config.json')['destination'], 'old')

    def test_legacy_index_prevents_duplicate_after_file_renamed(self):
        source = self.root / '报告.pptx'
        source.write_bytes(b'existing content')
        destination = self.root / 'copies'
        archive = Archive(destination)
        first = archive.archive(source)
        archive.close()
        (destination / '.pptsave').rename(destination / '.zhiting')
        renamed = self.root / 'new-name.pptx'
        source.rename(renamed)
        archive = Archive(destination)
        try:
            result = archive.archive(renamed)
            self.assertEqual(result.status, 'duplicate')
            self.assertEqual(result.destination, first.destination)
            self.assertFalse((destination / '.zhiting').exists())
            self.assertEqual(len(list(destination.rglob('*.pptx'))), 1)
        finally:
            archive.close()

    def test_partial_upgrade_merges_both_indexes_without_losing_copies(self):
        destination = self.root / 'copies'
        old = self.root / 'old.pdf'
        new = self.root / 'new.pdf'
        old.write_bytes(b'old pdf')
        new.write_bytes(b'new pdf')
        archive = Archive(destination)
        first = archive.archive(old)
        archive.close()
        # Stage the legacy history outside the destination until a new index exists.
        staged = self.root / 'legacy'
        (destination / '.pptsave').rename(staged)
        archive = Archive(destination)
        second = archive.archive(new)
        archive.close()
        staged.rename(destination / '.zhiting')
        archive = Archive(destination)
        try:
            old.rename(self.root / 'renamed.pdf')
            self.assertEqual(archive.archive(self.root / 'renamed.pdf').destination, first.destination)
            self.assertEqual(archive.archive(new).destination, second.destination)
            self.assertEqual(len(list(destination.rglob('*.pdf'))), 2)
        finally:
            archive.close()


class StartupUpgradeTests(unittest.TestCase):
    def setUp(self):
        self.values = {}
        def query(key, name):
            if name not in self.values:
                raise FileNotFoundError(name)
            return self.values[name], 1
        def delete(key, name):
            if name not in self.values:
                raise FileNotFoundError(name)
            del self.values[name]
        registry = types.SimpleNamespace(
            HKEY_CURRENT_USER=1, REG_SZ=1,
            OpenKey=lambda *args: nullcontext(1), CreateKey=lambda *args: nullcontext(1),
            QueryValueEx=query, DeleteValue=delete,
            SetValueEx=lambda key, name, reserved, kind, value: self.values.update({name: value}))
        self.patch = patch.dict(sys.modules, {'winreg': registry})
        self.patch.start()
        self.addCleanup(self.patch.stop)

    def test_disabled_startup_stays_disabled(self):
        self.assertFalse(migrate_startup())
        self.assertEqual(self.values, {})

    def test_legacy_startup_moves_to_current_executable(self):
        self.values['ZhitingPPTSave'] = 'old executable --background'
        with patch.object(sys, 'frozen', True, create=True), patch.object(sys, 'executable', 'C:/Tools/PPTSave/PPTSave.exe'):
            self.assertTrue(migrate_startup())
        self.assertEqual(self.values, {'PPTSave': 'C:/Tools/PPTSave/PPTSave.exe --background'})

    def test_new_startup_refreshes_location_and_cleans_old_entry(self):
        self.values.update(PPTSave='previous path', ZhitingPPTSave='old path')
        with patch.object(sys, 'frozen', True, create=True), patch.object(sys, 'executable', 'C:/New Folder/PPTSave.exe'):
            self.assertTrue(migrate_startup())
        self.assertEqual(self.values, {'PPTSave': '"C:/New Folder/PPTSave.exe" --background'})
        set_startup(False)
        self.assertEqual(self.values, {})

    def test_failed_replacement_keeps_legacy_startup(self):
        self.values['ZhitingPPTSave'] = 'old path'
        with patch.object(sys.modules['winreg'], 'SetValueEx', side_effect=PermissionError):
            with self.assertRaises(PermissionError):
                migrate_startup()
        self.assertEqual(self.values, {'ZhitingPPTSave': 'old path'})
