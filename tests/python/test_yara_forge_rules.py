"""Pinning and download contracts for the scheduled YARA Forge updater."""

from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import yara_forge_rules  # noqa: E402


class YaraForgeRuleTests(unittest.TestCase):
    def test_update_only_newer_valid_package(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pin = root / "pin.json"
            pin.write_text(json.dumps({"release": "20260726", "sha256": "a" * 64}))
            archive = root / "rules.zip"
            with zipfile.ZipFile(archive, "w") as package:
                package.writestr("packages/core/rules.yar", "rule test { condition: true }")
            with patch.object(yara_forge_rules, "PIN", pin):
                self.assertFalse(yara_forge_rules.update("20260726", archive))
                self.assertTrue(yara_forge_rules.update("20260927", archive))
                self.assertEqual(yara_forge_rules.read_pin(), {
                    "release": "20260927",
                    "sha256": hashlib.sha256(archive.read_bytes()).hexdigest(),
                })

    def test_rejects_unexpected_archive_without_changing_pin(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pin = root / "pin.json"
            original = json.dumps({"release": "20260726", "sha256": "a" * 64})
            pin.write_text(original)
            archive = root / "rules.zip"
            with zipfile.ZipFile(archive, "w") as package:
                package.writestr("unexpected.txt", "no rules")
            with patch.object(yara_forge_rules, "PIN", pin):
                with self.assertRaisesRegex(ValueError, "unexpected rule layout"):
                    yara_forge_rules.update("20260927", archive)
            self.assertEqual(pin.read_text(), original)

    def test_download_rejects_changed_content(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pin = root / "pin.json"
            pin.write_text(json.dumps({"release": "20260726", "sha256": "a" * 64}))
            output = root / "rules.zip"
            response = unittest.mock.MagicMock()
            response.__enter__.return_value.read.return_value = b"changed"
            with (patch.object(yara_forge_rules, "PIN", pin),
                  patch.object(yara_forge_rules.urllib.request, "urlopen",
                               return_value=response)):
                with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
                    yara_forge_rules.download(output)
            self.assertFalse(output.exists())
