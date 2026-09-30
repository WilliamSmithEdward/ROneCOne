"""The YARA Forge download contract, and the pin file the weekly updater moves."""

from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))
import yara_forge_rules  # noqa: E402

URL = "https://github.com/YARAHQ/yara-forge/releases/download/20260726/yara-forge-rules-core.zip"


class YaraForgeRuleTests(unittest.TestCase):
    def test_download_rejects_changed_content(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pin = root / "yara.json"
            pin.write_text(json.dumps({"yara_forge": {
                "release": "20260726", "asset": "yara-forge-rules-core.zip", "url": URL, "sha256": "a" * 64}}))
            output = root / "rules.zip"
            response = MagicMock()
            response.__enter__.return_value.read.return_value = b"changed"
            with (patch.object(yara_forge_rules, "PIN", pin),
                  patch.object(yara_forge_rules.urllib.request, "urlopen", return_value=response)):
                with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
                    yara_forge_rules.download(output)
            self.assertFalse(output.exists())

    def test_the_pin_file_is_one_the_standard_updater_accepts(self):
        spec = importlib.util.spec_from_file_location(
            "yara_update", ROOT / ".github" / "security" / "yara_update.py")
        updater = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(updater)
        pins = json.loads((ROOT / ".github" / "security" / "yara.json").read_text(encoding="utf-8"))
        updater.check_move(pins, pins)
        self.assertEqual(yara_forge_rules.read_pin(), pins["yara_forge"])


if __name__ == "__main__":
    unittest.main()
