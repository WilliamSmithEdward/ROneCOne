"""Download the reviewed YARA Forge Core package the malware scan runs.

The pin is the `yara_forge` entry of .github/security/yara.json, which the
standard YARA updater (.github/security/yara_update.py) moves.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import urllib.request
from pathlib import Path

PIN = Path(__file__).resolve().parents[1] / ".github" / "security" / "yara.json"
ASSET = "yara-forge-rules-core.zip"
MAX_BYTES = 50 * 1024 * 1024


def read_pin() -> dict[str, str]:
    pin = json.loads(PIN.read_text(encoding="utf-8"))["yara_forge"]
    url = f"https://github.com/YARAHQ/yara-forge/releases/download/{pin['release']}/{ASSET}"
    if (not re.fullmatch(r"20\d{6}", pin["release"]) or pin["url"] != url
            or not re.fullmatch(r"[0-9a-f]{64}", pin["sha256"])):
        raise ValueError("invalid YARA Forge pin")
    return pin


def download(output: Path) -> None:
    pin = read_pin()
    request = urllib.request.Request(pin["url"], headers={"User-Agent": "ROneCOne-security-ci"})
    with urllib.request.urlopen(request, timeout=45) as response:
        data = response.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise ValueError("YARA Forge archive exceeds the 50 MiB limit")
    digest = hashlib.sha256(data).hexdigest()
    if digest != pin["sha256"]:
        raise ValueError(f"YARA Forge archive SHA-256 mismatch: {digest}")
    output.write_bytes(data)
    print(f"YARA Forge {pin['release']}: SHA-256 {digest}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    get = sub.add_parser("download")
    get.add_argument("output", type=Path)
    args = parser.parse_args()
    try:
        download(args.output)
        return 0
    except (OSError, ValueError) as error:
        print(f"YARA Forge {args.command} failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
