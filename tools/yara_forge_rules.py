"""Download a reviewed YARA Forge Core package or propose a newer pin."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import urllib.request
import zipfile
from pathlib import Path

PIN = Path(__file__).with_name("yara_forge_pin.json")
ASSET = "yara-forge-rules-core.zip"
MAX_BYTES = 50 * 1024 * 1024


def valid_release(release: str) -> bool:
    return bool(re.fullmatch(r"20\d{6}", release))


def read_pin() -> dict[str, str]:
    pin = json.loads(PIN.read_text(encoding="utf-8"))
    if (set(pin) != {"release", "sha256"} or not valid_release(pin["release"])
            or not re.fullmatch(r"[0-9a-f]{64}", pin["sha256"])):
        raise ValueError("invalid YARA Forge pin")
    return pin


def download(output: Path) -> None:
    pin = read_pin()
    url = ("https://github.com/YARAHQ/yara-forge/releases/download/"
           f"{pin['release']}/{ASSET}")
    request = urllib.request.Request(url, headers={"User-Agent": "ROneCOne-security-ci"})
    with urllib.request.urlopen(request, timeout=45) as response:
        data = response.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise ValueError("YARA Forge archive exceeds the 50 MiB limit")
    digest = hashlib.sha256(data).hexdigest()
    if digest != pin["sha256"]:
        raise ValueError(f"YARA Forge archive SHA-256 mismatch: {digest}")
    output.write_bytes(data)
    print(f"YARA Forge {pin['release']}: SHA-256 {digest}")


def update(release: str, archive: Path) -> bool:
    if not valid_release(release):
        raise ValueError("YARA Forge release must be an eight-digit date")
    current = read_pin()
    if release <= current["release"]:
        return False
    if archive.stat().st_size > MAX_BYTES:
        raise ValueError("YARA Forge archive exceeds the 50 MiB limit")
    with zipfile.ZipFile(archive) as package:
        names = [name for name in package.namelist() if name.lower().endswith(".yar")]
        if (len(names) != 1 or any(item.file_size > MAX_BYTES for item in package.infolist())
                or package.testzip() is not None):
            raise ValueError("YARA Forge archive is invalid or has an unexpected rule layout")
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    PIN.write_text(json.dumps({"release": release, "sha256": digest}, indent=2) + "\n",
                   encoding="utf-8", newline="\n")
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    get = sub.add_parser("download")
    get.add_argument("output", type=Path)
    bump = sub.add_parser("update")
    bump.add_argument("release")
    bump.add_argument("archive", type=Path)
    args = parser.parse_args()
    try:
        if args.command == "download":
            download(args.output)
        else:
            print(f"changed={'true' if update(args.release, args.archive) else 'false'}")
        return 0
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"YARA Forge {args.command} failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
