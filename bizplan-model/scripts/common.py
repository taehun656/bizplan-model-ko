# -*- coding: utf-8 -*-
"""UTF-8 file output and ASCII-only console text."""
from pathlib import Path


def write_text(path, text):
    dest = Path(path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(text, encoding="utf-8", newline="\n")


def say(text):
    print(str(text).encode("ascii", "replace").decode("ascii"))


def wrote(path, label):
    dest = Path(path)
    size = dest.stat().st_size if dest.exists() else 0
    token = "".join(ch for ch in str(label) if ch.isascii() and (ch.isalnum() or ch in "._-")) or "file"
    say(f"WROTE {token} ({size} bytes)")
