#!/usr/bin/env python3
"""Watch source files and rebuild the website data when they change."""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WATCHED = [
    ROOT / "calottino_categoria_gestione.xlsx",
    ROOT / "statuto_calottino_categoria_avionici.docx",
    ROOT / "WhatsApp Image 2026-06-16 at 15.08.20.jpeg",
]
BUILDER = ROOT / "scripts" / "build_site_data.py"


def snapshot() -> dict[Path, float]:
    return {path: path.stat().st_mtime for path in WATCHED if path.exists()}


def rebuild() -> None:
    subprocess.run([sys.executable, str(BUILDER)], cwd=ROOT, check=True)


def main() -> None:
    print("Monitoraggio attivo. Premi Ctrl+C per fermare.")
    rebuild()
    last = snapshot()

    while True:
        time.sleep(2)
        current = snapshot()
        if current != last:
            print("Modifica rilevata, aggiorno i dati del sito...")
            rebuild()
            last = current


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nMonitoraggio interrotto.")
