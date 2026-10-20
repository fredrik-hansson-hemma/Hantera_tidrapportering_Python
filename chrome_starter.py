"""
Startar Chrome med fjärrfelsökning (CDP) om den inte redan är igång.

Chrome startas fristående i bakgrunden med en egen profil i chrome_profile/,
så att Antura-inloggningen sparas och fönstret ligger kvar när skriptet är klart.
"""

import os
import socket
import subprocess
import time
from pathlib import Path

from config import ANTURA_URL, CDP_PORT, CHROME_PROFILE_DIR


def find_chrome() -> Path | None:
    """Letar efter chrome.exe på de vanliga installationsställena."""
    candidates = [
        Path(os.environ.get(var, "")) / "Google" / "Chrome" / "Application" / "chrome.exe"
        for var in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA")
        if os.environ.get(var)
    ]
    return next((p for p in candidates if p.exists()), None)


def is_port_open(port: int) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False


def ensure_chrome(
    port: int = CDP_PORT,
    profile_dir: Path = CHROME_PROFILE_DIR,
    url: str = ANTURA_URL,
    timeout_s: float = 20,
) -> None:
    """
    Säkerställer att en Chrome med CDP lyssnar på given port.
    Startar en ny om ingen finns. Avslutar programmet om det inte går.
    """
    if is_port_open(port):
        print(f"  Chrome är redan igång (port {port})")
        return

    chrome = find_chrome()
    if chrome is None:
        raise SystemExit(
            "❌ Hittade inte chrome.exe på något standardställe.\n"
            "   Installera Chrome eller lägg till sökvägen i chrome_starter.find_chrome()."
        )

    print(f"  Startar Chrome (port {port}, profil {profile_dir.name})…")
    subprocess.Popen(
        [
            str(chrome),
            f"--remote-debugging-port={port}",
            f"--user-data-dir={profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            url,
        ],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        # Fristående process: Chrome lever vidare när skriptet avslutas
        creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
    )

    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        if is_port_open(port):
            return
        time.sleep(0.5)

    raise SystemExit(
        f"❌ Chrome startade men svarar inte på port {port}.\n"
        "   Stäng alla Chrome-fönster som använder chrome_profile och försök igen."
    )
