# Created by Jeoff Krontz via claude on 3/15/2026
#
# This script monitors a NTFY feed which is being updated by Google webhook when a calendar changes
# When detected, this script will switch focus to outlook and hit key sequence ALT+3
# - This is to forces Gsyncit to refresh the calendar sync.
# - Gsyncit doesn't have a method for that, so the Outlook plugin calendar function must be added to quick-task bar in Outlook.
# - It is assigned Alt+3 as the hotkey if it is in the correct position (3) on the quick-task bar.
#
# Revision History:
# ---------------------------------
# 03/15/2026: First working version
# 03/22/2026: Added notes at top and started revision history
#
#!/usr/bin/env python3
"""
ntfy → GSyncIt Monitor
======================
Monitors an ntfy topic for Google Calendar webhook notifications,
triggers GSyncIt on each new message, and lives in the system tray
with a right-click menu to exit.

Requirements:
    pip install requests pystray pillow pyautogui

Build as .exe (run on Windows):
    pip install pyinstaller
    pyinstaller --onefile --noconsole ^
      --collect-data certifi ^
      --hidden-import pyautogui ^
      --hidden-import pystray ^
      --hidden-import PIL ^
      --hidden-import requests ^
      ntfy_gsyncit_monitor.py

Usage:
    python ntfy_gsyncit_monitor.py
    python ntfy_gsyncit_monitor.py --dry-run
"""

import argparse
import ctypes
import ctypes.wintypes
import json
import logging
import os
import sys
import threading
import time
from pathlib import Path


# ── PyInstaller SSL fix ───────────────────────────────────────────────────────
import certifi
os.environ["SSL_CERT_FILE"] = certifi.where()
os.environ["REQUESTS_CA_BUNDLE"] = certifi.where()

# ┌─────────────────────────────────────────────────────────────────────────┐
# │                        CONFIGURATION DEFAULTS                           │
# │  These values are used if no gsyncit_monitor.ini file is found.         │
# │  Create gsyncit_monitor.ini next to the .exe to override any of these.  │
# └─────────────────────────────────────────────────────────────────────────┘

NTFY_TOPIC_URL     = "https://ntfy.sh/jeoff_calendar_987"
OUTLOOK_SYNC_KEY   = "alt+3"
SYNC_DELAY_SECONDS = 3
DEBOUNCE_SECONDS   = 15
LOG_FILE           = None   # None = auto: next to the .exe

# ┌─────────────────────────────────────────────────────────────────────────┐
# │                        END CONFIGURATION DEFAULTS                       │
# └─────────────────────────────────────────────────────────────────────────┘


def load_config():
    """
    Load gsyncit_monitor.ini from the same directory as the exe/script.

    Example gsyncit_monitor.ini:
    ─────────────────────────────
    [monitor]
    ntfy_topic_url     = https://ntfy.sh/my_other_topic
    outlook_sync_key   = alt+3
    sync_delay_seconds = 5
    debounce_seconds   = 30
    log_file           = C:\\Logs\\gsyncit_monitor.log
    ─────────────────────────────
    """
    import configparser
    global NTFY_TOPIC_URL, OUTLOOK_SYNC_KEY, SYNC_DELAY_SECONDS, DEBOUNCE_SECONDS, LOG_FILE

    exe_dir = Path(os.path.dirname(os.path.abspath(sys.argv[0])))
    ini_path = exe_dir / "gsyncit_monitor.ini"

    if LOG_FILE is None:
        LOG_FILE = exe_dir / "ntfy_gsyncit_monitor.log"

    if not ini_path.exists():
        return

    cfg = configparser.ConfigParser()
    cfg.read(ini_path, encoding="utf-8")
    section = "monitor"
    if not cfg.has_section(section):
        print(f"Warning: {ini_path} has no [monitor] section — using defaults.")
        return

    def get(key):
        return cfg.get(section, key, fallback=None)

    if (v := get("ntfy_topic_url"))    is not None: NTFY_TOPIC_URL     = v
    if (v := get("outlook_sync_key"))  is not None: OUTLOOK_SYNC_KEY   = v
    if (v := get("log_file"))          is not None: LOG_FILE            = Path(v)
    if (v := get("sync_delay_seconds")) is not None:
        try:    SYNC_DELAY_SECONDS = int(v)
        except ValueError: print(f"Warning: invalid sync_delay_seconds '{v}', using default.")
    if (v := get("debounce_seconds"))  is not None:
        try:    DEBOUNCE_SECONDS = int(v)
        except ValueError: print(f"Warning: invalid debounce_seconds '{v}', using default.")

    print(f"Loaded config from {ini_path}")


load_config()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ── Tray icon ─────────────────────────────────────────────────────────────────

def make_tray_image(color="royalblue"):
    from PIL import Image, ImageDraw
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, size - 4, size - 4), fill=color)
    return img

def make_tray_image_error():    return make_tray_image("crimson")
def make_tray_image_syncing():  return make_tray_image("mediumseagreen")


# ── GSyncIt trigger ───────────────────────────────────────────────────────────

def find_outlook_hwnd():
    user32 = ctypes.windll.user32
    found = ctypes.wintypes.HWND(0)

    def enum_cb(hwnd, _):
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        if "Outlook" in buf.value and user32.IsWindowVisible(hwnd):
            found.value = hwnd
            return False
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_cb), 0)
    return found.value


def trigger_gsyncit(dry_run: bool) -> bool:
    if dry_run:
        log.info(f"[DRY RUN] Would send '{OUTLOOK_SYNC_KEY}' to Outlook")
        return True

    hwnd = find_outlook_hwnd()
    if not hwnd:
        log.error("Outlook window not found — is Outlook open?")
        return False

    user32 = ctypes.windll.user32

    try:
        buf = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(hwnd, buf, 256)
        log.info(f"Found Outlook: hwnd={hwnd}, title='{buf.value}'")

        # Check current state so we don't change maximized → windowed
        class WINDOWPLACEMENT(ctypes.Structure):
            _fields_ = [("length",           ctypes.wintypes.UINT),
                        ("flags",            ctypes.wintypes.UINT),
                        ("showCmd",          ctypes.wintypes.UINT),
                        ("ptMinPosition",    ctypes.wintypes.POINT),
                        ("ptMaxPosition",    ctypes.wintypes.POINT),
                        ("rcNormalPosition", ctypes.wintypes.RECT)]
        wp = WINDOWPLACEMENT()
        wp.length = ctypes.sizeof(WINDOWPLACEMENT)
        user32.GetWindowPlacement(hwnd, ctypes.byref(wp))
        if wp.showCmd == 2:    # SW_SHOWMINIMIZED — restore it
            user32.ShowWindow(hwnd, 9)
        elif wp.showCmd == 3:  # SW_SHOWMAXIMIZED — keep it maximized
            user32.ShowWindow(hwnd, 3)
        else:                  # normal — just bring to front
            user32.ShowWindow(hwnd, 5)
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)

        parts    = OUTLOOK_SYNC_KEY.lower().split("+")
        digit    = [p for p in parts if p.isdigit()][0]
        vk_digit = 0x30 + int(digit)
        VK_ALT   = 0x12
        KEYEVENTF_KEYUP = 0x0002
        user32.keybd_event(VK_ALT,   0, 0,              0)
        user32.keybd_event(vk_digit, 0, 0,              0)
        user32.keybd_event(vk_digit, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_ALT,   0, KEYEVENTF_KEYUP, 0)
        log.info(f"Sent '{OUTLOOK_SYNC_KEY}' to Outlook — GSyncIt sync triggered.")
        return True

    except Exception as exc:
        log.error(f"Failed to trigger GSyncIt: {exc}")
        return False


# ── ntfy helpers ──────────────────────────────────────────────────────────────

def get_latest_message_id():
    try:
        import requests
        url = NTFY_TOPIC_URL.rstrip("/") + "/json?poll=1&since=all&limit=1"
        resp = requests.get(url, timeout=10)
        if resp.status_code == 404 or not resp.text.strip():
            return ""
        lines = [l.strip() for l in resp.text.strip().splitlines() if l.strip()]
        if not lines:
            return ""
        msg = json.loads(lines[-1])
        latest_id = msg.get("id", "")
        log.info(f"Startup: latest existing message ID is '{latest_id}' — will only process newer messages.")
        return latest_id
    except Exception as exc:
        log.warning(f"Could not fetch latest message ID: {exc} — will ignore messages for first {DEBOUNCE_SECONDS}s.")
        return None


# ── Monitor thread ────────────────────────────────────────────────────────────

def stream_loop(stop_event, last_sync_at, startup_message_id, dry_run, tray_ref):
    import requests
    sse_url = NTFY_TOPIC_URL.rstrip("/") + "/sse"

    while not stop_event.is_set():
        try:
            with requests.get(sse_url, stream=True, timeout=90) as resp:
                resp.raise_for_status()
                for raw_line in resp.iter_lines(decode_unicode=True):
                    if stop_event.is_set():
                        return
                    if not raw_line or not raw_line.startswith("data:"):
                        continue
                    data_str = raw_line[len("data:"):].strip()
                    if not data_str:
                        continue
                    try:
                        msg = json.loads(data_str)
                    except json.JSONDecodeError:
                        continue
                    if msg.get("event") != "message":
                        continue

                    event_msg = msg.get("message", "").strip() or "(no body)"
                    msg_id    = msg.get("id", "")

                    if startup_message_id is None:
                        if time.monotonic() < last_sync_at[0] + DEBOUNCE_SECONDS:
                            log.info("Startup cooldown: skipping message.")
                            continue
                    elif msg_id == startup_message_id or msg_id == "":
                        log.info(f"Startup: skipping pre-existing message (id={msg_id}).")
                        continue

                    log.info(f"New notification: {event_msg}")

                    now = time.monotonic()
                    since_last = now - last_sync_at[0]
                    if since_last < DEBOUNCE_SECONDS:
                        log.info(f"Debounced — last sync was {since_last:.0f}s ago, skipping.")
                        continue

                    if SYNC_DELAY_SECONDS > 0:
                        log.info(f"Waiting {SYNC_DELAY_SECONDS}s before syncing...")
                        for _ in range(SYNC_DELAY_SECONDS * 10):
                            if stop_event.is_set():
                                return
                            time.sleep(0.1)

                    tray = tray_ref[0]
                    if tray:
                        try:
                            tray.icon  = make_tray_image_syncing()
                            tray.title = "Calendar Monitor — syncing..."
                        except Exception:
                            pass

                    if trigger_gsyncit(dry_run):
                        last_sync_at[0] = time.monotonic()

                    if tray:
                        try:
                            tray.icon  = make_tray_image()
                            tray.title = "Calendar Monitor — listening"
                        except Exception:
                            pass

        except requests.exceptions.ConnectionError as exc:
            if not stop_event.is_set():
                log.warning(f"Connection lost: {exc} — retrying in 10s...")
                tray = tray_ref[0]
                if tray:
                    try:
                        tray.icon  = make_tray_image_error()
                        tray.title = "Calendar Monitor — connection lost"
                    except Exception:
                        pass
                stop_event.wait(timeout=10)
                if tray:
                    try:
                        tray.icon  = make_tray_image()
                        tray.title = "Calendar Monitor — listening"
                    except Exception:
                        pass
        except requests.exceptions.Timeout:
            log.debug("SSE stream timed out (normal), reconnecting...")
        except Exception as exc:
            if not stop_event.is_set():
                log.error(f"Unexpected error: {exc} — retrying in 15s...")
                stop_event.wait(timeout=15)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="ntfy → GSyncIt tray monitor")
    parser.add_argument("--dry-run", action="store_true",
                        help="Log sync triggers without actually running GSyncIt")
    args = parser.parse_args()

    try:
        import requests
    except ImportError:
        sys.exit("Missing package: pip install requests pystray pillow pyautogui")

    try:
        import pystray
        from PIL import Image
    except ImportError:
        sys.exit("Missing packages: pip install pystray pillow")

    log.info("Starting ntfy → GSyncIt monitor")
    if args.dry_run:
        log.info("DRY RUN mode — GSyncIt will not actually be called")

    startup_message_id = get_latest_message_id()
    stop_event   = threading.Event()
    last_sync_at = [0.0]
    tray_ref     = [None]

    t = threading.Thread(
        target=stream_loop,
        args=(stop_event, last_sync_at, startup_message_id, args.dry_run, tray_ref),
        daemon=True,
    )
    t.start()

    def on_exit(icon, item):
        log.info("Exit requested from tray menu.")
        stop_event.set()
        icon.stop()

    def on_open_log(icon, item):
        os.startfile(str(LOG_FILE))

    def on_force_sync(icon, item):
        log.info("Manual sync requested from tray menu.")
        def run():
            try:
                icon.icon  = make_tray_image_syncing()
                icon.title = "Calendar Monitor — syncing..."
            except Exception:
                pass
            trigger_gsyncit(args.dry_run)
            last_sync_at[0] = time.monotonic()
            try:
                icon.icon  = make_tray_image()
                icon.title = "Calendar Monitor — listening"
            except Exception:
                pass
        threading.Thread(target=run, daemon=True).start()

    menu = pystray.Menu(
        pystray.MenuItem("Calendar Monitor", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Force sync now", on_force_sync),
        pystray.MenuItem("Open log file",  on_open_log),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Exit", on_exit),
    )

    icon = pystray.Icon(
        name="ntfy_gsyncit",
        icon=make_tray_image(),
        title="Calendar Monitor — listening",
        menu=menu,
    )
    tray_ref[0] = icon

    log.info(f"Tray icon started. Log file: {LOG_FILE}")
    icon.run()

    stop_event.set()
    t.join(timeout=3)
    log.info("Exited cleanly.")


if __name__ == "__main__":
    main()
