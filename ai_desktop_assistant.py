"""
AI Desktop Assistant  (text + voice)
====================================
Designed for Windows 10/11 (most features also work on macOS/Linux).

Features
--------
1. Web search in the default browser     -> "search python tutorials", "google weather"
2. File operations (File Explorer)       -> create / delete / rename / copy / move / list / find / read / open
3. Open desktop apps                     -> "open notepad", "launch chrome", "start calculator"
4. Alarms & timers                       -> "set alarm for 6:30 am", "set timer for 10 minutes"
5. Calendar events                       -> "add event team meeting tomorrow at 3 pm"
6. System status analysis                -> "system status", "battery", "disk", "top processes"

Install dependencies
--------------------
    pip install psutil SpeechRecognition pyaudio pyttsx3

(psutil is required for system status. SpeechRecognition + pyaudio are only needed for
voice input; pyttsx3 only for spoken replies. The assistant still runs without them.)

Run
---
    python ai_desktop_assistant.py

Notes on Clock / Calendar
-------------------------
Windows does not offer a public API to create alarms inside the Clock app or events
inside the Calendar app programmatically. So this assistant:
  * runs its own alarm/timer scheduler (beeps, speaks and prints when time is up,
    works while the assistant is running) and also opens the Windows Clock app, and
  * creates a standard .ics calendar file for each event and opens it, which hands
    it to your default calendar app (Outlook / Windows Calendar / Google, etc.)
    for you to press "Save".
"""

import difflib
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple
from urllib.parse import quote_plus

# --------------------------------------------------------------------------- #
# Optional dependencies (assistant degrades gracefully if they are missing)
# --------------------------------------------------------------------------- #
try:
    import psutil
except ImportError:
    psutil = None

try:
    import speech_recognition as sr
except ImportError:
    sr = None

try:
    import pyttsx3
except ImportError:
    pyttsx3 = None

IS_WINDOWS = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"

try:
    from send2trash import send2trash  # optional: delete to the Recycle Bin instead of forever
except ImportError:
    send2trash = None

import logging
from logging.handlers import RotatingFileHandler

__version__ = "1.0.0"

HOME = Path.home()
# AIDA_HOME lets tests / portable installs keep data somewhere else.
DATA_DIR = Path(os.environ.get("AIDA_HOME", HOME / ".ai_desktop_assistant"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
ALARM_FILE = DATA_DIR / "alarms.json"
EVENT_DIR = DATA_DIR / "events"
EVENT_DIR.mkdir(exist_ok=True)
CONFIG_FILE = DATA_DIR / "config.json"
LOG_FILE = DATA_DIR / "assistant.log"

# --------------------------------------------------------------------------- #
# Config + logging
# --------------------------------------------------------------------------- #
DEFAULT_CONFIG = {
    "speak_replies": False,        # speak every reply out loud
    "primary_language": "en-IN",   # first-choice speech language
    "secondary_language": "en-US", # second-opinion language
    "confirm_destructive": True,   # always ask before delete / overwrite
    "use_recycle_bin": True,       # delete -> Recycle Bin when send2trash is installed
    "max_capture_retries": 3,      # voice re-recording limit
    "log_transcripts": False,      # privacy: off by default, transcripts are NOT written to the log
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        if CONFIG_FILE.exists():
            cfg.update({k: v for k, v in json.loads(CONFIG_FILE.read_text()).items() if k in DEFAULT_CONFIG})
        else:
            CONFIG_FILE.write_text(json.dumps(DEFAULT_CONFIG, indent=2))
    except Exception:
        pass  # a broken config file must never stop the assistant from starting
    return cfg


CONFIG = load_config()

log = logging.getLogger("aida")
log.setLevel(logging.INFO)
if not log.handlers:
    try:
        _h = RotatingFileHandler(LOG_FILE, maxBytes=500_000, backupCount=2, encoding="utf-8")
        _h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(_h)
    except Exception:
        log.addHandler(logging.NullHandler())


def _known_folder(name):
    """Find a user folder, including OneDrive-redirected ones (e.g. ~/OneDrive/Desktop)."""
    for base in (HOME, HOME / "OneDrive"):
        p = base / name
        if p.exists():
            return p
    return HOME / name


# Friendly folder names -> real paths
PLACES = {
    "home": HOME,
    "desktop": _known_folder("Desktop"),
    "documents": _known_folder("Documents"),
    "downloads": _known_folder("Downloads"),
    "pictures": _known_folder("Pictures"),
    "music": _known_folder("Music"),
    "videos": _known_folder("Videos"),
}

# Spoken app names -> launch target
APPS = {
    "notepad": "notepad",
    "calculator": "calc" if IS_WINDOWS else "calculator",
    "calc": "calc",
    "paint": "mspaint",
    "word": "winword",
    "excel": "excel",
    "powerpoint": "powerpnt",
    "outlook": "outlook",
    "chrome": "chrome",
    "google chrome": "chrome",
    "edge": "msedge",
    "microsoft edge": "msedge",
    "firefox": "firefox",
    "brave": "brave",
    "vscode": "code",
    "vs code": "code",
    "visual studio code": "code",
    "file explorer": "explorer",
    "explorer": "explorer",
    "cmd": "cmd",
    "command prompt": "cmd",
    "terminal": "wt" if IS_WINDOWS else "gnome-terminal",
    "powershell": "powershell",
    "task manager": "taskmgr",
    "control panel": "control",
    "settings": "ms-settings:",
    "clock": "ms-clock:",
    "alarms": "ms-clock:",
    "calendar": "outlookcal:",
    "store": "ms-windows-store:",
    "snipping tool": "snippingtool",
    "spotify": "spotify",
    "whatsapp": "whatsapp:",
    "teams": "msteams:",
    "zoom": "zoom",
    "vlc": "vlc",
    "camera": "microsoft.windows.camera:",
    "photos": "ms-photos:",
    "maps": "bingmaps:",
    "wordpad": "write",
    "registry editor": "regedit",
    "device manager": "devmgmt.msc",
    "disk management": "diskmgmt.msc",
}

# Spoken site names (spaces removed) -> URL.  "go to you tube" == "go to youtube".
SITES = {
    "youtube": "https://www.youtube.com",
    "google": "https://www.google.com",
    "gmail": "https://mail.google.com",
    "github": "https://github.com",
    "wikipedia": "https://www.wikipedia.org",
    "twitter": "https://twitter.com",
    "linkedin": "https://www.linkedin.com",
    "facebook": "https://www.facebook.com",
    "instagram": "https://www.instagram.com",
    "reddit": "https://www.reddit.com",
    "amazon": "https://www.amazon.in",
    "flipkart": "https://www.flipkart.com",
    "netflix": "https://www.netflix.com",
    "stackoverflow": "https://stackoverflow.com",
    "maps": "https://maps.google.com",
    "googlemaps": "https://maps.google.com",
    "drive": "https://drive.google.com",
    "googledrive": "https://drive.google.com",
    "news": "https://news.google.com",
    "claude": "https://claude.ai",
    "whatsappweb": "https://web.whatsapp.com",
}

# Browser navigation by voice: (spoken phrases, keys to press, spoken feedback)
BROWSER_ACTIONS = [
    (("scroll down", "page down"), ("pagedown",), "Scrolling down"),
    (("scroll up", "page up"), ("pageup",), "Scrolling up"),
    (("scroll to top", "go to top", "top of page"), ("home",), "Jumping to the top"),
    (("scroll to bottom", "go to bottom", "bottom of page"), ("end",), "Jumping to the bottom"),
    (("new tab", "open new tab", "open a new tab"), ("ctrl", "t"), "Opening a new tab"),
    (("close tab", "close this tab", "close the tab"), ("ctrl", "w"), "Closing the tab"),
    (("next tab", "switch tab"), ("ctrl", "tab"), "Switching to the next tab"),
    (("previous tab", "last tab"), ("ctrl", "shift", "tab"), "Switching to the previous tab"),
    (("go back", "back", "previous page"), ("alt", "left"), "Going back"),
    (("go forward", "forward", "next page"), ("alt", "right"), "Going forward"),
    (("refresh", "reload", "refresh page", "reload page"), ("f5",), "Refreshing the page"),
    (("zoom in",), ("ctrl", "plus"), "Zooming in"),
    (("zoom out",), ("ctrl", "minus"), "Zooming out"),
    (("reset zoom",), ("ctrl", "0"), "Resetting zoom"),
    (("full screen", "fullscreen"), ("f11",), "Toggling full screen"),
    (("address bar", "focus address bar"), ("ctrl", "l"), "Focusing the address bar"),
    (("show history",), ("ctrl", "h"), "Opening history"),
    (("show downloads",), ("ctrl", "j"), "Opening downloads"),
]
BROWSER_PHRASES = {p for phrases, _, _ in BROWSER_ACTIONS for p in phrases}

# When True, skills only *describe* what they would do (used by --demo-noisy).
DRY_RUN = False

_VK = {
    "ctrl": 0x11, "shift": 0x10, "alt": 0x12, "tab": 0x09, "left": 0x25, "right": 0x27,
    "f5": 0x74, "f11": 0x7A, "pageup": 0x21, "pagedown": 0x22, "end": 0x23, "home": 0x24,
    "t": 0x54, "w": 0x57, "l": 0x4C, "h": 0x48, "j": 0x4A, "plus": 0xBB, "minus": 0xBD, "0": 0x30,
}
_EXTENDED = {"left", "right", "pageup", "pagedown", "end", "home"}


def focus_browser():
    """Bring the most recent browser window to the foreground (Windows). Returns True on success."""
    if not IS_WINDOWS:
        return False
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    names = ("Google Chrome", "Edge", "Firefox", "Brave", "Opera", "Vivaldi")
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def _cb(hwnd, _lparam):
        if user32.IsWindowVisible(hwnd):
            n = user32.GetWindowTextLengthW(hwnd)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(hwnd, buf, n + 1)
                if any(t in buf.value for t in names):
                    found.append(hwnd)
                    return False
        return True

    user32.EnumWindows(_cb, 0)
    if not found:
        return False
    hwnd = found[0]
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    user32.keybd_event(0x12, 0, 0, 0)  # tapping Alt lets SetForegroundWindow succeed
    user32.keybd_event(0x12, 0, 2, 0)
    user32.SetForegroundWindow(hwnd)
    return True


def send_keys(*keys):
    """Press a key combination (Windows) with no extra dependencies."""
    import ctypes

    user32 = ctypes.windll.user32
    for k in keys:
        user32.keybd_event(_VK[k], 0, 1 if k in _EXTENDED else 0, 0)
    for k in reversed(keys):
        user32.keybd_event(_VK[k], 0, (1 if k in _EXTENDED else 0) | 2, 0)

# --------------------------------------------------------------------------- #
# Speech output / helpers
# --------------------------------------------------------------------------- #
_speak_lock = threading.Lock()
_tts_engine = None
SPEAK_ENABLED = False


def init_tts():
    global _tts_engine
    if pyttsx3 is None:
        return False
    try:
        _tts_engine = pyttsx3.init()
        _tts_engine.setProperty("rate", 175)
        return True
    except Exception:
        _tts_engine = None
        return False


def say(text, speak=None):
    """Print the assistant's reply, and speak it if voice replies are on."""
    print(f"\nAssistant: {text}\n")
    if (SPEAK_ENABLED if speak is None else speak) and _tts_engine is not None:
        with _speak_lock:
            try:
                _tts_engine.say(text)
                _tts_engine.runAndWait()
            except Exception:
                pass


def beep_pattern(times=6):
    try:
        if IS_WINDOWS:
            import winsound
            for _ in range(times):
                winsound.Beep(1000, 400)
                time.sleep(0.15)
        else:
            for _ in range(times):
                print("\a", end="", flush=True)
                time.sleep(0.5)
    except Exception:
        pass


def notify(title, message):
    """Best-effort desktop toast, then a console message."""
    # Alarm labels come from the user's own words, so never paste them into a script unescaped.
    title = re.sub(r"[^\w .,:!?\-]", "", str(title))[:60]
    message = re.sub(r"[^\w .,:!?\-]", "", str(message))[:200]
    try:
        if IS_WINDOWS:
            ps = (
                "[reflection.assembly]::loadwithpartialname('System.Windows.Forms') | Out-Null;"
                "$n = New-Object System.Windows.Forms.NotifyIcon;"
                "$n.Icon = [System.Drawing.SystemIcons]::Information;"
                "$n.Visible = $true;"
                f"$n.ShowBalloonTip(10000, '{title}', '{message}', 'Info');"
                "Start-Sleep -Seconds 6; $n.Dispose()"
            )
            subprocess.Popen(
                ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                creationflags=0x08000000,
            )
        elif IS_MAC:
            subprocess.Popen(
                ["osascript", "-e", f'display notification "{message}" with title "{title}"']
            )
        elif shutil.which("notify-send"):
            subprocess.Popen(["notify-send", title, message])
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Date / time parsing
# --------------------------------------------------------------------------- #
WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def parse_duration(text):
    """'1 hour 30 minutes 10 sec' -> seconds (None if nothing found)."""
    total, found = 0, False
    for num, unit in re.findall(
        r"(\d+(?:\.\d+)?)\s*(hours?|hrs?|h|minutes?|mins?|m|seconds?|secs?|s)\b", text.lower()
    ):
        n = float(num)
        u = unit[0]
        if u == "h":
            total += n * 3600
        elif u == "m":
            total += n * 60
        else:
            total += n
        found = True
    # "half an hour"
    if "half an hour" in text.lower():
        total += 1800
        found = True
    return int(total) if found else None


def parse_time_of_day(text):
    """Return (hour, minute) or None. Supports '6:30 pm', '6pm', '18:45', 'noon', 'midnight'."""
    t = text.lower()
    if "noon" in t:
        return 12, 0
    if "midnight" in t:
        return 0, 0
    m = re.search(r"\b(\d{1,2})(?::(\d{2}))?\s*([ap])\.?\s*m\b\.?", t)
    if m:
        h, mi, ap = int(m.group(1)), int(m.group(2) or 0), m.group(3)
        if h == 12:
            h = 0
        if ap == "p":
            h += 12
        if 0 <= h < 24 and 0 <= mi < 60:
            return h, mi
    m = re.search(r"\b([01]?\d|2[0-3]):([0-5]\d)\b", t)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def parse_date(text, now=None):
    """Return a date or None. Supports today/tomorrow/weekday/ISO/dd-mm-yyyy/'15 oct'/'in 3 days'."""
    now = now or datetime.now()
    t = text.lower()
    today = now.date()

    if "day after tomorrow" in t:
        return today + timedelta(days=2)
    if "tomorrow" in t:
        return today + timedelta(days=1)
    if "today" in t or "tonight" in t:
        return today

    m = re.search(r"\bin\s+(\d+)\s+days?\b", t)
    if m:
        return today + timedelta(days=int(m.group(1)))
    m = re.search(r"\bin\s+(\d+)\s+weeks?\b", t)
    if m:
        return today + timedelta(weeks=int(m.group(1)))

    m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", t)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
        except ValueError:
            pass

    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b", t)  # dd/mm/yyyy (India style)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).date()
        except ValueError:
            pass

    month_names = "|".join(MONTHS)
    m = re.search(rf"\b(\d{{1,2}})(?:st|nd|rd|th)?\s+(?:of\s+)?({month_names})[a-z]*(?:,?\s+(\d{{4}}))?", t)
    if not m:
        m2 = re.search(rf"\b({month_names})[a-z]*\s+(\d{{1,2}})(?:st|nd|rd|th)?(?:,?\s+(\d{{4}}))?", t)
        if m2:
            day, mon, yr = int(m2.group(2)), MONTHS[m2.group(1)[:3]], m2.group(3)
        else:
            day = None
    else:
        day, mon, yr = int(m.group(1)), MONTHS[m.group(2)[:3]], m.group(3)
    if day:
        year = int(yr) if yr else today.year
        try:
            d = datetime(year, mon, day).date()
            if not yr and d < today:
                d = datetime(year + 1, mon, day).date()
            return d
        except ValueError:
            pass

    for i, wd in enumerate(WEEKDAYS):
        if re.search(rf"\b{wd}\b", t):
            delta = (i - today.weekday()) % 7
            if delta == 0 or "next" in t:
                delta = delta or 7
            return today + timedelta(days=delta)
    return None


def fmt_dt(dt):
    return dt.strftime("%a %d %b %Y, %I:%M %p")


# --------------------------------------------------------------------------- #
# Alarm / timer manager (background thread)
# --------------------------------------------------------------------------- #
class AlarmManager:
    def __init__(self):
        self.items = []  # dicts: id, kind, label, when (iso)
        self._next_id = 1
        self._lock = threading.Lock()
        self._load()
        threading.Thread(target=self._loop, daemon=True).start()

    # persistence
    def _load(self):
        try:
            data = json.loads(ALARM_FILE.read_text())
            now = datetime.now()
            self.items = [i for i in data if datetime.fromisoformat(i["when"]) > now]
            self._next_id = max([i["id"] for i in self.items] + [0]) + 1
        except Exception:
            self.items = []

    def _save(self):
        try:
            ALARM_FILE.write_text(json.dumps(self.items, indent=2))
        except Exception:
            pass

    def add(self, kind, when, label=""):
        with self._lock:
            item = {"id": self._next_id, "kind": kind, "label": label, "when": when.isoformat()}
            self._next_id += 1
            self.items.append(item)
            self._save()
        return item

    def remove(self, ident=None):
        with self._lock:
            before = len(self.items)
            if ident is None:
                self.items = []
            else:
                self.items = [i for i in self.items if i["id"] != ident]
            self._save()
            return before - len(self.items)

    def list(self):
        with self._lock:
            return sorted(self.items, key=lambda i: i["when"])

    def _loop(self):
        while True:
            time.sleep(1)
            now = datetime.now()
            due = []
            with self._lock:
                for i in self.items:
                    if datetime.fromisoformat(i["when"]) <= now:
                        due.append(i)
                if due:
                    self.items = [i for i in self.items if i not in due]
                    self._save()
            for i in due:
                self._ring(i)

    def _ring(self, item):
        what = "Alarm" if item["kind"] == "alarm" else "Timer"
        label = f" - {item['label']}" if item["label"] else ""
        msg = f"{what} finished{label}!"
        notify(f"{what}", msg)
        print(f"\n\a*** {msg} ***")
        beep_pattern()
        say(msg)
        print("You: ", end="", flush=True)


ALARMS = AlarmManager()


# --------------------------------------------------------------------------- #
# Skills
# --------------------------------------------------------------------------- #
class Skills:
    # ------------------------- web search ------------------------- #
    @staticmethod
    def web_search(query):
        if not query.strip():
            return "What would you like me to search for?"
        url = "https://www.google.com/search?q=" + quote_plus(query)
        if DRY_RUN:
            return f"[dry-run] would search the web for '{query}'"
        webbrowser.open(url)  # opens in the system's default browser
        return f"Searching the web for '{query}' in your default browser."

    @staticmethod
    def open_website(target):
        target = target.strip()
        if not re.match(r"^https?://", target):
            target = "https://" + target
        if DRY_RUN:
            return f"[dry-run] would open {target} in the default browser"
        webbrowser.open(target)
        return f"Opening {target}"

    @staticmethod
    def browser_key(keys, label):
        """Voice-control the browser: focus it, then press a shortcut (Windows)."""
        if DRY_RUN:
            return f"[dry-run] would focus the browser and press {'+'.join(keys).upper()}  ({label})"
        if not IS_WINDOWS:
            return "Browser keyboard control is only available on Windows."
        if not focus_browser():
            return "I couldn't find an open browser window. Say 'open chrome' first."
        time.sleep(0.25)
        send_keys(*keys)
        return label + "."

    # ------------------------- apps ------------------------------- #
    @staticmethod
    def open_app(name):
        key = name.lower().strip()
        if DRY_RUN:
            return f"[dry-run] would open the app '{name}'"
        target = APPS.get(key)
        if target is None:
            # fuzzy: partial match
            for k, v in APPS.items():
                if (len(key) >= 3 and key in k) or (len(k) >= 4 and re.search(rf"\b{re.escape(k)}\b", key)):
                    target = v
                    break
        if target is None:
            # Unknown app: only accept plain program names, never anything that could carry
            # shell syntax (& | ; > < ` $ quotes ...).  This is what stops "open x & del y".
            if not re.fullmatch(r"[\w .\-]{1,60}", key):
                log.warning("refused to launch suspicious app name: %r", name)
                return f"I won't launch '{name}' - app names can only contain letters, numbers, spaces, dots and dashes."
            target = key
        try:
            if IS_WINDOWS:
                if target.endswith(":"):
                    os.startfile(target)  # URI protocol such as ms-settings:
                else:
                    # list form, no shell: the name is passed as ONE argument
                    subprocess.Popen(["cmd", "/c", "start", "", target], creationflags=0x08000000)
            elif IS_MAC:
                subprocess.Popen(["open", "-a", name])
            else:
                exe = shutil.which(target) or shutil.which(key)
                if not exe:
                    return f"I couldn't find an app called '{name}'."
                subprocess.Popen([exe])
            return f"Opening {name}."
        except Exception as e:
            return f"Couldn't open {name}: {e}"

    # ------------------------- files ------------------------------ #
    @staticmethod
    def resolve(path_text, base=None):
        """Turn 'report.txt', 'desktop/report.txt' or a full path into a Path."""
        p = path_text.strip().strip('"').strip("'")
        low = p.lower()
        for name, real in PLACES.items():
            if low == name:
                return real
            if low.startswith(name + "/") or low.startswith(name + "\\"):
                return real / p[len(name) + 1:]
        pp = Path(os.path.expanduser(p))
        if not pp.is_absolute():
            pp = (base or Skills.cwd) / pp
        return pp

    cwd = HOME

    @staticmethod
    def split_location(text):
        """'notes.txt in documents' -> ('notes.txt', Path(Documents))."""
        m = re.search(r"^(.*?)\s+(?:in|on|inside|at)\s+(?:the\s+)?(.+)$", text.strip(), re.I)
        if m:
            loc = Skills.resolve(m.group(2))
            if loc.exists() and loc.is_dir():
                return m.group(1).strip(), loc
        return text.strip(), None

    @staticmethod
    def _p(text):
        name, base = Skills.split_location(text)
        return Skills.resolve(name, base)

    @staticmethod
    def create_folder(text):
        p = Skills._p(text)
        p.mkdir(parents=True, exist_ok=True)
        return f"Folder created: {p}"

    @staticmethod
    def create_file(text):
        p = Skills._p(text)
        p.parent.mkdir(parents=True, exist_ok=True)
        if p.exists():
            return f"That file already exists: {p}"
        p.touch()
        return f"File created: {p}"

    @staticmethod
    def is_protected(p):
        """Paths the assistant must never delete/move/rename, however politely asked."""
        try:
            rp = Path(p).resolve()
        except Exception:
            return True
        if rp == Path(rp.anchor) or rp == HOME.resolve() or rp == DATA_DIR.resolve():
            return True                                   # drive root, home folder, our own data
        if any(rp == v.resolve() for v in PLACES.values()):
            return True                                   # Desktop, Documents, Downloads ... themselves
        for var in ("SystemRoot", "ProgramFiles", "ProgramFiles(x86)", "ProgramData"):
            base = os.environ.get(var)
            if base and (rp == Path(base).resolve() or Path(base).resolve() in rp.parents):
                return True
        if not IS_WINDOWS and any(rp == Path(x) or Path(x) in rp.parents for x in ("/bin", "/etc", "/usr", "/boot", "/sys", "/proc", "/dev", "/lib", "/sbin", "/System", "/Library")):
            return True
        return False

    @staticmethod
    def delete(text, confirm_fn):
        p = Skills._p(text)
        if not p.exists():
            return f"I can't find {p}"
        if Skills.is_protected(p):
            log.warning("blocked delete of protected path %s", p)
            return f"I won't delete '{p}' - it's a protected system or personal folder."
        to_bin = bool(CONFIG.get("use_recycle_bin")) and send2trash is not None
        how = "move to the Recycle Bin" if to_bin else "PERMANENTLY delete (cannot be undone)"
        if not confirm_fn(f"{how.capitalize()}: '{p}'? (yes/no): "):
            return "Cancelled. Nothing was deleted."
        try:
            if to_bin:
                send2trash(str(p))
            elif p.is_dir():
                shutil.rmtree(p)
            else:
                p.unlink()
            log.info("deleted %s (%s)", p, "bin" if to_bin else "permanent")
            return f"{'Moved to Recycle Bin' if to_bin else 'Deleted'}: {p}"
        except Exception as e:
            return f"Couldn't delete: {e}"

    @staticmethod
    def rename(src, new_name):
        s = Skills._p(src)
        if not s.exists():
            return f"I can't find {s}"
        if Skills.is_protected(s):
            return f"I won't rename '{s}' - it's protected."
        new_name = new_name.strip().strip('"')
        if not new_name or any(c in new_name for c in '\\/:*?"<>|'):
            return "That isn't a valid file name (no \\ / : * ? \" < > |)."
        d = s.with_name(new_name)
        if d.exists():
            return f"'{d.name}' already exists - I won't overwrite it."
        s.rename(d)
        return f"Renamed to {d.name}"

    @staticmethod
    def copy(src, dst):
        s, d = Skills._p(src), Skills._p(dst)
        if not s.exists():
            return f"I can't find {s}"
        target = d / s.name if d.is_dir() else d
        if target.exists():
            return f"'{target}' already exists - I won't overwrite it."
        if s.is_dir():
            shutil.copytree(s, target)
        else:
            shutil.copy2(s, target)
        return f"Copied {s.name} to {target.parent}"

    @staticmethod
    def move(src, dst):
        s, d = Skills._p(src), Skills._p(dst)
        if not s.exists():
            return f"I can't find {s}"
        if Skills.is_protected(s):
            return f"I won't move '{s}' - it's protected."
        target = d / s.name if d.is_dir() else d
        if target.exists():
            return f"'{target}' already exists - I won't overwrite it."
        shutil.move(str(s), str(target))
        return f"Moved {s.name} to {target.parent}"

    @staticmethod
    def list_dir(text=""):
        p = Skills._p(text) if text.strip() else Skills.cwd
        if not p.is_dir():
            return f"That's not a folder: {p}"
        entries = sorted(p.iterdir(), key=lambda x: (not x.is_dir(), x.name.lower()))
        if not entries:
            return f"{p} is empty."
        lines = [f"Contents of {p}:"]
        for e in entries[:60]:
            lines.append(f"  [DIR]  {e.name}" if e.is_dir() else f"  {e.name}  ({e.stat().st_size:,} bytes)")
        if len(entries) > 60:
            lines.append(f"  ... and {len(entries) - 60} more")
        return "\n".join(lines)

    @staticmethod
    def find(name, base_text=None):
        base = Skills.resolve(base_text) if base_text else HOME
        hits = []
        needle = name.lower().strip().strip('"')
        deadline = time.monotonic() + 8          # never hang the assistant on a huge disk
        timed_out = False
        for root, dirs, files in os.walk(base):
            if time.monotonic() > deadline:
                timed_out = True
                break
            dirs[:] = [d for d in dirs if not d.startswith(".") and d not in ("AppData", "node_modules", "$Recycle.Bin")]
            for f in files + dirs:
                if needle in f.lower():
                    hits.append(os.path.join(root, f))
                    if len(hits) >= 15:
                        break
            if len(hits) >= 15:
                break
        note = "\n  (search stopped after 8 s - narrow it, e.g. 'find file invoice in documents')" if timed_out else ""
        if not hits:
            return f"No files matching '{name}' found under {base}." + note
        return "Found:\n" + "\n".join("  " + h for h in hits) + note

    @staticmethod
    def read_file(text):
        p = Skills._p(text)
        if not p.is_file():
            return f"I can't find the file {p}"
        try:
            data = p.read_text(errors="replace")
        except Exception as e:
            return f"Couldn't read: {e}"
        return data[:1500] + ("\n... (truncated)" if len(data) > 1500 else "") or "(empty file)"

    @staticmethod
    def open_path(text):
        """Open a file with its default app, or a folder in File Explorer."""
        p = Skills._p(text)
        if not p.exists():
            return f"I can't find {p}"
        try:
            if IS_WINDOWS:
                os.startfile(str(p))
            elif IS_MAC:
                subprocess.Popen(["open", str(p)])
            else:
                subprocess.Popen(["xdg-open", str(p)])
            return f"Opened {p}"
        except Exception as e:
            return f"Couldn't open: {e}"

    @staticmethod
    def write_file(text, content):
        p = Skills._p(text)
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "a", encoding="utf-8") as f:
            f.write(content + "\n")
        return f"Wrote to {p}"

    # ------------------------- alarm / timer ---------------------- #
    @staticmethod
    def set_alarm(text):
        tod = parse_time_of_day(text)
        if tod is None:
            return "Please give a time, e.g. 'set alarm for 6:30 am'."
        now = datetime.now()
        d = parse_date(text, now)
        when = datetime.combine(d or now.date(), datetime.min.time()).replace(hour=tod[0], minute=tod[1])
        if when <= now:
            if d is None:
                when += timedelta(days=1)
            else:
                return "That time is already in the past."
        label = ""
        m = re.search(r"(?:called|named|labell?ed|for the|to)\s+(.+)$", text, re.I)
        if m and not parse_time_of_day(m.group(1)):
            label = m.group(1).strip()
        item = ALARMS.add("alarm", when, label)
        Skills.open_app("clock")  # also open the Windows Clock app
        return (f"Alarm #{item['id']} set for {fmt_dt(when)}. "
                "Keep this assistant running so it can ring. (I also opened the Clock app.)")

    @staticmethod
    def set_timer(text):
        secs = parse_duration(text)
        if not secs:
            return "How long? e.g. 'set timer for 10 minutes'."
        when = datetime.now() + timedelta(seconds=secs)
        item = ALARMS.add("timer", when)
        h, rem = divmod(secs, 3600)
        mi, s = divmod(rem, 60)
        parts = [f"{h}h" if h else "", f"{mi}m" if mi else "", f"{s}s" if s else ""]
        return f"Timer #{item['id']} started for {' '.join(p for p in parts if p)} (rings at {when:%I:%M:%S %p})."

    @staticmethod
    def show_alarms():
        items = ALARMS.list()
        if not items:
            return "You have no active alarms or timers."
        lines = ["Active alarms/timers:"]
        for i in items:
            when = datetime.fromisoformat(i["when"])
            left = when - datetime.now()
            left_s = str(left).split(".")[0]
            lines.append(f"  #{i['id']} {i['kind']:<5} {fmt_dt(when)}  (in {left_s}) {i['label']}")
        return "\n".join(lines)

    @staticmethod
    def cancel_alarm(text):
        m = re.search(r"#?(\d+)", text)
        if "all" in text.lower():
            n = ALARMS.remove()
            return f"Cancelled {n} alarm(s)/timer(s)."
        if m:
            n = ALARMS.remove(int(m.group(1)))
            return "Cancelled." if n else "No alarm with that number."
        return "Say 'cancel alarm 2' or 'cancel all alarms'."

    # ------------------------- calendar --------------------------- #
    @staticmethod
    def add_event(text):
        now = datetime.now()
        # title = text after the keyword, before the first date/time word
        m = re.search(
            r"(?:event|meeting|appointment|reminder|birthday|schedule|mark)\s*(?:called|named|titled|for|of)?\s*(.*)",
            text, re.I,
        )
        rest = m.group(1) if m else text
        title = re.split(
            r"\s+(?:on|at|tomorrow|today|tonight|next|from|in|for\s+\d)\b", rest, maxsplit=1, flags=re.I
        )[0].strip(" ,.")
        if not title:
            title = "New event"
        title = title[:1].upper() + title[1:]

        d = parse_date(text, now) or now.date()
        tod = parse_time_of_day(text)
        all_day = tod is None
        dur = 3600
        m2 = re.search(r"for\s+(\d.*?)(?:$|\s+(?:on|at|tomorrow|today))", text, re.I)
        if m2 and parse_duration(m2.group(1)):
            dur = parse_duration(m2.group(1))

        if all_day:
            start = datetime.combine(d, datetime.min.time())
            end = start + timedelta(days=1)
        else:
            start = datetime.combine(d, datetime.min.time()).replace(hour=tod[0], minute=tod[1])
            end = start + timedelta(seconds=dur)

        def esc(s):
            return s.replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;")

        uid = f"{int(time.time())}@ai-desktop-assistant"
        stamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
        if all_day:
            dts = f"DTSTART;VALUE=DATE:{start:%Y%m%d}\nDTEND;VALUE=DATE:{end:%Y%m%d}"
        else:
            dts = f"DTSTART:{start:%Y%m%dT%H%M%S}\nDTEND:{end:%Y%m%dT%H%M%S}"
        ics = (
            "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//AI Desktop Assistant//EN\n"
            "BEGIN:VEVENT\n"
            f"UID:{uid}\nDTSTAMP:{stamp}\nSUMMARY:{esc(title)}\n{dts}\n"
            "BEGIN:VALARM\nTRIGGER:-PT15M\nACTION:DISPLAY\nDESCRIPTION:Reminder\nEND:VALARM\n"
            "END:VEVENT\nEND:VCALENDAR\n"
        ).replace("\n", "\r\n")
        safe = re.sub(r"[^\w\-]+", "_", title)[:40]
        path = EVENT_DIR / f"{safe}_{start:%Y%m%d_%H%M}.ics"
        path.write_text(ics, encoding="utf-8")
        try:
            if IS_WINDOWS:
                os.startfile(str(path))  # opens in default calendar app -> click Save
            elif IS_MAC:
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            pass
        when_txt = f"{start:%a %d %b %Y}" if all_day else fmt_dt(start)
        return (f"Event '{title}' prepared for {when_txt}. I opened it in your calendar app - "
                f"press Save/Add there. (File: {path})")

    # ------------------------- system status ---------------------- #
    @staticmethod
    def system_status(detail="all"):
        if psutil is None:
            return "Please install psutil first:  pip install psutil"
        d = detail.lower()
        out, warnings = [], []

        cpu = psutil.cpu_percent(interval=1)
        vm = psutil.virtual_memory()
        sw = psutil.swap_memory()
        boot = datetime.fromtimestamp(psutil.boot_time())
        up = datetime.now() - boot

        if d in ("all", "cpu"):
            freq = psutil.cpu_freq()
            out.append(f"CPU: {cpu:.0f}% used | {psutil.cpu_count(logical=False)} cores / "
                       f"{psutil.cpu_count()} threads" + (f" | {freq.current:.0f} MHz" if freq else ""))
        if cpu > 85:
            warnings.append("CPU usage is very high - check the top processes.")

        if d in ("all", "memory", "ram"):
            out.append(f"Memory: {vm.percent:.0f}% used ({vm.used / 2**30:.1f} of {vm.total / 2**30:.1f} GB) "
                       f"| Swap {sw.percent:.0f}%")
        if vm.percent > 85:
            warnings.append("Memory is almost full - close unused apps.")

        if d in ("all", "disk", "storage"):
            for part in psutil.disk_partitions(all=False):
                try:
                    u = psutil.disk_usage(part.mountpoint)
                except Exception:
                    continue
                out.append(f"Disk {part.device}: {u.percent:.0f}% used "
                           f"({u.used / 2**30:.0f} of {u.total / 2**30:.0f} GB, {u.free / 2**30:.0f} GB free)")
                if u.percent > 90:
                    warnings.append(f"Drive {part.device} is nearly full.")

        if d in ("all", "battery"):
            bat = psutil.sensors_battery() if hasattr(psutil, "sensors_battery") else None
            if bat:
                state = "charging" if bat.power_plugged else "on battery"
                left = f", ~{int(bat.secsleft // 60)} min left" if bat.secsleft and bat.secsleft > 0 and not bat.power_plugged else ""
                out.append(f"Battery: {bat.percent:.0f}% ({state}{left})")
                if bat.percent < 20 and not bat.power_plugged:
                    warnings.append("Battery is low - plug in your charger.")
            elif d == "battery":
                out.append("No battery detected (desktop PC).")

        if d in ("all", "network"):
            n = psutil.net_io_counters()
            out.append(f"Network: sent {n.bytes_sent / 2**20:.0f} MB, received {n.bytes_recv / 2**20:.0f} MB since boot")

        if d == "all":
            out.append(f"OS: {platform.system()} {platform.release()} ({platform.version()})")
            out.append(f"Uptime: {str(up).split('.')[0]} (since {boot:%d %b %H:%M})")
            if up.days >= 7:
                warnings.append("Your PC hasn't restarted in over a week - consider restarting.")

        if d in ("all", "processes", "process", "apps"):
            procs = []
            for p in psutil.process_iter(["name", "memory_percent"]):
                try:
                    procs.append((p.info["memory_percent"] or 0, p.info["name"]))
                except Exception:
                    pass
            procs.sort(reverse=True)
            out.append("Top memory users: " + ", ".join(f"{n} ({m:.1f}%)" for m, n in procs[:5]))

        verdict = "Verdict: System looks healthy." if not warnings else "Verdict: Attention needed:\n  - " + "\n  - ".join(warnings)
        return "\n".join(out + [verdict])


# --------------------------------------------------------------------------- #
# Natural-language command router
# --------------------------------------------------------------------------- #
HELP_TEXT = """\
I can do the following (just say it naturally):

  WEB       search python decorators | google weather in delhi | open website github.com
  FILES     create folder Projects in documents | create file notes.txt on desktop
            delete file old.txt | rename a.txt to b.txt | copy a.txt to desktop
            move report.pdf to documents | list files in downloads | find file invoice
            read file notes.txt | open folder downloads | write "hello" to notes.txt
  APPS      open notepad | launch chrome | start calculator | open settings
  BROWSER   go to youtube | new tab | close tab | next tab | go back | refresh
            scroll down | scroll up | scroll to top | zoom in | full screen
  ALARM     set alarm for 6:30 am | set alarm tomorrow at 5 pm | show alarms | cancel alarm 1
  TIMER     set timer for 10 minutes | timer 1 hour 30 minutes
  CALENDAR  add event team meeting tomorrow at 3 pm | mark birthday party on 25 october
  SYSTEM    system status | cpu | memory | disk | battery | network | top processes
  MODES     voice mode | text mode | speak on | speak off | listen (one voice command)
  OTHER     time | date | help | exit
"""


def strip_polite(text):
    t = text.strip()
    t = re.sub(r"^(?:hey |hi |ok |okay )?(?:assistant|jarvis|computer)[, ]+", "", t, flags=re.I)
    t = re.sub(r"^(?:please|can you|could you|would you|i want you to|i want to)\s+", "", t, flags=re.I)
    t = re.sub(r"\s+please$", "", t, flags=re.I)
    return t.strip()


def handle(raw, confirm_fn):
    """Return the reply string for a command, or None to quit."""
    text = strip_polite(raw)
    low = text.lower().strip().rstrip(".!?")
    if not low:
        return "I didn't catch that."

    # --- exit / help / time ---
    if low in ("exit", "quit", "bye", "goodbye", "close assistant", "stop assistant"):
        return None
    if low in ("help", "what can you do", "commands"):
        return HELP_TEXT
    if re.fullmatch(r"(what(?:'s| is) )?(the )?(current )?time( now)?|what time is it", low):
        return f"It's {datetime.now():%I:%M %p}."
    if re.fullmatch(r"(what(?:'s| is) )?(the )?(today'?s? )?date( today)?|what day is it( today)?", low):
        return f"Today is {datetime.now():%A, %d %B %Y}."

    # --- browser navigation by voice (acts on your most recent browser window) ---
    for phrases, keys, label in BROWSER_ACTIONS:
        if low in phrases:
            return Skills.browser_key(keys, label)
    m = re.match(r"(?:open|go to|goto|visit|launch|take me to|navigate to)\s+(?:the\s+)?(?:website\s+|site\s+)?(.+)$", low)
    if m and m.group(1).replace(" ", "") in SITES:
        return Skills.open_website(SITES[m.group(1).replace(" ", "")])

    # --- alarms & timers (check before generic 'set') ---
    if re.search(r"\b(show|list|view|what)\b.*\b(alarms?|timers?)\b", low) or low in ("alarms", "timers"):
        return Skills.show_alarms()
    if re.search(r"\b(cancel|delete|remove|stop|clear)\b.*\b(alarms?|timers?)\b", low):
        return Skills.cancel_alarm(low)
    if re.search(r"\btimer\b", low) or re.search(r"\bcountdown\b", low):
        return Skills.set_timer(low)
    if re.search(r"\b(alarm|wake me)\b", low):
        return Skills.set_alarm(low)

    # --- calendar ---
    if re.search(r"\b(event|meeting|appointment|calendar|birthday|schedule)\b", low) and \
            re.search(r"\b(add|create|mark|set|schedule|put|new|book)\b", low):
        return Skills.add_event(text)
    if low in ("open calendar", "show calendar"):
        return Skills.open_app("calendar")

    # --- system status ---
    if re.search(r"\b(system|pc|computer|windows|laptop)\b.*\b(status|health|performance|info|analy[sz]e|analysis|report)\b", low) \
            or low in ("status", "system status", "analyze system", "analyse system", "analyze windows", "diagnostics"):
        return Skills.system_status("all")
    for kw, d in (("cpu", "cpu"), ("processor", "cpu"), ("memory", "memory"), ("ram", "memory"),
                  ("disk", "disk"), ("storage", "disk"), ("battery", "battery"),
                  ("network", "network"), ("internet usage", "network"),
                  ("top processes", "processes"), ("running processes", "processes"), ("running apps", "processes")):
        if re.search(rf"\b{kw}\b", low) and not re.search(r"\b(open|launch|start|file|folder)\b", low):
            return Skills.system_status(d)

    # --- file operations ---
    m = re.match(r"(?:create|make|new)\s+(?:a\s+|an\s+)?(?:new\s+)?(?:folder|directory)\s+(?:called\s+|named\s+)?(.+)", text, re.I)
    if m:
        return Skills.create_folder(m.group(1))
    m = re.match(r"(?:create|make|new)\s+(?:a\s+|an\s+)?(?:new\s+)?file\s+(?:called\s+|named\s+)?(.+)", text, re.I)
    if m:
        return Skills.create_file(m.group(1))
    m = re.match(r"(?:delete|remove)\s+(?:the\s+)?(?:file|folder|directory)\s+(.+)", text, re.I)
    if m:
        return Skills.delete(m.group(1), confirm_fn)
    m = re.match(r"rename\s+(?:the\s+)?(?:file\s+|folder\s+)?(.+?)\s+(?:to|as)\s+(.+)", text, re.I)
    if m:
        return Skills.rename(m.group(1), m.group(2))
    m = re.match(r"copy\s+(?:the\s+)?(?:file\s+|folder\s+)?(.+?)\s+to\s+(.+)", text, re.I)
    if m:
        return Skills.copy(m.group(1), m.group(2))
    m = re.match(r"move\s+(?:the\s+)?(?:file\s+|folder\s+)?(.+?)\s+to\s+(.+)", text, re.I)
    if m:
        return Skills.move(m.group(1), m.group(2))
    m = re.match(r"(?:list|show)\s+(?:all\s+)?(?:the\s+)?(?:files|folders|contents)(?:\s+(?:in|of|inside|from)\s+(.+))?", text, re.I)
    if m:
        return Skills.list_dir(m.group(1) or "")
    m = re.match(r"(?:find|search for|locate)\s+(?:a\s+|the\s+)?(?:file|folder)s?\s+(?:called\s+|named\s+)?(.+?)(?:\s+(?:in|under)\s+(.+))?$", text, re.I)
    if m:
        return Skills.find(m.group(1), m.group(2))
    m = re.match(r"(?:read|show|display|cat)\s+(?:the\s+)?(?:file\s+|contents of\s+)(.+)", text, re.I)
    if m:
        return Skills.read_file(m.group(1))
    m = re.match(r"(?:write|append|add)\s+[\"']?(.+?)[\"']?\s+(?:to|into)\s+(?:file\s+)?(.+)", text, re.I)
    if m and not re.search(r"\b(event|calendar|alarm)\b", low):
        return Skills.write_file(m.group(2), m.group(1))
    m = re.match(r"(?:go to|cd|change directory to|set directory to)\s+(.+)", text, re.I)
    if m:
        p = Skills.resolve(m.group(1))
        if p.is_dir():
            Skills.cwd = p
            return f"Working folder is now {p}"
        return f"That folder doesn't exist: {p}"
    if low in ("where am i", "current folder", "current directory", "pwd"):
        return f"Working folder: {Skills.cwd}"
    m = re.match(r"(?:open|show|reveal)\s+(?:the\s+)?(?:file|folder|directory)\s+(.+)", text, re.I)
    if m:
        return Skills.open_path(m.group(1))
    m = re.match(r"open\s+(?:the\s+)?(desktop|documents|downloads|pictures|music|videos|home)(?:\s+folder)?$", low)
    if m:
        return Skills.open_path(m.group(1))

    # --- websites / search ---
    m = re.match(r"(?:open|go to|visit|launch)\s+(?:the\s+)?(?:website|site|url|page)?\s*((?:https?://)?(?:www\.)?[\w\-]+\.(?:com|org|net|in|io|edu|gov|co|ai|dev|app)(?:/\S*)?)$", text, re.I)
    if m:
        return Skills.open_website(m.group(1))
    m = re.match(r"(?:search|google|look up|lookup|find|browse|bing)\s+(?:for\s+|about\s+|on (?:google|the web|the internet)\s+)?(.+)", text, re.I)
    if m:
        return Skills.web_search(m.group(1))
    m = re.match(r"(?:what is|who is|who was|how to|how do i|why is|where is|tell me about)\s+.+", text, re.I)
    if m:
        return Skills.web_search(text)

    # --- apps ---
    m = re.match(r"(?:open|launch|start|run)\s+(?:the\s+)?(?:app\s+|application\s+|program\s+)?(.+)", text, re.I)
    if m:
        return Skills.open_app(m.group(1))

    return ("Sorry, I didn't understand that. Say 'help' to see what I can do, "
            "or I can search the web for it - just say 'search " + text + "'.")


# --------------------------------------------------------------------------- #
# Robust voice transcription: detect errors -> recover -> adapt
# --------------------------------------------------------------------------- #
#
# DESIGN (accuracy vs. latency)
# -----------------------------
# The normal path is exactly what the assistant always did: calibrate briefly,
# capture once, ask the recognizer once.  If that result is confident (or valid
# for the assistant), it is returned immediately - ZERO extra latency.
#
# Only when an error is DETECTED does the recovery ladder start.  Detection:
#   * recognizer raised "unintelligible" / timed out / engine unreachable
#   * confidence below CONF_MIN
#   * text is not a command the assistant understands AND confidence is not high
#
# The ladder is ordered from cheapest to most expensive, and each rung stops the
# climb as soon as it succeeds:
#   1. n-best alternatives   (0 s   - already returned by the recognizer)
#   2. fuzzy repair          (~0 s   - local text fix: "scroll dub" -> "scroll down")
#   3. second-opinion engine (~0.5 s - same audio, different language model)
#   4. offline engine        (~1 s   - only if the online service is unreachable)
#   5. re-capture with a stronger noise profile (~3-4 s, at most MAX_CAPTURES total)
#   6. cross-attempt consensus, then "did you mean ...?" confirmation
#
# The latency cost is bounded by MAX_CAPTURES and BUDGET_S, and the noise level
# is remembered: after a noisy command the next one starts directly with the
# stronger profile (fast adaptation); after a few clean commands it relaxes back.

PRIMARY_LANG = "en-IN"    # first-choice recognition language
SECONDARY_LANG = "en-US"  # "second opinion" language model (same audio)


class NoSpeechError(Exception):
    """Nothing usable was heard. silence=True means a pure timeout (user never spoke)."""

    def __init__(self, msg, silence=False):
        super().__init__(msg)
        self.silence = silence


class EngineError(Exception):
    """The recognition engine itself failed (network down, engine missing...)."""


class TranscriptionFailed(Exception):
    """Every recovery rung was tried and none produced a trustworthy command."""


@dataclass
class Hypothesis:
    text: str
    conf: float
    engine: str = ""


@dataclass
class AudioProfile:
    name: str
    calibrate: float      # seconds spent measuring background noise
    energy_mult: float    # how far above the noise floor speech must rise
    pause: float          # silence (s) that ends a phrase
    phrase_limit: int     # hard cap on phrase length (s) - stops noise from dragging a capture on


PROFILES = [
    AudioProfile("normal", 0.6, 1.0, 0.8, 12),      # identical to the original behaviour
    AudioProfile("noisy", 1.2, 1.4, 0.6, 8),
    AudioProfile("very-noisy", 1.5, 1.9, 0.5, 6),
]


@dataclass
class TranscriptionResult:
    text: str
    how: str            # direct | n-best | fuzzy-repair | second-engine | offline-engine | consensus | confirmed-guess
    captures: int
    elapsed: float
    profile: str
    trace: List[str]

    @property
    def recovered(self):
        return self.captures > 1 or self.how != "direct"


class Clock:
    def now(self):
        return time.monotonic()


class VirtualClock(Clock):
    """Simulated time, so the demo can report realistic latencies instantly."""

    def __init__(self):
        self.t = 0.0

    def now(self):
        return self.t

    def advance(self, seconds):
        self.t += seconds


# Typical noisy-room mishearings -> what was probably said.
CONFUSIONS = {
    "you tube": "youtube", "u tube": "youtube", "git hub": "github", "g mail": "gmail",
    "note pad": "notepad", "crumb": "chrome", "crome": "chrome", "goat": "go",
    "go two": "go to", "go too": "go to", "tap": "tab", "tub": "tab",
    "lunch": "launch", "opened": "open", "opens": "open", "scrolled": "scroll",
}


class CommandVocabulary:
    """Knows what the assistant can understand: validates transcripts and repairs near-misses."""

    VERBS = set("""
        search google look lookup find browse bing open launch start run go goto visit take navigate
        create make new delete remove rename copy move list show read display cat write append add
        set cancel stop clear view alarm alarms timer timers wake mark schedule book put
        what who whats how why where tell system pc computer windows laptop status analyze analyse
        cpu processor memory ram disk storage battery network top running current pwd cd change
        help exit quit bye goodbye time date voice text speak talk listen type
        commands diagnostics locate reveal hey ok okay please
    """.split())
    OPEN_VERBS = ("open", "launch", "start", "run", "go to", "goto", "visit", "take me to", "navigate to")
    SOLO_OK = {"cpu", "memory", "ram", "disk", "storage", "battery", "network", "status", "help", "exit",
               "quit", "bye", "goodbye", "time", "date", "alarms", "timers", "listen", "pwd", "commands",
               "diagnostics", "processor"}

    def __init__(self):
        self.slots = {}
        for name in list(APPS) + list(SITES) + list(PLACES):
            self.slots[name.replace(" ", "")] = name
        self.phrases = set(BROWSER_PHRASES) | {
            "system status", "top processes", "show alarms", "voice mode", "text mode", "speak on",
            "speak off", "what time is it", "what can you do", "close assistant", "stop assistant",
        }

    @staticmethod
    def norm(text):
        t = strip_polite(text).lower()
        t = re.sub(r"[^\w\s:/.\-'@#]", " ", t)
        return re.sub(r"\s+", " ", t).strip()

    def split_verb(self, t):
        for v in sorted(self.OPEN_VERBS, key=len, reverse=True):
            if t == v or t.startswith(v + " "):
                return v, t[len(v):].strip()
        head, _, rest = t.partition(" ")
        return head, rest

    def slot_ok(self, slot):
        slot = re.sub(r"\s+(folder|app|application|website|site)$", "", re.sub(r"^(the|a|an)\s+", "", slot))
        if slot.replace(" ", "") in self.slots:
            return True
        if re.match(r"^(file|folder|directory|website|site|app|application|program|url|page)\b", slot):
            return True
        return bool(re.search(r"[\w\-]+\.\w{2,5}", slot))  # a domain or a file name

    def is_valid(self, text):
        """Would handle() understand this?  (No side effects.)"""
        t = self.norm(text)
        if not t:
            return False
        if t in self.phrases:
            return True
        verb, rest = self.split_verb(t)
        if verb in self.OPEN_VERBS:
            return bool(rest) and self.slot_ok(rest)
        if verb not in self.VERBS:
            return False
        if verb == "set" and not re.search(r"alarm|timer", rest):
            return False
        return bool(rest) or verb in self.SOLO_OK

    def repair(self, text):
        """Try to turn a near-miss into a valid command. Returns (fixed_text | None, similarity 0..1)."""
        t = self.norm(text)
        sims = []
        for wrong, right in CONFUSIONS.items():
            pat = rf"\b{re.escape(wrong)}\b"
            if re.search(pat, t):
                t = re.sub(pat, right, t)
                sims.append(0.9)
        if self.is_valid(t):
            return t, min(sims or [1.0])
        # whole-phrase fuzzy match (short browser/system commands)
        near = difflib.get_close_matches(t, sorted(self.phrases), n=1, cutoff=0.7)
        if near:
            return near[0], min(sims + [difflib.SequenceMatcher(None, t, near[0]).ratio()])
        # verb + slot fuzzy match
        verb, rest = self.split_verb(t)
        single_open = [v for v in self.OPEN_VERBS if " " not in v]
        if verb not in self.VERBS and verb not in self.OPEN_VERBS:
            cand = difflib.get_close_matches(verb, sorted(self.VERBS | set(single_open)), n=1, cutoff=0.7)
            if not cand:
                return None, 0.0
            sims.append(difflib.SequenceMatcher(None, verb, cand[0]).ratio())
            verb = cand[0]
        if verb in self.OPEN_VERBS and rest and not self.slot_ok(rest):
            cand = difflib.get_close_matches(rest.replace(" ", ""), list(self.slots), n=1, cutoff=0.6)
            if not cand:
                return None, 0.0
            sims.append(difflib.SequenceMatcher(None, rest.replace(" ", ""), cand[0]).ratio())
            rest = self.slots[cand[0]]
        fixed = f"{verb} {rest}".strip()
        return (fixed, min(sims or [1.0])) if self.is_valid(fixed) else (None, 0.0)


class RobustTranscriber:
    CONF_OK = 0.80       # trust a transcript this confident without further checks
    CONF_MIN = 0.45      # minimum confidence for an alternative that is a valid command
    CONF_FLOOR = 0.25    # below this a transcript is treated as noise, never repaired
    REPAIR_MIN = 0.75    # minimum text similarity for fuzzy repair to be accepted
    MAX_CAPTURES = 3     # hard bound on re-recording (the expensive step)
    BUDGET_S = 20.0      # no new capture is started after this many seconds
    RELAX_AFTER = 2      # clean first-try commands needed to step the noise profile back down

    def __init__(self, backend, vocab, notify=print, confirm=None, clock=None):
        self.b, self.v = backend, vocab
        self.notify = notify
        self.confirm = confirm
        self.clock = clock or Clock()
        self.level = 0           # index into PROFILES (remembered between commands)
        self.clean_streak = 0
        self.trace: List[str] = []
        self._cands: Dict[str, Tuple[str, float]] = {}
        self._engine_err = ""

    # ---- small helpers ---------------------------------------------------- #
    def _log(self, msg):
        self.trace.append(msg)

    def _tell(self, msg):
        self.trace.append("» " + msg)
        self.notify(msg)

    def _escalate(self, why):
        if self.level < len(PROFILES) - 1:
            self.level += 1
            self._tell(f"{why} - switching to the '{PROFILES[self.level].name}' audio profile and retrying...")
        else:
            self._tell(f"{why} - retrying...")

    @staticmethod
    def _fmt(hyps):
        return "; ".join(f"'{h.text}' {h.conf:.2f}" for h in hyps[:3])

    # ---- selection: pick a command out of recognizer hypotheses ------------ #
    def _select(self, hyps):
        if not hyps:
            return None
        top = hyps[0]
        if top.conf >= self.CONF_OK:                      # confident -> trust, no extra work
            if self.v.is_valid(top.text):
                return top.text, "direct"
            fixed, sim = self.v.repair(top.text)
            if fixed and sim >= 0.85:                     # only fix a confident transcript if the fix is very close
                return fixed, "fuzzy-repair"
            return top.text, "direct"
        for h in hyps:                                    # rung 1: n-best alternatives
            if h.conf >= self.CONF_MIN and self.v.is_valid(h.text):
                return h.text, ("direct" if h is top else "n-best")
        for h in hyps:                                    # rung 2: fuzzy repair
            if h.conf >= self.CONF_FLOOR:
                fixed, sim = self.v.repair(h.text)
                if fixed and sim >= self.REPAIR_MIN and fixed != self.v.norm(h.text):
                    return fixed, "fuzzy-repair"
        return None

    def _collect_candidates(self, pool):
        """Remember plausible-but-unconfirmed candidates so a later capture can vote on them."""
        cands = {}
        for h in pool:
            if h.conf < self.CONF_FLOOR:
                continue
            text, sim = (h.text, 1.0) if self.v.is_valid(h.text) else self.v.repair(h.text)
            if text and sim >= self.REPAIR_MIN:
                key = self.v.norm(text)
                if key not in cands or h.conf > cands[key][1]:
                    cands[key] = (text, h.conf)
        self._cands = cands

    # ---- one captured clip: rungs 1-4 -------------------------------------- #
    def _attempt(self, audio):
        pool: List[Hypothesis] = []
        primary_down = False
        for label, eng in (("primary", self.b.primary), ("second-opinion", self.b.secondary)):
            if eng is None:
                continue
            try:
                hyps = eng(audio)
            except NoSpeechError:
                self._log(f"  {label}: nothing intelligible")
                continue
            except EngineError as e:
                self._engine_err = str(e)
                self._log(f"  {label}: engine unreachable ({e})")
                primary_down = primary_down or label == "primary"
                break                                     # same service would fail again -> go offline
            pool += hyps
            self._log(f"  {label}: {self._fmt(hyps)}")
            sel = self._select(hyps if label == "primary" else pool)
            if sel:
                text, how = sel
                if label == "second-opinion" and how in ("direct", "n-best"):
                    how = "second-engine"
                return text, how
        if primary_down and self.b.offline is not None:
            try:
                hyps = self.b.offline(audio)
                pool += hyps
                self._log(f"  offline: {self._fmt(hyps)}")
                sel = self._select(hyps)
                if sel:
                    return sel[0], "offline-engine"
            except (NoSpeechError, EngineError) as e:
                self._log(f"  offline: failed ({e})")
        self._collect_candidates(pool)
        return None

    # ---- public entry point ------------------------------------------------- #
    def transcribe(self):
        self.trace = []
        t0 = self.clock.now()
        votes: Dict[str, int] = defaultdict(int)
        score: Dict[str, float] = defaultdict(float)
        best_text: Dict[str, str] = {}
        last_err = "Sorry, I couldn't understand the audio."
        captures = 0

        while captures < self.MAX_CAPTURES:
            if captures and self.clock.now() - t0 > self.BUDGET_S:
                self._log("latency budget reached - stopping recovery")
                break
            profile = PROFILES[self.level]
            captures += 1
            self._cands, self._engine_err = {}, ""
            try:
                audio = self.b.capture(profile)
            except NoSpeechError as e:
                self._log(f"capture {captures} [{profile.name}]: {e}")
                last_err = str(e)
                if captures < self.MAX_CAPTURES:
                    if e.silence:
                        self._tell("I didn't hear you - say it again...")
                    else:
                        self._escalate("Too much background noise")
                continue
            self._log(f"capture {captures} [{profile.name}]")

            res = self._attempt(audio)
            if res:
                return self._finish(res[0], res[1], captures, t0)

            for key, (text, conf) in self._cands.items():       # rung 6a: cross-attempt consensus
                votes[key] += 1
                score[key] += conf
                best_text[key] = text
                if votes[key] >= 2:
                    return self._finish(text, "consensus", captures, t0)
            if self._engine_err:
                last_err = f"Speech service error (internet needed): {self._engine_err}"
            else:
                last_err = "Sorry, I couldn't understand the audio over the background noise."
            if captures < self.MAX_CAPTURES:
                self._escalate("I couldn't make that out")

        if score and self.confirm is not None:                   # rung 6b: ask the user
            key = max(score, key=score.get)
            if self.confirm(f"Did you mean '{best_text[key]}'? (yes/no): "):
                return self._finish(best_text[key], "confirmed-guess", captures, t0)
        raise TranscriptionFailed(last_err + " You can type the command instead.")

    def _finish(self, text, how, captures, t0):
        if captures == 1 and how == "direct":
            self.clean_streak += 1
            if self.clean_streak >= self.RELAX_AFTER and self.level > 0:
                self.level -= 1
                self.clean_streak = 0
                self._log(f"quiet again - relaxing to the '{PROFILES[self.level].name}' profile")
        else:
            self.clean_streak = 0
        return TranscriptionResult(text, how, captures, self.clock.now() - t0,
                                   PROFILES[self.level].name, list(self.trace))


# --------------------------------------------------------------------------- #
# Real microphone + engines
# --------------------------------------------------------------------------- #
class MicBackend:
    def __init__(self):
        self.rec = sr.Recognizer()
        self.rec.dynamic_energy_threshold = True
        self.primary = self._primary
        self.secondary = lambda audio: self._google(audio, SECONDARY_LANG)
        try:
            import pocketsphinx  # noqa: F401  (optional offline engine: pip install pocketsphinx)
            self.offline = self._sphinx
        except Exception:
            self.offline = None

    def capture(self, p):
        try:
            with sr.Microphone() as mic:
                self.rec.pause_threshold = p.pause
                self.rec.adjust_for_ambient_noise(mic, duration=p.calibrate)
                self.rec.energy_threshold *= p.energy_mult
                tag = "" if p.name == "normal" else f" [{p.name} mode]"
                print(f"Listening... speak now.{tag}")
                return self.rec.listen(mic, timeout=6, phrase_time_limit=p.phrase_limit)
        except sr.WaitTimeoutError:
            raise NoSpeechError("I didn't hear anything.", silence=True)

    def _primary(self, audio):
        print("Recognizing...")
        return self._google(audio, PRIMARY_LANG)

    def _google(self, audio, lang):
        try:
            raw = self.rec.recognize_google(audio, language=lang, show_all=True)
        except sr.UnknownValueError:
            raise NoSpeechError("unintelligible")
        except sr.RequestError as e:
            raise EngineError(str(e))
        alts = raw.get("alternative", []) if isinstance(raw, dict) else []
        if not alts:
            raise NoSpeechError("unintelligible")
        top_conf = alts[0].get("confidence")
        out = []
        for i, a in enumerate(alts):
            conf = a.get("confidence")
            if conf is None:   # Google only scores the first alternative; decay the rest
                conf = (top_conf if top_conf is not None else 0.7) * (0.9 ** i)
            if a.get("transcript", "").strip():
                out.append(Hypothesis(a["transcript"].strip(), float(conf), f"google-{lang}"))
        if not out:
            raise NoSpeechError("unintelligible")
        return out

    def _sphinx(self, audio):
        try:
            text = self.rec.recognize_sphinx(audio)
        except sr.UnknownValueError:
            raise NoSpeechError("unintelligible")
        except Exception as e:
            raise EngineError(str(e))
        return [Hypothesis(text, 0.5, "sphinx")] if text else []


# --------------------------------------------------------------------------- #
# Voice input (same interface as before: listen() -> (text, error))
# --------------------------------------------------------------------------- #
class Listener:
    def __init__(self, confirm_fn=None, verbose=True):
        self.ok = sr is not None
        self.verbose = verbose
        self.tr = None
        if self.ok:
            self.tr = RobustTranscriber(MicBackend(), CommandVocabulary(), notify=print, confirm=confirm_fn)

    def listen(self):
        if not self.ok:
            return None, "Voice input needs:  pip install SpeechRecognition pyaudio"
        try:
            r = self.tr.transcribe()
        except TranscriptionFailed as e:
            return None, str(e)
        except Exception as e:  # missing mic / pyaudio, etc.
            return None, f"Microphone problem: {e}"
        if r.recovered and self.verbose:
            print(f"  [voice recovery: {r.how} | {r.captures} capture(s) | {r.elapsed:.1f}s | profile={r.profile}]")
        return r.text, None


# --------------------------------------------------------------------------- #
# Demo: noisy cafe, navigating the browser by voice  (python this_file.py --demo-noisy)
# --------------------------------------------------------------------------- #
class ScriptedBackend:
    """Simulated microphone + engines: replays what a recognizer would return in a noisy cafe."""

    def __init__(self, clock):
        self.clock = clock
        self.script: list = []
        self.primary = self._engine("primary", 0.5)
        self.secondary = self._engine("secondary", 0.5)
        self.offline = self._engine("offline", 1.2)

    def load(self, captures):
        self.script = list(captures)

    def capture(self, profile):
        self.clock.advance(profile.calibrate + 2.0)       # noise calibration + ~2 s of speech
        item = self.script.pop(0)
        if isinstance(item, Exception):
            raise item
        return item

    def _engine(self, name, cost):
        def run(audio):
            self.clock.advance(cost)
            r = audio.get(name)
            if isinstance(r, Exception):
                raise r
            if r is None:
                raise NoSpeechError("unintelligible")
            return [Hypothesis(t, c, name) for t, c in r]
        return run


def run_noisy_demo():
    global DRY_RUN
    DRY_RUN = True                                         # describe actions instead of performing them
    clock = VirtualClock()
    backend = ScriptedBackend(clock)
    tr = RobustTranscriber(backend, CommandVocabulary(), notify=lambda m: None,
                           confirm=lambda q: False, clock=clock)
    unintel = NoSpeechError("unintelligible")

    scenes = [
        ("0. QUIET baseline - no recovery should kick in",
         "open notepad",
         [{"primary": [("open notepad", 0.93)]}]),
        ("1. Café chatter garbles the first words  ->  n-best alternatives",
         "go to youtube",
         [{"primary": [("goat to you tube", 0.55), ("go to YouTube", 0.50)]}]),
        ("2. Clattering cups clip the last word  ->  fuzzy repair",
         "scroll down",
         [{"primary": [("scroll doubt", 0.38), ("scroll dumb", 0.34)]}]),
        ("3. A loud burst drowns the speech  ->  noise profile + second-opinion engine",
         "new tab",
         [{}, {"primary": [("you tap", 0.30)], "secondary": [("new tab", 0.62)]}]),
        ("4. The café Wi-Fi drops  ->  offline engine failover",
         "search lo-fi music",
         [{"primary": EngineError("network unreachable"), "offline": [("search lo fi music", 0.50)]}]),
        ("5. Announcement over the PA covers everything  ->  bounded retries, graceful failure",
         "close tab",
         [{}, {}, {"primary": [("plus sub", 0.27)]}]),
    ]

    summary = []
    for title, said, captures in scenes:
        backend.load(captures)
        clock.t = 0.0
        print("\n" + "=" * 74)
        print(title)
        print(f'  User says: "{said}"        (starting profile: {PROFILES[tr.level].name})')
        try:
            r = tr.transcribe()
            for line in r.trace:
                print("   ", line)
            action = handle(r.text, lambda q: False)
            plural = "capture" if r.captures == 1 else "captures"
            print(f"  => heard as \"{r.text}\"  via {r.how.upper()}  ({r.captures} {plural}, {r.elapsed:.1f}s)")
            print(f"  => action: {action}")
            summary.append((said, r.how, r.captures, r.elapsed))
        except TranscriptionFailed as e:
            for line in tr.trace:
                print("   ", line)
            print(f"  => FAILED after {clock.t:.1f}s (bounded) - assistant says: \"{e}\"")
            summary.append((said, "failed gracefully", tr.MAX_CAPTURES, clock.t))

    print("\n" + "=" * 74)
    print("6. Back in a quiet corner - four clean commands relax the profile again")
    for said in ("go back", "refresh", "next tab", "scroll up"):
        backend.load([{"primary": [(said, 0.94)]}])
        clock.t = 0.0
        r = tr.transcribe()
        print(f'   "{said}" -> {r.how} in {r.elapsed:.1f}s  (profile now: {r.profile})')

    print("\n" + "=" * 74)
    print(f"{'command':<20}{'recovered via':<22}{'captures':<10}{'time':<8}")
    for said, how, caps, secs in summary:
        print(f"{said:<20}{how:<22}{caps:<10}{secs:.1f}s")
    print("\nSimulated times: ~3 s is the unchanged quiet-room baseline; every extra second is recovery cost.")


# --------------------------------------------------------------------------- #
# Main loop
# --------------------------------------------------------------------------- #
def confirm(prompt):
    try:
        return input(prompt).strip().lower() in ("y", "yes", "yeah", "sure", "ok", "okay")
    except EOFError:
        return False


def main():
    global SPEAK_ENABLED
    if "--demo-noisy" in sys.argv:
        run_noisy_demo()
        return
    listener = Listener(confirm)
    tts_ok = init_tts()
    voice_mode = False

    print("=" * 62)
    print("  AI Desktop Assistant - type a command, or say 'help'")
    print("  Voice: 'voice mode' (continuous) or 'listen' (one command)")
    print("=" * 62)
    if psutil is None:
        print("(tip: pip install psutil  -> enables system status analysis)")
    if not listener.ok:
        print("(tip: pip install SpeechRecognition pyaudio  -> enables voice input)")
    if not tts_ok:
        print("(tip: pip install pyttsx3  -> enables spoken replies)")

    pending = ALARMS.list()
    if pending:
        say(f"You have {len(pending)} active alarm(s)/timer(s). Say 'show alarms' to view them.", speak=False)

    while True:
        raw = None
        try:
            if voice_mode:
                heard, err = listener.listen()
                if err:
                    say(err)
                    if "needs:" in err or "Microphone problem" in err:
                        voice_mode = False
                        say("Switching back to text mode.")
                    continue
                print(f"You (voice): {heard}")
                raw = heard
            else:
                raw = input("You: ")
        except (KeyboardInterrupt, EOFError):
            print()
            break

        low = (raw or "").strip().lower()
        if low in ("voice mode", "voice on", "start voice", "voice"):
            if listener.ok:
                voice_mode = True
                say("Voice mode on. Say 'text mode' to go back to typing.")
            else:
                say("Voice input needs:  pip install SpeechRecognition pyaudio")
            continue
        if low in ("text mode", "voice off", "stop voice", "type mode"):
            voice_mode = False
            say("Text mode on.")
            continue
        if low in ("listen", "l", "v"):
            heard, err = listener.listen()
            if err:
                say(err)
                continue
            print(f"You (voice): {heard}")
            raw = heard
        if low in ("speak on", "talk on", "voice replies on"):
            if tts_ok:
                SPEAK_ENABLED = True
                say("I will now speak my replies.")
            else:
                say("Spoken replies need:  pip install pyttsx3")
            continue
        if low in ("speak off", "talk off", "voice replies off"):
            SPEAK_ENABLED = False
            say("Okay, I'll stay quiet.")
            continue

        try:
            reply = handle(raw, confirm)
        except Exception as e:  # never crash the assistant
            reply = f"Something went wrong: {e}"

        if reply is None:
            say("Goodbye!")
            break
        say(reply)


if __name__ == "__main__":
    main()
