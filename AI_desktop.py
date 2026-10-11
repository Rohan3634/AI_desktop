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
from datetime import datetime, timedelta
from pathlib import Path
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

HOME = Path.home()
DATA_DIR = HOME / ".ai_desktop_assistant"
DATA_DIR.mkdir(exist_ok=True)
ALARM_FILE = DATA_DIR / "alarms.json"
EVENT_DIR = DATA_DIR / "events"
EVENT_DIR.mkdir(exist_ok=True)

# Friendly folder names -> real paths
PLACES = {
    "home": HOME,
    "desktop": HOME / "Desktop",
    "documents": HOME / "Documents",
    "downloads": HOME / "Downloads",
    "pictures": HOME / "Pictures",
    "music": HOME / "Music",
    "videos": HOME / "Videos",
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
        webbrowser.open(url)  # opens in the system's default browser
        return f"Searching the web for '{query}' in your default browser."

    @staticmethod
    def open_website(target):
        target = target.strip()
        if not re.match(r"^https?://", target):
            target = "https://" + target
        webbrowser.open(target)
        return f"Opening {target}"

    # ------------------------- apps ------------------------------- #
    @staticmethod
    def open_app(name):
        key = name.lower().strip()
        target = APPS.get(key)
        if target is None:
            # fuzzy: partial match
            for k, v in APPS.items():
                if key in k or k in key:
                    target = v
                    break
        target = target or key
        try:
            if IS_WINDOWS:
                if target.endswith(":") or ":" in target and not target[1:3] == ":\\":
                    os.startfile(target)  # URI protocol such as ms-settings:
                else:
                    subprocess.Popen(f'start "" "{target}"', shell=True)
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
    def delete(text, confirm_fn):
        p = Skills._p(text)
        if not p.exists():
            return f"I can't find {p}"
        if confirm_fn(f"Delete '{p}'? This cannot be undone. (yes/no): "):
            try:
                if p.is_dir():
                    shutil.rmtree(p)
                else:
                    p.unlink()
                return f"Deleted {p}"
            except Exception as e:
                return f"Couldn't delete: {e}"
        return "Cancelled. Nothing was deleted."

    @staticmethod
    def rename(src, new_name):
        s = Skills._p(src)
        if not s.exists():
            return f"I can't find {s}"
        d = s.with_name(new_name.strip().strip('"'))
        s.rename(d)
        return f"Renamed to {d.name}"

    @staticmethod
    def copy(src, dst):
        s, d = Skills._p(src), Skills._p(dst)
        if not s.exists():
            return f"I can't find {s}"
        if s.is_dir():
            shutil.copytree(s, d / s.name if d.is_dir() else d)
        else:
            shutil.copy2(s, d)
        return f"Copied {s.name} to {d}"

    @staticmethod
    def move(src, dst):
        s, d = Skills._p(src), Skills._p(dst)
        if not s.exists():
            return f"I can't find {s}"
        shutil.move(str(s), str(d))
        return f"Moved {s.name} to {d}"

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
        for root, dirs, files in os.walk(base):
            dirs[:] = [d for d in dirs if not d.startswith(".") and d not in ("AppData", "node_modules")]
            for f in files + dirs:
                if needle in f.lower():
                    hits.append(os.path.join(root, f))
                    if len(hits) >= 15:
                        break
            if len(hits) >= 15:
                break
        if not hits:
            return f"No files matching '{name}' found under {base}."
        return "Found:\n" + "\n".join("  " + h for h in hits)

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
# Voice input
# --------------------------------------------------------------------------- #
class Listener:
    def __init__(self):
        self.ok = sr is not None
        self.rec = sr.Recognizer() if self.ok else None

    def listen(self):
        if not self.ok:
            return None, "Voice input needs:  pip install SpeechRecognition pyaudio"
        try:
            with sr.Microphone() as mic:
                self.rec.adjust_for_ambient_noise(mic, duration=0.6)
                print("Listening... speak now.")
                audio = self.rec.listen(mic, timeout=6, phrase_time_limit=12)
            print("Recognizing...")
            return self.rec.recognize_google(audio), None
        except sr.WaitTimeoutError:
            return None, "I didn't hear anything."
        except sr.UnknownValueError:
            return None, "Sorry, I couldn't understand the audio."
        except sr.RequestError as e:
            return None, f"Speech service error (internet needed): {e}"
        except Exception as e:  # missing mic / pyaudio, etc.
            return None, f"Microphone problem: {e}"


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
    listener = Listener()
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
    