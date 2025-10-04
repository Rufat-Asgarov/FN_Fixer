import os
import sys
import time
import ctypes
import platform
import subprocess
import threading
from queue import Queue, Empty
from dataclasses import dataclass
from typing import Optional, List

# Third-party libraries
import keyboard
import pystray
from pystray import MenuItem as Item, Menu
from PIL import Image, ImageDraw, ImageFont

import screen_brightness_control as sbc

# Mic control (WASAPI via pycaw/comtypes)
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import IAudioEndpointVolume, AudioUtilities

IS_WINDOWS = (platform.system() == "Windows")
if not IS_WINDOWS:
    print("This script currently supports Windows only.")
    sys.exit(1)

# =============================
# Admin helpers (auto-elevate)
# =============================
def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def _quote_arg(s: str) -> str:
    if not s or any(c in s for c in ' \t"'):
        s = s.replace('"', r'\"')
        return f'"{s}"'
    return s

def run_as_admin():
    """
    Relaunch with admin rights (UAC). Works for python/pythonw and frozen EXE.
    """
    if getattr(sys, 'frozen', False):
        exe = sys.executable
        params = " ".join(map(_quote_arg, sys.argv[1:]))
    else:
        exe = sys.executable
        script = os.path.abspath(sys.argv[0])
        params = " ".join([_quote_arg(script)] + list(map(_quote_arg, sys.argv[1:])))
    ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, None, 1)
    if ret > 32:
        sys.exit(0)  # elevated instance launched; quit this one

if not is_admin():
    run_as_admin()

# =============================
# Global state
# =============================
@dataclass
class ModeState:
    media_mode: bool = False
    hotkey_handles: List = None

state = ModeState(media_mode=False, hotkey_handles=[])

# Brightness method cache to avoid slow probing every time
BRIGHT_METHOD_LAST: Optional[str] = None  # "wmi", "generic", or "ps"

# =============================
# Tray helpers
# =============================
def _measure_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont):
    if hasattr(draw, "textbbox"):
        l, t, r, b = draw.textbbox((0, 0), text, font=font)
        return r - l, b - t
    if hasattr(font, "getbbox"):
        l, t, r, b = font.getbbox(text)
        return r - l, b - t
    if hasattr(font, "getsize"):
        return font.getsize(text)
    return (8 * len(text), 16)

def make_icon(text: str, fill=(0, 0, 0), bg=(255, 255, 255)) -> Image.Image:
    size = (64, 64)
    img = Image.new("RGB", size, bg)
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 28)
    except Exception:
        font = ImageFont.load_default()
    w, h = _measure_text(d, text, font)
    d.text(((size[0]-w)/2, (size[1]-h)/2), text, fill=fill, font=font)
    return img

def update_tray_icon(icon: pystray.Icon):
    icon.icon = make_icon("FN" if not state.media_mode else "AU")
    icon.title = f"FnLock: {'Media' if state.media_mode else 'Normal'} mode"

def flash_tray_title(icon: pystray.Icon, msg: str, duration: float = 1.0):
    original = icon.title
    icon.title = f"{original} • {msg}"
    def _restore():
        time.sleep(duration)
        icon.title = original
    threading.Thread(target=_restore, daemon=True).start()

# =============================
# Background worker (non-blocking hotkeys)
# =============================
class ActionWorker:
    def __init__(self, icon: pystray.Icon):
        self.icon = icon
        self.q: Queue = Queue()
        self.thread = threading.Thread(target=self._loop, daemon=True)
        self._stop = threading.Event()
        self._bright_lock = threading.Lock()
        self._bright_pending = 0  # aggregated delta

    def start(self):
        self.thread.start()

    def stop(self):
        self._stop.set()
        self.q.put(("__quit__", None))
        self.thread.join(timeout=0.8)

    def enqueue(self, action: str, data=None):
        # Generic enqueue for quick tasks
        self.q.put((action, data))

    def add_brightness(self, delta: int):
        # Aggregate brightness changes to avoid spamming slow WMI/PS calls
        with self._bright_lock:
            self._bright_pending += delta
        self.q.put(("brightness_flush", None))

    def _loop(self):
        while not self._stop.is_set():
            try:
                action, data = self.q.get(timeout=0.5)
            except Empty:
                continue
            if action == "__quit__":
                break
            try:
                if action == "brightness_flush":
                    # debounce a bit to coalesce multiple F5/F6 presses
                    time.sleep(0.12)
                    with self._bright_lock:
                        delta = self._bright_pending
                        self._bright_pending = 0
                    if delta != 0:
                        safe_change_brightness(delta, self.icon)
                elif action == "mic_toggle":
                    safe_toggle_mic(self.icon)
                elif action == "volume":
                    # volume actions are fast but still off the hook thread
                    keyboard.send(data)
                elif action == "flash":
                    flash_tray_title(self.icon, str(data))
            except Exception:
                # never let the worker die
                try:
                    flash_tray_title(self.icon, "Action error")
                except Exception:
                    pass

WORKER: Optional[ActionWorker] = None  # set in run_tray()

# =============================
# Brightness controls (robust + cached)
# =============================
def clamp(val, lo=0, hi=100):
    return max(lo, min(hi, val))

def _avg(nums: List[int]) -> Optional[int]:
    vals = [n for n in nums if isinstance(n, (int, float))]
    if not vals:
        return None
    return int(round(sum(vals) / len(vals)))

def _ps_run(cmd: str):
    try:
        return subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
            capture_output=True, text=True, timeout=5
        )
    except Exception:
        return None

def _ps_get_brightness() -> Optional[int]:
    r = _ps_run("(Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightness | Select -First 1).CurrentBrightness")
    if not r or r.returncode != 0:
        return None
    for token in reversed(r.stdout.strip().split()):
        try:
            return int(token)
        except ValueError:
            pass
    return None

def _ps_set_brightness(val: int) -> bool:
    val = clamp(int(val), 0, 100)
    cmd = (
        f"$v={val}; "
        "Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightnessMethods "
        "| Invoke-CimMethod -MethodName WmiSetBrightness -Arguments @{Timeout=1; Brightness=$v} | Out-Null"
    )
    r = _ps_run(cmd)
    return bool(r and r.returncode == 0)

def _list_monitors_safe() -> List[str]:
    try:
        return sbc.list_monitors()
    except Exception:
        return []

def safe_change_brightness(delta: int, icon: Optional[pystray.Icon] = None):
    """
    Non-blocking-safe brightness change (called in worker thread).
    Tries cached method first, then WMI -> generic -> PowerShell; caches winner.
    """
    global BRIGHT_METHOD_LAST
    changed = False
    results: List[int] = []

    # Preferred target list
    monitors = _list_monitors_safe()
    internal = [m for m in monitors if any(k in m.lower() for k in ("integrated", "internal", "built-in", "edp"))]
    targets = internal if internal else monitors

    def try_wmi_targets():
        ok = False
        for m in targets:
            try:
                curr = sbc.get_brightness(display=m, method="wmi")
                curr_val = curr if isinstance(curr, int) else (curr[0] if curr else None)
                if curr_val is not None:
                    new_val = clamp(curr_val + delta, 0, 100)
                    sbc.set_brightness(new_val, display=m, method="wmi")
                    results.append(new_val)
                    ok = True
            except Exception:
                pass
        return ok

    def try_generic_all():
        ok = False
        for m in monitors or []:
            try:
                sbc.set_brightness(f"{'+' if delta>0 else ''}{delta}", display=m)
                try:
                    post = sbc.get_brightness(display=m)
                    post_val = post if isinstance(post, int) else (post[0] if post else None)
                    if post_val is not None:
                        results.append(post_val)
                except Exception:
                    pass
                ok = True
            except Exception:
                pass
        return ok

    def try_ps():
        cur = _ps_get_brightness()
        if cur is None:
            return False
        new_val = clamp(cur + delta, 0, 100)
        if _ps_set_brightness(new_val):
            results.append(new_val)
            return True
        return False

    # Build attempt order, preferring cached method
    methods = []
    if BRIGHT_METHOD_LAST:
        methods.append(BRIGHT_METHOD_LAST)
    for m in ("wmi", "generic", "ps"):
        if m not in methods:
            methods.append(m)

    for method in methods:
        ok = False
        if method == "wmi":
            ok = try_wmi_targets()
        elif method == "generic":
            ok = try_generic_all()
        elif method == "ps":
            ok = try_ps()
        if ok:
            BRIGHT_METHOD_LAST = method
            changed = True
            break

    # Feedback (never block)
    if icon:
        if changed:
            avg = _avg(results)
            if avg is not None:
                flash_tray_title(icon, f"Brightness: {avg}%")
            else:
                flash_tray_title(icon, f"Brightness: adjusted ({'+' if delta>0 else ''}{delta})")
        else:
            flash_tray_title(icon, "Brightness: not supported")

# =============================
# Microphone mute/unmute (safe)
# =============================
WM_APPCOMMAND = 0x0319
APPCOMMAND_MIC_MUTE = 0x180000
user32 = ctypes.windll.user32

def _mic_toggle_via_appcommand() -> bool:
    try:
        hwnd = user32.GetForegroundWindow()
        user32.SendMessageW(hwnd, WM_APPCOMMAND, 0, APPCOMMAND_MIC_MUTE)
        return True
    except Exception:
        return False

DATAFLOW_CAPTURE = 1  # EDataFlow.eCapture
ROLE_CONSOLE = 0      # ERole.eConsole
ROLE_COMMUNICATIONS = 2  # ERole.eCommunications

def _get_default_capture_endpoint_volume(role: int):
    try:
        dev = AudioUtilities.GetDefaultAudioEndpoint(DATAFLOW_CAPTURE, role)
        interface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(interface, POINTER(IAudioEndpointVolume))
    except Exception:
        return None

def safe_toggle_mic(icon: Optional[pystray.Icon] = None):
    try:
        vol = _get_default_capture_endpoint_volume(ROLE_CONSOLE)
        if vol is None:
            vol = _get_default_capture_endpoint_volume(ROLE_COMMUNICATIONS)

        status = None
        if vol is not None:
            try:
                is_muted = vol.GetMute()
                new_state = 0 if is_muted else 1
                vol.SetMute(new_state, None)
                status = "unmuted" if new_state == 0 else "muted"
            except Exception:
                vol = None

        if vol is None:
            ok = _mic_toggle_via_appcommand()
            status = "toggled" if ok else "error"

        if icon:
            flash_tray_title(icon, f"Mic: {status}")
    except Exception:
        if icon:
            flash_tray_title(icon, "Mic: error")

# =============================
# Hotkeys (bind/unbind)
# =============================
def bind_media_hotkeys(icon: Optional[pystray.Icon] = None):
    """Bind F1..F6 in Media mode. Callbacks enqueue work; never heavy inline."""
    unbind_media_hotkeys()
    # Volume via worker (fast anyway)
    h1 = keyboard.add_hotkey('f1', lambda: WORKER.enqueue("volume", 'volume mute'), suppress=True)
    h2 = keyboard.add_hotkey('f2', lambda: WORKER.enqueue("volume", 'volume down'), suppress=True)
    h3 = keyboard.add_hotkey('f3', lambda: WORKER.enqueue("volume", 'volume up'),   suppress=True)
    # Mic & Brightness via worker (non-blocking)
    h4 = keyboard.add_hotkey('f4', lambda: WORKER.enqueue("mic_toggle"), suppress=True)
    h5 = keyboard.add_hotkey('f5', lambda: WORKER.add_brightness(-10),   suppress=True)
    h6 = keyboard.add_hotkey('f6', lambda: WORKER.add_brightness(+10),   suppress=True)
    state.hotkey_handles = [h1, h2, h3, h4, h5, h6]

def unbind_media_hotkeys():
    if state.hotkey_handles:
        for h in state.hotkey_handles:
            try:
                keyboard.remove_hotkey(h)
            except KeyError:
                pass
        state.hotkey_handles = []

def set_mode(media: bool, tray_icon: Optional[pystray.Icon] = None):
    state.media_mode = media
    if state.media_mode:
        bind_media_hotkeys(tray_icon)
    else:
        unbind_media_hotkeys()
    if tray_icon is not None:
        update_tray_icon(tray_icon)

# =============================
# Num Lock toggle
# =============================
def on_numlock_toggle(tray_icon: Optional[pystray.Icon] = None):
    set_mode(not state.media_mode, tray_icon)
    if tray_icon is not None:
        flash_tray_title(tray_icon, "Toggled")

def install_numlock_listener(tray_icon: Optional[pystray.Icon]):
    keyboard.add_hotkey('num lock', lambda: on_numlock_toggle(tray_icon), suppress=False)

# =============================
# Tray application
# =============================
def run_tray():
    global WORKER

    icon = pystray.Icon("FnLock")

    def toggle_action(icon, item):
        on_numlock_toggle(icon)

    def quit_action(icon, item):
        try:
            if WORKER:
                WORKER.stop()
        except Exception:
            pass
        icon.visible = False
        time.sleep(0.2)
        try:
            unbind_media_hotkeys()
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        icon.stop()

    def mode_label(_):
        return ("Mode: Media "
                "(F1=Mute, F2=Vol-, F3=Vol+, F4=Mic, F5=Bright-, F6=Bright+)"
                if state.media_mode else
                "Mode: Normal (F1–F6)")

    icon.menu = Menu(
        Item(mode_label, lambda *_: None, enabled=False),
        Item("Toggle (Num Lock)", toggle_action),
        Item("Quit", quit_action)
    )
    update_tray_icon(icon)

    # Start background worker and listeners
    WORKER = ActionWorker(icon)
    WORKER.start()

    # Start in Normal mode
    set_mode(False, icon)
    install_numlock_listener(icon)

    icon.run()

if __name__ == "__main__":
    run_tray()
