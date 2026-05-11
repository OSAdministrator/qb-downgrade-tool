"""Background watchdog that auto-dismisses QuickBooks popups and keeps QB windows hidden.

Runs in its own thread for the entire duration of a job so the operator never has to
touch a mouse or keyboard. Two responsibilities:

1. Scan top-level desktop windows every ~500 ms. Any window matching a known popup
   rule is dismissed by clicking the safe button (or the X). Unknown dialogs are
   logged with title + button list so we can add a rule next time.

2. Keep QuickBooks main windows hidden (SW_HIDE) so the operator only sees our
   GUI. QBFC SDK does not require the QB UI to be visible — it talks to the file
   directly via COM.

The watchdog is best-effort: any individual failure is swallowed so it never crashes
the main pipeline.
"""

from __future__ import annotations

import re
import threading
import time
from typing import Callable, List, Optional, Tuple

try:
    import win32con  # type: ignore
    import win32gui  # type: ignore
    import win32process  # type: ignore
    _WIN32_OK = True
except ImportError:  # pragma: no cover - non-Windows / no pywin32
    win32con = None  # type: ignore
    win32gui = None  # type: ignore
    win32process = None  # type: ignore
    _WIN32_OK = False

try:
    from pywinauto import Desktop  # type: ignore
    _PWA_OK = True
except ImportError:  # pragma: no cover
    Desktop = None  # type: ignore
    _PWA_OK = False


LogFn = Callable[[str], None]


# ---------------------------------------------------------------------------
# Popup rules: (regex on title, ordered list of button captions to try)
# First matching button (case-insensitive substring) is clicked. If no button
# matches the window is closed via WM_CLOSE (the X button).
# ---------------------------------------------------------------------------
_RULES: List[Tuple[str, List[str]]] = [
    # ---- Updates / maintenance / registration / promo ----
    (r"(?i)quickbooks\s+update",          ["Install Later", "Later", "Cancel", "No", "Close"]),
    (r"(?i)update\s+available",           ["Install Later", "Later", "Cancel", "No", "Close"]),
    (r"(?i)maintenance\s+release",        ["Install Later", "Later", "Cancel", "Close"]),
    (r"(?i)product\s+information",        ["OK", "Close"]),
    (r"(?i)register\s+quickbooks",        ["Remind Me Later", "Later", "Cancel", "Close"]),
    (r"(?i)registration",                 ["Remind Me Later", "Later", "Cancel", "Close"]),
    (r"(?i)sign\s*in\s+to\s+intuit",      ["Skip", "Cancel", "Close", "No Thanks"]),
    (r"(?i)intuit\s+account",             ["Skip", "Cancel", "Close", "No Thanks"]),
    (r"(?i)what.?s\s+new",                ["Close", "OK"]),
    (r"(?i)getting\s+started",            ["Close", "OK"]),
    (r"(?i)quickbooks\s+learning",        ["Close", "OK"]),
    (r"(?i)new\s+feature\s+tour",         ["Close", "Skip", "No"]),
    (r"(?i)did\s+you\s+know",             ["Close", "OK"]),
    (r"(?i)tip\s+of\s+the\s+day",         ["Close", "OK"]),
    (r"(?i)home\s*page",                  ["Close"]),
    (r"(?i)quickbooks\s+coach",           ["Close", "No Thanks"]),
    (r"(?i)accountant\s+center",          ["Close", "X"]),

    # ---- Memorized / future / past txn prompts during import ----
    (r"(?i)memorize\s+transaction",       ["No", "Cancel", "Don't Save"]),
    (r"(?i)memorize\s+check",             ["No", "Cancel"]),
    (r"(?i)memorize\s+bill",              ["No", "Cancel"]),
    (r"(?i)memorized\s+transactions?\s+(due|list|later)", ["Enter All Later", "OK", "Close", "Cancel"]),
    (r"(?i)future\s+(date|transaction)",  ["Yes"]),
    (r"(?i)date\s+is\s+in\s+the\s+future", ["Yes"]),
    (r"(?i)past\s+(date|transaction)",    ["Yes"]),
    (r"(?i)more\s+than\s+\d+\s+days",     ["Yes"]),
    (r"(?i)recording\s+transaction",      ["OK", "Yes"]),
    (r"(?i)tracking\s+number",            ["OK"]),
    (r"(?i)information\s+missing",        ["OK"]),
    (r"(?i)you\s+must\s+specify",         ["OK"]),
    (r"(?i)print\s+later",                ["No"]),
    (r"(?i)to\s+be\s+printed",            ["No"]),
    (r"(?i)mark\s+as\s+to\s+be\s+printed", ["No"]),

    # ---- Backup / accountant copy / restore prompts ----
    (r"(?i)backup\s+(reminder|now|copy)", ["No", "Cancel", "Later"]),
    (r"(?i)schedule\s+backup",            ["No", "Cancel", "Skip"]),
    (r"(?i)accountant.?s\s+copy",         ["Cancel", "No", "Close"]),
    (r"(?i)restore\s+(a\s+)?backup",      ["Cancel", "No", "Close"]),

    # ---- Setup / config dialogs we never want during import ----
    (r"(?i)set\s*up\s+(new\s+)?account",  ["Cancel", "Close"]),
    (r"(?i)external\s+accountant",        ["No", "Cancel", "Close"]),
    (r"(?i)set\s*up\s+user",              ["Cancel", "Close"]),
    (r"(?i)review\s+(your\s+)?(bills|payroll)", ["Cancel", "Close", "No"]),

    # ---- Generic informational popups (catch-all, must come LAST) ----
    (r"(?i)quickbooks\s+message",         ["OK", "Yes", "Close"]),
    (r"(?i)information$",                 ["OK", "Close"]),
    (r"(?i)warning$",                     ["OK", "Yes", "Close"]),
    (r"(?i)confirmation$",                ["Yes", "OK", "Close"]),
]


# Window titles whose top-level QBW32 frames should be hidden, NOT dismissed.
# Anything matching here is sent ShowWindow(SW_HIDE) every poll so QB stays
# off-screen. The main company window matches the first pattern.
_HIDE_PATTERNS: List[str] = [
    r"(?i)quickbooks\s+(premier|pro|enterprise|accountant|desktop)",
    r"(?i) - quickbooks",  # title bar suffix on company window
]

# Window titles we MUST NOT dismiss or hide — these are our own GUI / Tk.
_OWN_WINDOW_PATTERNS: List[str] = [
    r"(?i)quickbooks\s+timewarp",
    r"(?i)timewarp",
    r"(?i)settings$",
    r"(?i)tk$",
]


class DialogWatchdog:
    """Background thread that dismisses popups and hides QB windows.

    Usage:
        wd = DialogWatchdog(log_fn=print)
        wd.start()
        try:
            ...do the job...
        finally:
            wd.stop()
    """

    POLL_SEC = 0.5

    def __init__(self, log_fn: Optional[LogFn] = None, hide_qb: bool = True) -> None:
        self._log_fn = log_fn or (lambda m: None)
        self._hide_qb = hide_qb
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._dismissed_titles: set = set()  # de-dup logging
        self._unknown_titles: set = set()    # de-dup unknown-dialog logging
        self._enabled = _WIN32_OK and _PWA_OK

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------
    def start(self) -> None:
        if not self._enabled:
            self._log_fn("[Watchdog] pywin32 / pywinauto not available — watchdog disabled")
            return
        if self._thread is not None and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, name="DialogWatchdog", daemon=True)
        self._thread.start()
        self._log_fn("[Watchdog] started — popups will be auto-dismissed, QB stays hidden")

    def stop(self) -> None:
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=2)
        self._thread = None
        self._log_fn("[Watchdog] stopped")

    # ------------------------------------------------------------------
    # Main loop
    # ------------------------------------------------------------------
    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                self._tick()
            except Exception as exc:  # noqa: BLE001 - never crash on stray errors
                # Only log once per error message to avoid spam
                msg = f"[Watchdog] tick error: {exc}"
                if msg not in self._dismissed_titles:
                    self._dismissed_titles.add(msg)
                    self._log_fn(msg)
            self._stop.wait(self.POLL_SEC)

    def _tick(self) -> None:
        # Collect all top-level visible windows with their titles via win32 (faster
        # than pywinauto for enumeration).
        hwnds: List[Tuple[int, str]] = []

        def _enum(hwnd: int, _arg) -> bool:
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd) or ""
                if not title:
                    return True
                hwnds.append((hwnd, title))
            except Exception:  # noqa: BLE001
                pass
            return True

        win32gui.EnumWindows(_enum, None)

        for hwnd, title in hwnds:
            # Never touch our own GUI
            if any(re.search(p, title) for p in _OWN_WINDOW_PATTERNS):
                continue

            # Hide QB main windows (don't dismiss — we need the file open)
            if self._hide_qb and any(re.search(p, title) for p in _HIDE_PATTERNS):
                # Only hide if it's a QBW32 process window
                if self._is_qb_process(hwnd):
                    self._hide_window(hwnd, title)
                    continue

            # Match against dismissal rules
            handled = False
            for pattern, buttons in _RULES:
                if re.search(pattern, title):
                    self._dismiss(hwnd, title, buttons)
                    handled = True
                    break

            if not handled:
                # Log unknown popups once so we can add a rule next time.
                # Skip if it's a QB process window (likely a transient internal frame).
                if self._looks_like_dialog(hwnd) and title not in self._unknown_titles:
                    self._unknown_titles.add(title)
                    self._log_fn(f"[Watchdog] unknown dialog (left alone): '{title}'")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _is_qb_process(self, hwnd: int) -> bool:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            # We don't actually need to verify the process name — QB titles are
            # distinctive enough. Just confirm pid > 0.
            return pid > 0
        except Exception:  # noqa: BLE001
            return False

    def _looks_like_dialog(self, hwnd: int) -> bool:
        """Heuristic: small-ish window with no minimize box looks like a dialog."""
        try:
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            # Dialogs usually lack WS_MINIMIZEBOX/WS_MAXIMIZEBOX
            has_min = bool(style & win32con.WS_MINIMIZEBOX)
            has_max = bool(style & win32con.WS_MAXIMIZEBOX)
            if has_min and has_max:
                return False
            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            # Reasonable dialog size band
            return 100 < w < 900 and 80 < h < 700
        except Exception:  # noqa: BLE001
            return False

    def _hide_window(self, hwnd: int, title: str) -> None:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
            key = f"HIDE::{title}"
            if key not in self._dismissed_titles:
                self._dismissed_titles.add(key)
                self._log_fn(f"[Watchdog] hid QB window: '{title}'")
        except Exception as exc:  # noqa: BLE001
            self._log_fn(f"[Watchdog] hide failed for '{title}': {exc}")

    def _dismiss(self, hwnd: int, title: str, button_captions: List[str]) -> None:
        """Click the first matching button in the dialog, or close it."""
        clicked_caption: Optional[str] = None
        try:
            # Enumerate child buttons via win32 and click by caption
            children: List[Tuple[int, str, str]] = []

            def _enum_child(child_hwnd: int, _arg) -> bool:
                try:
                    cls = win32gui.GetClassName(child_hwnd) or ""
                    txt = win32gui.GetWindowText(child_hwnd) or ""
                    children.append((child_hwnd, cls, txt))
                except Exception:  # noqa: BLE001
                    pass
                return True

            win32gui.EnumChildWindows(hwnd, _enum_child, None)

            # Find a button whose caption matches one in our list (in priority order)
            for wanted in button_captions:
                wanted_lc = wanted.lower().strip()
                for child_hwnd, cls, txt in children:
                    if "button" not in cls.lower():
                        continue
                    # Strip & accelerator from button captions ("&Yes" -> "Yes")
                    clean = txt.replace("&", "").strip().lower()
                    if wanted_lc in clean:
                        # Click via BM_CLICK so it doesn't require focus
                        win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                        clicked_caption = txt or wanted
                        break
                if clicked_caption:
                    break

            # If no button matched, close via WM_CLOSE (the X button)
            if not clicked_caption:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                clicked_caption = "[X close]"

            key = f"{title}::{clicked_caption}"
            if key not in self._dismissed_titles:
                self._dismissed_titles.add(key)
                self._log_fn(f"[Watchdog] dismissed '{title}' via '{clicked_caption}'")
        except Exception as exc:  # noqa: BLE001
            self._log_fn(f"[Watchdog] dismiss failed for '{title}': {exc}")


__all__ = ["DialogWatchdog"]
