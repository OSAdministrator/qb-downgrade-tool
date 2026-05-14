"""QuickBooks desktop automation engine.

Designed for Windows + QuickBooks Desktop 2023 and 2021 environments.
Uses pywinauto/pyautogui when available, with a dry-run fallback for testing.
"""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

from config import AppConfig
from transaction_parser import TransactionParser
from validator import CompanyValidationSnapshot, ValidationReportBuilder

try:
    import pyautogui
except Exception:  # noqa: BLE001
    pyautogui = None

try:
    from pywinauto import Desktop
    from pywinauto.application import Application
    from pywinauto.keyboard import send_keys
except Exception:  # noqa: BLE001
    Application = None
    Desktop = None
    send_keys = None


logger = logging.getLogger(__name__)

LogFn = Callable[[str], None]
ProgressFn = Callable[[float], None]


@dataclass
class CompanyJob:
    qbw_path: Path
    password: str
    output_dir: Path
    target_company_name: Optional[str] = None


@dataclass
class CompanyJobResult:
    qbw_path: str
    success: bool
    message: str
    generated_files: Dict[str, str] = field(default_factory=dict)


class QuickBooksAutomationEngine:
    """End-to-end orchestrator for one company file."""

    # Match QB windows but EXCLUDE our own GUI ("QuickBooks TimeWarp®")
    QB_WINDOW_RE = r"(?i).*quickbooks.*"
    _OWN_GUI_KEYWORDS = ("timewarp", "downgrade tool")

    def __init__(self, config: AppConfig, logger: Optional[logging.Logger] = None):
        self.config = config
        self.logger = logger or logging.getLogger("qb_downgrade")
        self.tx_parser = TransactionParser()
        self.validator = ValidationReportBuilder()
        self._qb2023_app: Optional[object] = None
        self._qb2021_app: Optional[object] = None
        self._watchdog: Optional[object] = None  # DialogWatchdog reference for pause/resume
        # Tracks whether password was already entered during QB startup dialogs.
        # Used to prevent redundant _open_company_file() call in _export_from_qb2023().
        # See bug note in _export_from_qb2023() for details.
        self._startup_password_handled: bool = False

    def _emit(self, msg: str, log_fn: Optional[LogFn]) -> None:
        # Sanitize non-ASCII characters that break Windows console/file encoding
        safe_msg = msg.encode('ascii', 'replace').decode('ascii')
        self.logger.info(safe_msg)
        if log_fn:
            log_fn(safe_msg)

    def _with_retries(self, fn: Callable[[], None], step_name: str, log_fn: Optional[LogFn]) -> None:
        attempts = self.config.retry_attempts + 1
        for attempt in range(1, attempts + 1):
            try:
                fn()
                return
            except Exception as exc:  # noqa: BLE001
                screenshot_path = self._take_error_screenshot(step_name.replace(" ", "_"))
                self._emit(
                    f"{step_name} failed (attempt {attempt}/{attempts}): {exc}. "
                    f"Screenshot: {screenshot_path if screenshot_path else 'unavailable'}",
                    log_fn,
                )
                if attempt == attempts:
                    raise
                time.sleep(self.config.retry_delay_seconds)

    def _ensure_automation_ready(self) -> None:
        if os.name != "nt":
            raise RuntimeError("QuickBooks automation only runs on Windows.")
        if Application is None or Desktop is None or send_keys is None:
            raise RuntimeError("pywinauto is not available. Install dependencies on Windows.")

    def _wait_until(self, condition: Callable[[], bool], timeout_s: int, interval_s: float = 0.5) -> bool:
        start = time.time()
        while time.time() - start < timeout_s:
            try:
                if condition():
                    return True
            except Exception:  # noqa: BLE001
                pass
            time.sleep(interval_s)
        return False

    def _take_error_screenshot(self, label: str) -> Optional[Path]:
        if pyautogui is None:
            return None
        try:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_label = re.sub(r"[^a-zA-Z0-9_.-]+", "_", label)
            screenshot_dir = Path(self.config.default_output_dir) / "logs"
            screenshot_dir.mkdir(parents=True, exist_ok=True)
            screenshot_path = screenshot_dir / f"error_{safe_label}_{stamp}.png"
            pyautogui.screenshot(str(screenshot_path))
            return screenshot_path
        except Exception:  # noqa: BLE001
            return None

    def _get_desktop(self):
        self._ensure_automation_ready()
        return Desktop(backend="uia")

    def _find_qb_main_window(self, app: Optional[object], version_hint: Optional[str], timeout_s: int, include_hidden: bool = False, log_fn: Optional[LogFn] = None):
        """Find the QB main window.

        When *include_hidden* is True the visibility check is skipped so
        windows hidden by the watchdog are still returned.  This is
        necessary for the QB 2021 import phase because the watchdog hides
        the main window before the company fully loads.
        """
        def _is_own_gui(title: str) -> bool:
            """Return True if *title* belongs to our own GUI, not real QB."""
            tl = title.lower()
            return any(kw in tl for kw in self._OWN_GUI_KEYWORDS)

        def _pick_window() -> Optional[object]:
            if app is not None:
                try:
                    for win in app.windows():
                        if not win.exists():
                            continue
                        if not include_hidden and not win.is_visible():
                            continue
                        title = win.window_text()
                        if _is_own_gui(title):
                            continue
                        if version_hint and version_hint not in title:
                            continue
                        if re.search(self.QB_WINDOW_RE, title):
                            return win
                    top = app.top_window()
                    if top.exists() and (include_hidden or top.is_visible()):
                        title = top.window_text() or ""
                        if not _is_own_gui(title):
                            return top
                except Exception:  # noqa: BLE001
                    pass

            desktop = self._get_desktop()
            candidates = []
            for w in desktop.windows(title_re=self.QB_WINDOW_RE, visible_only=not include_hidden):
                try:
                    if not include_hidden and not w.is_visible():
                        continue
                    title = w.window_text()
                    if _is_own_gui(title):
                        continue
                    score = 0
                    if version_hint and version_hint in title:
                        score += 10
                    if "No Company Open" not in title:
                        score += 2
                    candidates.append((score, w))
                except Exception:  # noqa: BLE001
                    continue
            if not candidates:
                return None
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1]
        box: Dict[str, object] = {}
        poll_count = 0

        def _cond() -> bool:
            nonlocal poll_count
            poll_count += 1
            win = _pick_window()
            if win is not None:
                box["window"] = win
                return True
            # Log a window census every 30 polls to aid debugging
            if poll_count % 30 == 0:
                try:
                    import win32gui  # type: ignore[import-untyped]
                    titles = []
                    def _enum_cb(hwnd, _):
                        if win32gui.IsWindowVisible(hwnd):
                            t = win32gui.GetWindowText(hwnd)
                            if t:
                                titles.append(t)
                        return True
                    win32gui.EnumWindows(_enum_cb, None)
                    if log_fn:
                        log_fn(f"  [WindowCensus] Poll #{poll_count}: {len(titles)} visible windows")
                        for t in titles:
                            if "quickbooks" in t.lower() or "qb" in t.lower() or "intuit" in t.lower():
                                log_fn(f"    QB-related: '{t}'")
                except Exception:
                    pass
            return False

        if not self._wait_until(_cond, timeout_s, 1.0):
            # Last-ditch: try win32gui.FindWindow for any QB window
            try:
                import win32gui  # type: ignore[import-untyped]
                all_qb = []
                def _enum_all(hwnd, _):
                    t = win32gui.GetWindowText(hwnd)
                    if t and "quickbooks" in t.lower():
                        all_qb.append((hwnd, t, win32gui.IsWindowVisible(hwnd)))
                    return True
                win32gui.EnumWindows(_enum_all, None)
                if all_qb and log_fn:
                    log_fn(f"  [win32gui] Found {len(all_qb)} QB windows after timeout:")
                    for hwnd, t, vis in all_qb:
                        log_fn(f"    hwnd={hwnd} vis={vis} title='{t}'")
                    # If we found a matching window, try to wrap it
                    for hwnd, t, vis in all_qb:
                        if version_hint and version_hint in t:
                            if not vis:
                                import win32con  # type: ignore[import-untyped]
                                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                            try:
                                desktop = self._get_desktop()
                                for w in desktop.windows(visible_only=False):
                                    try:
                                        if w.handle == hwnd:
                                            box["window"] = w
                                            if log_fn:
                                                log_fn(f"  [win32gui] Recovered QB window via hwnd={hwnd}")
                                            return box["window"]
                                    except Exception:
                                        continue
                            except Exception:
                                pass
            except Exception:
                pass

            hint_text = f" ({version_hint})" if version_hint else ""
            raise RuntimeError(f"Could not find QuickBooks main window{hint_text} within {timeout_s}s")

        return box["window"]
        return box["window"]

    def _focus_window(self, window) -> None:
        try:
            window.set_focus()
        except Exception:  # noqa: BLE001
            pass

    def _click_first_button(self, dialog, labels: List[str]) -> bool:
        for label in labels:
            try:
                btn = dialog.child_window(title_re=fr"(?i){re.escape(label)}", control_type="Button")
                if btn.exists(timeout=1):
                    btn.wrapper_object().click_input()
                    return True
            except Exception:  # noqa: BLE001
                continue
        # Fallback: try without control_type filter (some buttons are custom controls)
        for label in labels:
            try:
                btn = dialog.child_window(title_re=fr"(?i){re.escape(label)}")
                if btn.exists(timeout=0.5):
                    try:
                        btn.wrapper_object().click_input()
                    except Exception:
                        # Last resort: focus and send Enter
                        try:
                            btn.set_focus()
                            time.sleep(0.1)
                            if send_keys is not None:
                                send_keys("{ENTER}")
                        except Exception:
                            continue
                    return True
            except Exception:  # noqa: BLE001
                continue
        return False

    def _set_edit_value(self, parent, value: str, edit_index: int = 0) -> bool:
        try:
            edits = parent.descendants(control_type="Edit")
            if not edits or edit_index >= len(edits):
                return False
            edit = edits[edit_index]
            edit.set_focus()
            time.sleep(0.1)
            # Select all existing text first, then replace
            if send_keys is not None:
                send_keys("^a")
                time.sleep(0.05)
            try:
                edit.set_edit_text(value)
            except Exception:  # noqa: BLE001
                # Fallback: type the value character by character
                if send_keys is not None:
                    send_keys("{DELETE}")
                    time.sleep(0.05)
                    edit.type_keys(value, with_spaces=True, pause=0.02)
            return True
        except Exception:  # noqa: BLE001
            return False

    def _find_active_dialog(self, title_re: Optional[str] = None, parent_window=None):
        """Find an active dialog window matching title_re.

        Args:
            title_re: Optional regex to match against window titles.
            parent_window: Optional parent window to search children of.
                If provided, searches parent_window.children(visible_only=True)
                instead of Desktop().windows().

        # FIX: Search child dialogs of QB main window instead of desktop-level windows.
        # QB Export dialogs are modal children, not top-level, so Desktop().windows()
        # can't see them. This was causing the tool to loop ~20 times then timeout
        # (found in user testing 2026-05-08).
        #
        # WHY we search child windows: QuickBooks creates its export/save/import dialogs
        # as modal children of the main application window. These child dialogs do not
        # appear in the list returned by Desktop().windows(), which only enumerates
        # top-level windows. By searching parent_window.children() when a parent is
        # provided, we can detect these modal dialogs reliably and avoid the timeout loop.
        """
        if parent_window is not None:
            # Search child windows of the given parent (e.g., QB main window).
            # This is the correct approach for modal dialogs like Export/Save As
            # which are children of the main QB window, not top-level desktop windows.
            try:
                children = parent_window.children()
            except Exception:  # noqa: BLE001
                children = []
            for w in children:
                try:
                    if not w.is_visible():
                        continue
                    title = w.window_text() or ""
                    if not title:
                        continue
                    if title_re and not re.search(title_re, title):
                        continue
                    return w
                except Exception:  # noqa: BLE001
                    continue
            # Fall through to desktop-level search: Windows common dialogs
            # (Save Print Output As, Open, etc.) are top-level windows, NOT
            # children of QB main_window, so we must also check Desktop().windows().

        # Fallback: search top-level desktop windows (for dialogs not tied to a parent)
        desktop = self._get_desktop()
        windows = desktop.windows()
        for w in windows:
            try:
                if not w.is_visible():
                    continue
                title = w.window_text() or ""
                if not title:
                    continue
                if title_re and not re.search(title_re, title):
                    continue
                return w
            except Exception:  # noqa: BLE001
                continue
        return None

    def _dismiss_common_dialogs(self, log_fn: Optional[LogFn]) -> None:
        desktop = self._get_desktop()
        for dialog in desktop.windows():
            try:
                if not dialog.is_visible():
                    continue
                title = dialog.window_text() or ""
                if not title:
                    continue
                # Skip the main QB window itself
                if re.search(self.QB_WINDOW_RE, title) and "login" not in title.lower():
                    continue

                title_l = title.lower()

                # Handle password/login dialogs at startup
                if "login" in title_l or "password" in title_l:
                    # Don't dismiss these - they need special handling
                    continue

                # Handle "Set Up an External Accountant User" dialog
                if "external accountant" in title_l or "accountant user" in title_l:
                    clicked = self._click_first_button(dialog, ["No", "Cancel", "Close"])
                    if clicked:
                        self._emit(f"Dismissed dialog: {title}", log_fn)
                    continue

                # Handle "Enter Memorized Transactions" dialog
                if "memorized transaction" in title_l:
                    clicked = self._click_first_button(dialog, ["Enter All Later", "Close", "Cancel", "No"])
                    if clicked:
                        self._emit(f"Dismissed dialog: {title}", log_fn)
                    time.sleep(0.5)
                    # There may be a follow-up "Enter Memorized Transactions Later" dialog
                    followup = self._find_active_dialog(title_re=r"(?i)memorized.*later")
                    if followup is not None:
                        self._click_first_button(followup, ["OK", "Close"])
                        try:
                            # Try clicking the X button
                            followup.close()
                        except Exception:  # noqa: BLE001
                            if send_keys is not None:
                                send_keys("{ENTER}")
                        self._emit("Dismissed memorized transactions follow-up dialog", log_fn)
                    continue

                if any(
                    x in title_l
                    for x in [
                        "update",
                        "backup",
                        "confirmation",
                        "quickbooks message",
                        "information",
                        "warning",
                        "microsoft office",
                        "register",
                        "product information",
                        "new feature",
                        "intuit",
                        "usage",
                        "analytics",
                        "study",
                        "survey",
                        "have a question",
                        "faq",
                    ]
                ):
                    clicked = self._click_first_button(
                        dialog,
                        ["Continue", "OK", "Yes", "No", "Don't Save", "Close", "Skip", "Later", "Cancel"],
                    )
                    if clicked:
                        self._emit(f"Dismissed dialog: {title}", log_fn)
            except Exception:  # noqa: BLE001
                continue

    def _close_popup_windows(self, main_window, log_fn: Optional[LogFn]) -> None:
        """Close QB popup/helper windows that steal focus after login.

        Common culprits include the Accountant Center, Getting Started,
        Learning Center, What's New, and Home Page windows.  These are
        enumerated from the desktop and closed individually so that the
        main QuickBooks company window can regain focus for menu navigation.
        """
        desktop = self._get_desktop()

        # Substrings (lower-cased) of window titles that should be closed
        popup_hints = [
            "accountant center",
            "getting started",
            "what's new",
            "whats new",
            "tips",
            "quickbooks learning center",
            "learning center",
            "home page",
            "quickbooks home",
            "new feature",
            "coach",
            "did you know",
            "quickbooks desktop",  # generic splash / promo windows
            "usage",
            "analytics",
            "study",
            "have a question",
            "faq",
            "enterprise",
            "upgrade",
            "get the latest",
        ]

        closed_any = False
        for win in desktop.windows():
            try:
                if not win.is_visible():
                    continue
                title = win.window_text() or ""
                if not title:
                    continue

                # Never close the main company window itself
                if win.handle == getattr(main_window, "handle", None):
                    continue
                # Also skip anything that looks like the main QB window
                if re.search(self.QB_WINDOW_RE, title) and not any(h in title.lower() for h in popup_hints):
                    continue

                title_l = title.lower()
                if any(hint in title_l for hint in popup_hints):
                    # Try clicking common close/dismiss buttons first
                    clicked = self._click_first_button(
                        win, ["Continue", "Close", "OK", "No", "Skip", "Later", "X", "Cancel"]
                    )
                    if not clicked:
                        # Fall back to closing the window directly
                        try:
                            win.close()
                        except Exception:  # noqa: BLE001
                            pass
                    self._emit(f"Closed popup window: {title}", log_fn)
                    closed_any = True
                    time.sleep(0.5)
            except Exception:  # noqa: BLE001
                continue

        if closed_any:
            time.sleep(1)

        # Ensure the main window has focus for subsequent menu navigation
        try:
            self._focus_window(main_window)
            time.sleep(0.5)
        except Exception:  # noqa: BLE001
            pass

    def _close_extra_windows(self, main_window, log_fn: Optional[LogFn]) -> None:
        """Close extra QB windows (reports, centers, forms) before export navigation.

        # FIX: Close extra QB windows before export navigation
        # Issue found 2026-05-08: Employee Center, reports, etc. intercept keyboard shortcuts
        # causing File→Utilities→Export to fail. Close all extra windows first.
        #
        # QuickBooks opens various child/tool windows within its MDI interface:
        #   - Employee Center, Customer Center, Vendor Center
        #   - Account Listing, Transaction Detail reports
        #   - Write Checks, Enter Bills, Pay Bills
        #   - Chart of Accounts, Item List, etc.
        # These windows receive keyboard focus and intercept Alt+F, Ctrl+key, and
        # other shortcuts meant for the main company window's menu bar.
        # By closing them first, we guarantee that menu navigation targets the
        # correct window.

        Strategy:
          1. Find all child windows of the main QB window
          2. Close any that match known report/center/form title patterns
          3. Keep only the main company window open
          4. Use safe close methods: click X button, send ESC, or close()
        """
        # Substrings (lower-cased) of child window titles that should be closed.
        # These are internal QB MDI child windows that steal keyboard focus.
        extra_window_hints = [
            "employee center",
            "customer center",
            "vendor center",
            "account listing",
            "transaction detail",
            "transaction list",
            "write checks",
            "enter bills",
            "pay bills",
            "chart of accounts",
            "item list",
            "sales tax",
            "payroll center",
            "report",
            "balance sheet",
            "profit & loss",
            "profit and loss",
            "trial balance",
            "a/r aging",
            "a/p aging",
            "general ledger",
            "journal",
            "register",
            "reconcile",
            "make deposits",
            "receive payments",
            "create invoices",
            "enter sales receipts",
            "credit memo",
            "purchase order",
            "sales order",
            "estimate",
            "statement",
            "home page",
            "quickbooks home",
        ]

        closed_count = 0

        # --- Approach 1: Close child windows of the main QB window ---
        try:
            children = main_window.children()
        except Exception:  # noqa: BLE001
            children = []

        for child in children:
            try:
                if not child.is_visible():
                    continue
                title = child.window_text() or ""
                if not title:
                    continue

                title_l = title.lower()

                # Skip the main company window itself (should not appear as child,
                # but guard against it)
                if child.handle == getattr(main_window, "handle", None):
                    continue

                if any(hint in title_l for hint in extra_window_hints):
                    self._emit(f"  [CloseExtra] Closing child window: '{title}'", log_fn)
                    # Try clicking Close/X button first
                    clicked = self._click_first_button(child, ["Close", "X", "Cancel"])
                    if not clicked:
                        # Try sending ESC to close
                        try:
                            child.set_focus()
                            time.sleep(0.1)
                            if send_keys is not None:
                                send_keys("{ESC}")
                                time.sleep(0.3)
                        except Exception:  # noqa: BLE001
                            pass
                        # Last resort: try .close() on the window
                        try:
                            child.close()
                        except Exception:  # noqa: BLE001
                            pass
                    closed_count += 1
                    time.sleep(0.3)
            except Exception:  # noqa: BLE001
                continue

        # --- Approach 2: Also check top-level desktop windows ---
        # Some QB windows (e.g., Employee Center) may appear as separate
        # top-level windows rather than MDI children.
        try:
            desktop = self._get_desktop()
            for win in desktop.windows():
                try:
                    if not win.is_visible():
                        continue
                    title = win.window_text() or ""
                    if not title:
                        continue

                    # Never close the main QB window
                    if win.handle == getattr(main_window, "handle", None):
                        continue
                    # Skip non-QB windows
                    if not re.search(self.QB_WINDOW_RE, title):
                        continue

                    title_l = title.lower()
                    if any(hint in title_l for hint in extra_window_hints):
                        self._emit(f"  [CloseExtra] Closing top-level QB window: '{title}'", log_fn)
                        clicked = self._click_first_button(win, ["Close", "X", "Cancel"])
                        if not clicked:
                            try:
                                win.set_focus()
                                time.sleep(0.1)
                                if send_keys is not None:
                                    send_keys("{ESC}")
                                    time.sleep(0.3)
                            except Exception:  # noqa: BLE001
                                pass
                            try:
                                win.close()
                            except Exception:  # noqa: BLE001
                                pass
                        closed_count += 1
                        time.sleep(0.3)
                except Exception:  # noqa: BLE001
                    continue
        except Exception:  # noqa: BLE001
            pass

        if closed_count > 0:
            self._emit(f"  [CloseExtra] Closed {closed_count} extra window(s).", log_fn)
            time.sleep(1)
        else:
            self._emit("  [CloseExtra] No extra windows found to close.", log_fn)

        # Restore focus to the main company window
        try:
            self._focus_window(main_window)
            time.sleep(0.5)
        except Exception:  # noqa: BLE001
            pass

    def _invoke_menu(self, window, menu_path: str, fallback_keys: Optional[str], log_fn: Optional[LogFn]) -> None:
        self._focus_window(window)
        try:
            window.menu_select(menu_path)
            return
        except Exception as exc:  # noqa: BLE001
            self._emit(f"menu_select failed for '{menu_path}': {exc}", log_fn)

        if fallback_keys:
            if send_keys is None:
                raise RuntimeError(f"Could not invoke menu '{menu_path}' and keyboard fallback unavailable.")
            self._emit(f"Trying fallback keystrokes for menu '{menu_path}': {fallback_keys}", log_fn)
            send_keys(fallback_keys, pause=0.08)
            time.sleep(1)
            return

        raise RuntimeError(f"Could not invoke menu path: {menu_path}")

    def _handle_standard_file_dialog(
        self,
        file_path: Path,
        save_mode: bool,
        timeout_s: int,
        log_fn: Optional[LogFn],
        parent_window=None,
    ) -> None:
        # FIX: Save As dialog is modal child of QB window, not desktop-level
        # Must search parent_window.children() like we do for Export Lists dialog
        # Bug found 2026-05-08: Tool selects lists correctly but can't find Save As dialog
        target = str(file_path)

        dialog_box: Dict[str, object] = {}

        def _cond() -> bool:
            dlg = self._find_active_dialog(
                title_re=r"(?i)(open|save\s+as|save\s+print|save document|save print output|import|export|create company)",
                parent_window=parent_window,
            )
            if dlg is not None:
                dialog_box["dlg"] = dlg
                return True
            return False

        if not self._wait_until(_cond, timeout_s, 0.5):
            raise RuntimeError("Timed out waiting for standard Open/Save dialog")

        dialog = dialog_box["dlg"]
        self._focus_window(dialog)

        # Try known File name field patterns first, then generic first edit control.
        edit_set = False
        for locator in [
            {"auto_id": "1148", "control_type": "Edit"},
            {"title_re": r"(?i)file\s*name", "control_type": "Edit"},
        ]:
            try:
                edit = dialog.child_window(**locator)
                if edit.exists(timeout=1):
                    edit.set_focus()
                    edit.set_edit_text(target)
                    edit_set = True
                    break
            except Exception:  # noqa: BLE001
                continue

        if not edit_set:
            edit_set = self._set_edit_value(dialog, target)

        if not edit_set:
            raise RuntimeError(f"Could not set file path in dialog: {target}")

        button_labels = ["Save", "Open", "Import", "OK", "Create", "Next"] if save_mode else ["Open", "Import", "OK"]
        if not self._click_first_button(dialog, button_labels):
            # Fallback: press Enter
            if send_keys is None:
                raise RuntimeError("Could not click Open/Save button and keyboard fallback unavailable")
            send_keys("{ENTER}")

        time.sleep(1)
        self._dismiss_common_dialogs(log_fn)

    def _handle_password_prompt(self, password: str, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        """Wait for password dialog, pause for manual entry, then resume automation.

        Manual password entry approach - automation pauses to allow user to
        enter credentials, then resumes once company opens.
        """
        _ = password  # Intentionally unused: password entry is now fully manual.

        wait_for_appear_s = min(timeout_s, 20)
        self._emit(
            f"  [Password] Waiting up to {wait_for_appear_s}s for password/login dialog to appear...",
            log_fn,
        )

        def _dialog_present() -> bool:
            return self._find_active_dialog(title_re=r"(?i)(password|login)") is not None

        if not self._wait_until(_dialog_present, wait_for_appear_s, 0.5):
            self._emit("  [Password] No password/login dialog detected; continuing.", log_fn)
            return

        self._emit(
            "PASSWORD DIALOG DETECTED - Waiting for user to enter password manually...",
            log_fn,
        )

        def _dialog_gone() -> bool:
            return self._find_active_dialog(title_re=r"(?i)(password|login)") is None

        if not self._wait_until(_dialog_gone, timeout_s, 0.5):
            raise RuntimeError(
                "Timed out waiting for manual password entry: password dialog is still open"
            )

        self._emit("  [Password] Password dialog closed. Resuming automation.", log_fn)

    def _wait_for_company_ready(
        self,
        main_window,
        timeout_s: int,
        log_fn: Optional[LogFn],
        company_hint: Optional[str] = None,
    ) -> None:
        """Wait for the company to finish loading by checking the window title.

        If *company_hint* is given (e.g. "Blank Template"), the title must
        contain that substring — which only appears AFTER the user has entered
        the password and the file has fully opened.  Without a hint we fall
        back to the old heuristic ("No Company Open" absent), which can
        false-positive when QB shows its product name before a company is
        actually loaded (e.g. while a password dialog is displayed).

        NOTE: This method does NOT dismiss dialogs during the wait loop.
        Popup/dialog dismissal should happen AFTER this method returns,
        ensuring the company is fully loaded first.
        """
        if company_hint:
            self._emit(
                f"[STEP 5] Waiting for '{company_hint}' to appear in title (timeout={timeout_s}s)...",
                log_fn,
            )
        else:
            self._emit(f"[STEP 5] Waiting for company to finish loading (timeout={timeout_s}s)...", log_fn)
        self._emit("  [Load] Polling window title every 1s for company name...", log_fn)

        poll_count = 0
        last_title = ""

        def _cond() -> bool:
            nonlocal poll_count, last_title
            poll_count += 1
            try:
                title = main_window.window_text() or ""
            except Exception as e:  # noqa: BLE001
                self._emit(f"  [Load] Poll #{poll_count}: ERROR reading title: {e}", log_fn)
                return False

            # Log title changes (avoid spamming identical titles)
            if title != last_title:
                self._emit(f"  [Load] Poll #{poll_count}: Title changed -> '{title}'", log_fn)
                last_title = title
            elif poll_count % 10 == 0:
                self._emit(f"  [Load] Poll #{poll_count}: Still waiting (title='{title}')...", log_fn)

            if company_hint:
                # Positive match: the file name must appear in the title
                # Normalize: strip non-ASCII, collapse whitespace
                import re as _re
                norm_title = _re.sub(r'\s+', ' ', title.strip()).lower()
                norm_hint = company_hint.strip().lower()
                found = norm_hint in norm_title
                if not found and poll_count <= 3:
                    self._emit(f"  [Load] DEBUG: norm_hint={repr(norm_hint)} norm_title={repr(norm_title)}", log_fn)
                    self._emit(f"  [Load] DEBUG: title hex={title.encode('utf-8', errors='replace').hex()}", log_fn)
            else:
                # Negative match: "No Company Open" must be absent
                found = "No Company Open" not in title

            if found:
                self._emit(f"  [Load] ✓ Company detected in title on poll #{poll_count}: '{title}'", log_fn)
            return found

        if not self._wait_until(_cond, timeout_s, 1.0):
            try:
                final_title = main_window.window_text() or ""
            except Exception:  # noqa: BLE001
                final_title = "<could not read>"
            self._emit(f"  [Load] TIMEOUT after {poll_count} polls. Final title: '{final_title}'", log_fn)
            raise RuntimeError("Company did not finish loading in QuickBooks")

        self._emit(f"[STEP 5] Company is fully loaded after {poll_count} poll(s).", log_fn)
        self._emit("  [Load] Pausing 2s for QB to stabilize after load...", log_fn)
        time.sleep(2)

    def _open_company_file(self, main_window, qbw_path: Path, password: str, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        """Open a QuickBooks company file using QB's command-line parameter.

        =====================================================================
        WHY WE USE COMMAND-LINE LAUNCH (2026-05-08 debugging session):
        =====================================================================
        Earlier versions tried to open company files via File → Open menu
        navigation. This was UNRELIABLE because:
          1. The File menu often didn't respond to Alt+F keystrokes when
             popup windows (Accountant Center, Learning Center, etc.) had
             stolen focus after QB launched.
          2. Navigating File → Open → file dialog → browse to .qbw was a
             multi-step process with many failure points.
          3. Each step required waiting for specific dialogs that might or
             might not appear depending on QB's state.

        The command-line approach is far simpler and more reliable:
          - Launch QB with: "QBW.exe" "path/to/file.qbw"
          - QB automatically opens the company file
          - QB automatically shows the password dialog
          - We just type the password and press Enter
          - Done!
        =====================================================================

        =====================================================================
        WHY POPUPS ARE DISMISSED AFTER COMPANY LOADS (not before):
        =====================================================================
        QB shows various popup windows (Accountant Center, Getting Started,
        Memorized Transactions, etc.) AFTER the company finishes loading.
        If we try to dismiss them too early:
          1. They haven't appeared yet, so we miss them
          2. Dismissing them while the company is still loading can cause
             focus issues that prevent the company from loading properly
          3. New popups can appear AFTER we've already dismissed earlier ones

        So the correct order is:
          1. Launch QB with .qbw → password dialog → type + Enter
          2. Wait for company to fully load (title bar shows company name)
          3. THEN dismiss all popups (they're all present now)
          4. THEN proceed with menu navigation (File → Export, etc.)
        =====================================================================

        Steps:
          1. Launch QB with the .qbw file as a parameter (auto-opens the file)
          2. Wait for the password dialog to appear
          3. Handle the password dialog (simplified: just type + Enter)
          4. Wait for the company to fully load
          5. Dismiss post-login popups
        """
        self._emit("=" * 60, log_fn)
        self._emit("  OPENING COMPANY FILE  –  Command-Line Parameter Approach", log_fn)
        self._emit("=" * 60, log_fn)
        self._emit(f"  File path : {qbw_path}", log_fn)
        self._emit(f"  Password  : {'(provided, length=%d)' % len(password) if password else '(none)'}", log_fn)
        self._emit(f"  Timeout   : {timeout_s}s", log_fn)

        if not qbw_path.exists():
            self._emit(f"  ERROR: File does not exist on disk: {qbw_path}", log_fn)
            raise FileNotFoundError(f"Source QBW file not found: {qbw_path}")
        self._emit(f"  File confirmed on disk ({qbw_path.stat().st_size:,} bytes).", log_fn)

        # ==================================================================
        # STEP 1 – Launch QB with the .qbw file as a command-line parameter
        # ==================================================================
        self._emit("[STEP 1/5] Launching QB with company file as command-line parameter...", log_fn)

        # Close the existing QB instance so we can re-launch with the file arg
        if self._qb2023_app is not None:
            self._emit("  [Launch] Closing existing QB instance to re-launch with file parameter...", log_fn)
            self._close_qb(self._qb2023_app, log_fn)
            self._qb2023_app = None
            time.sleep(2)

        # Launch QB with the .qbw path – QB will auto-open the company file
        qb_exe = self.config.install_paths.qb_2023_path
        self._qb2023_app = self._launch_qb(qb_exe, log_fn, qbw_path=qbw_path)

        self._emit("[STEP 1/5] ✓ QB launched with company file parameter.", log_fn)

        # Re-acquire the main window reference after relaunch
        self._emit("  [Launch] Waiting for QB main window to appear...", log_fn)
        main_window = self._find_qb_main_window(self._qb2023_app, "2023", self.config.timeouts.launch_qb_seconds)
        self._emit("  [Launch] ✓ Main window found.", log_fn)

        # ==================================================================
        # STEP 2-3 – Wait for password dialog → user enters password manually
        # ==================================================================
        self._emit("[STEP 2/5] Waiting for password/login dialog to appear...", log_fn)
        self._emit("  (QB should auto-prompt for the password after opening the company file.)", log_fn)
        self._handle_password_prompt(password, timeout_s=timeout_s, log_fn=log_fn)
        self._emit("[STEP 3/5] ✓ Manual password step complete.", log_fn)

        # ==================================================================
        # STEP 4 – Wait for the company to FULLY load
        # ==================================================================
        self._emit("[STEP 4/5] Waiting for company to fully load (title bar check)...", log_fn)
        self._emit("  (Looking for company name in title; 'No Company Open' must disappear.)", log_fn)
        self._wait_for_company_ready(main_window, timeout_s=timeout_s, log_fn=log_fn)
        self._emit("[STEP 4/5] ✓ Company is fully loaded.", log_fn)

        # ==================================================================
        # STEP 5 – Dismiss popup dialogs (Payroll, Accountant Center …)
        # ==================================================================
        self._emit("[STEP 5/5] Company loaded – dismissing post-login popup dialogs...", log_fn)
        self._emit("  [Popups] Pausing 1s for popups to finish rendering...", log_fn)
        time.sleep(1)

        self._emit("  [Popups] Running _dismiss_common_dialogs()...", log_fn)
        self._dismiss_common_dialogs(log_fn)
        self._emit("  [Popups] ✓ Common dialogs dismissed.", log_fn)

        self._emit("  [Popups] Running _close_popup_windows()...", log_fn)
        self._close_popup_windows(main_window, log_fn)
        self._emit("  [Popups] ✓ Popup windows closed.", log_fn)

        # Re-focus the main window after clearing popups
        self._emit("  [Focus] Re-focusing main QB window...", log_fn)
        self._focus_window(main_window)
        time.sleep(0.5)
        try:
            final_title = main_window.window_text() or ""
        except Exception:  # noqa: BLE001
            final_title = "<could not read>"
        self._emit(f"  [Focus] Main window title: '{final_title}'", log_fn)

        self._emit("=" * 60, log_fn)
        self._emit("  COMPANY FILE OPEN COMPLETE", log_fn)
        self._emit("=" * 60, log_fn)

    def _export_single_list_iif(
        self,
        main_window,
        list_name: str,
        out_path: Path,
        log_fn: Optional[LogFn],
        select_all: bool = False,
    ) -> None:
        mode = "all list panes" if select_all else f"list '{list_name}'"
        self._emit(f"Exporting {mode} -> {out_path}", log_fn)

        # Dismiss any stale menus / popups before starting
        # Note: pywinauto uses {ESC} not {ESCAPE}
        if send_keys is not None:
            send_keys("{ESC}")
            time.sleep(0.3)
            send_keys("{ESC}")
            time.sleep(0.3)

        # FIX: Close extra QB windows before export navigation
        # Issue found 2026-05-08: Employee Center, reports, etc. intercept keyboard shortcuts
        # causing File→Utilities→Export to fail. Close all extra windows first.
        self._close_extra_windows(main_window, log_fn)

        # Close any popup windows that may have stolen focus
        self._close_popup_windows(main_window, log_fn)
        self._dismiss_common_dialogs(log_fn)

        # Set focus back to main_window AFTER closing everything,
        # THEN attempt the export navigation
        self._focus_window(main_window)
        time.sleep(0.5)

        # Navigate: File -> Utilities -> Export -> Lists to IIF Files...
        # Try multiple approaches for menu navigation
        dlg = None

        # Approach 1: keyboard shortcuts with longer waits
        if send_keys is not None:
            self._emit("Trying menu navigation via keyboard shortcuts", log_fn)
            send_keys("%f")  # Alt+F for File menu
            time.sleep(1.0)
            send_keys("u")   # Utilities
            time.sleep(1.0)
            send_keys("e")   # Export
            time.sleep(1.0)
            send_keys("l")   # Lists to IIF Files
            time.sleep(2)

            # FIX: Search child dialogs of QB main window instead of desktop-level windows.
            # QB Export dialogs are modal children, not top-level, so Desktop().windows()
            # can't see them. This was causing the tool to loop ~20 times then timeout
            # (found in user testing 2026-05-08).
            dlg = self._find_active_dialog(title_re=r"(?i)(export|iif|list)", parent_window=main_window)

        # Approach 2: try menu_select via pywinauto
        if dlg is None:
            self._emit("Keyboard menu navigation failed, trying menu_select", log_fn)
            if send_keys is not None:
                send_keys("{ESC}")
                time.sleep(0.5)
            self._focus_window(main_window)
            time.sleep(0.5)
            try:
                self._invoke_menu(
                    main_window,
                    "File->Utilities->Export->Lists to IIF Files...",
                    None,
                    log_fn,
                )
                time.sleep(2)
                # Search child windows of main_window — export dialog is a modal child
                dlg = self._find_active_dialog(title_re=r"(?i)(export|iif|list)", parent_window=main_window)
            except Exception:  # noqa: BLE001
                pass

        # Approach 3: try alternative keyboard sequence
        if dlg is None and send_keys is not None:
            self._emit("Trying alternative keyboard sequence for export menu", log_fn)
            send_keys("{ESC}")
            time.sleep(0.5)
            self._focus_window(main_window)
            time.sleep(0.5)
            # Try Alt, then arrow keys through File menu
            send_keys("{ESC}")
            time.sleep(0.3)
            send_keys("%f")
            time.sleep(1.5)
            # Look for Utilities in the menu and use arrow keys
            send_keys("u")
            time.sleep(1.5)
            send_keys("e")
            time.sleep(1.5)
            send_keys("l")
            time.sleep(2)
            # Search child windows of main_window — export dialog is a modal child
            dlg = self._find_active_dialog(title_re=r"(?i)(export|iif|list)", parent_window=main_window)

        if dlg is None:
            raise RuntimeError("Export Lists to IIF dialog did not appear")

        self._focus_window(dlg)

        # ----------------------------------------------------------------
        # LIST SELECTION LOGIC
        # ----------------------------------------------------------------
        # QB's "Export Lists to IIF" dialog presents checkboxes for each
        # list type (Chart of Accounts, Customers, Vendors, Items,
        # Employees, etc.).  We MUST check the correct one AND uncheck
        # the others before clicking OK, otherwise QB shows:
        #   "Please select a file to export"
        #
        # Strategy (ordered by reliability):
        #   0. [QB 2023] Find Pane descendants — QB 2023 renders its
        #      checkboxes as custom Pane controls, NOT standard CheckBox
        #      or Button controls.  pywinauto sees them as type="Pane"
        #      even though they look like checkboxes on screen.
        #      Bug found 2026-05-08: Dialog shows checkboxes visually
        #      but exposes them as Pane type in the UI Automation tree.
        #   1. Find CheckBox descendants via pywinauto, match by text
        #   2. Try Button descendants (QB sometimes exposes checkboxes as
        #      Button controls with BS_CHECKBOX / BS_AUTOCHECKBOX style)
        #   3. Try List / Tree / ComboBox .select() API
        #   4. Keyboard fallback: Tab through controls and Space to toggle
        # ----------------------------------------------------------------

        selected = False

        # --- Attempt 0: Pane controls (QB 2023 custom checkbox rendering) ---
        # QB 2023 uses custom Pane controls instead of standard CheckBox
        # or Button controls for its Export Lists dialog.  The checkboxes
        # appear visually as checkboxes but are exposed to UI Automation
        # as generic Pane elements with the list name as window_text.
        # We click matching panes to toggle them on, and click non-matching
        # panes to toggle them off (ensuring only our target list is selected).
        try:
            panes = dlg.descendants(control_type="Pane")
            # Filter to panes that look like list-item checkboxes:
            # they have non-empty text and are NOT standard dialog buttons
            skip_texts = {"ok", "cancel", "help", "export", "save", ""}
            list_panes = []
            for pane in panes:
                try:
                    text = (pane.window_text() or "").strip()
                    if text.lower() not in skip_texts:
                        list_panes.append((pane, text))
                except Exception:  # noqa: BLE001
                    continue

            if list_panes:
                self._emit(f"  Found {len(list_panes)} Pane control(s) that look like list checkboxes", log_fn)
                if select_all:
                    for pane, text in list_panes:
                        self._emit(f"    Selecting pane: '{text}'", log_fn)
                        try:
                            pane.click_input()
                            selected = True
                        except Exception as exc:  # noqa: BLE001
                            self._emit(f"    Could not click pane '{text}': {exc}", log_fn)
                    if selected:
                        self._emit("  ✓ Selected all detectable pane-checkboxes", log_fn)
                else:
                    for pane, text in list_panes:
                        self._emit(f"    Pane: '{text}'", log_fn)
                        if list_name.lower() in text.lower():
                            # This is the list we want — click to check it
                            pane.click_input()
                            selected = True
                            self._emit(f"  ✓ Selected pane-checkbox: '{text}'", log_fn)
                        else:
                            # Uncheck other lists by clicking them off
                            try:
                                pane.click_input()
                                self._emit(f"    Toggled off pane: '{text}'", log_fn)
                            except Exception:  # noqa: BLE001
                                pass
        except Exception as exc:  # noqa: BLE001
            self._emit(f"  Pane search failed: {exc}", log_fn)

        # --- Attempt 1: CheckBox controls via pywinauto (older QB versions) ---
        if not selected:
            try:
                checkboxes = dlg.descendants(control_type="CheckBox")
                self._emit(f"  Found {len(checkboxes)} CheckBox control(s) in dialog", log_fn)
                for cb in checkboxes:
                    try:
                        cb_text = (cb.window_text() or "").strip()
                        self._emit(f"    CheckBox: '{cb_text}'", log_fn)
                        if select_all or list_name.lower() in cb_text.lower():
                            # Ensure selected for target list (or for all lists in select_all mode)
                            try:
                                if not cb.get_toggle_state():
                                    cb.toggle()
                            except Exception:  # noqa: BLE001
                                # toggle() failed — try clicking directly
                                cb.click_input()
                            selected = True
                            self._emit(f"  ✓ Selected checkbox: '{cb_text}'", log_fn)
                        elif not select_all:
                            # Uncheck any OTHER list so we only export the one we want
                            try:
                                if cb.get_toggle_state():
                                    cb.toggle()
                                    self._emit(f"    Unchecked: '{cb_text}'", log_fn)
                            except Exception:  # noqa: BLE001
                                pass
                    except Exception as exc:  # noqa: BLE001
                        self._emit(f"    (could not inspect checkbox: {exc})", log_fn)
                        continue
            except Exception as exc:  # noqa: BLE001
                self._emit(f"  CheckBox search failed: {exc}", log_fn)

        # --- Attempt 2: Button controls with checkbox style ---
        if not selected:
            try:
                buttons = dlg.descendants(control_type="Button")
                self._emit(f"  Trying {len(buttons)} Button control(s) as checkbox fallback", log_fn)
                for btn in buttons:
                    try:
                        btn_text = (btn.window_text() or "").strip()
                        if not btn_text:
                            continue
                        # Skip OK / Cancel / Export / etc.
                        if btn_text.lower() in ("ok", "cancel", "export", "save", "help"):
                            continue
                        self._emit(f"    Button: '{btn_text}'", log_fn)
                        if select_all or list_name.lower() in btn_text.lower():
                            btn.click_input()
                            selected = True
                            self._emit(f"  ✓ Clicked button-checkbox: '{btn_text}'", log_fn)
                    except Exception:  # noqa: BLE001
                        continue
            except Exception as exc:  # noqa: BLE001
                self._emit(f"  Button-checkbox search failed: {exc}", log_fn)

        # --- Attempt 3: List / Tree / ComboBox .select() ---
        if not selected:
            for ctrl_type in ["List", "ListView", "Tree", "TreeView", "ComboBox"]:
                try:
                    ctrls = dlg.descendants(control_type=ctrl_type)
                    for ctrl in ctrls:
                        try:
                            self._emit(f"  Trying {ctrl_type}.select('{list_name}')", log_fn)
                            ctrl.select(list_name)
                            selected = True
                            self._emit(f"  ✓ Selected via {ctrl_type}", log_fn)
                            break
                        except Exception:  # noqa: BLE001
                            continue
                    if selected:
                        break
                except Exception:  # noqa: BLE001
                    continue

        # --- Attempt 4: Keyboard fallback ---
        # Tab through all controls looking for one whose text matches,
        # then press Space to toggle the checkbox.
        if not selected and send_keys is not None:
            self._emit("  Trying keyboard Tab+Space fallback for list selection", log_fn)
            self._focus_window(dlg)
            time.sleep(0.3)
            # Press Home to move to top of the control list, then Tab through
            send_keys("{HOME}")
            time.sleep(0.2)
            for _attempt in range(20):
                send_keys("{TAB}")
                time.sleep(0.3)
                # Read the focused control's text
                try:
                    focused = dlg.get_focus()
                    focused_text = (focused.window_text() or "").strip() if focused else ""
                    self._emit(f"    Focused control: '{focused_text}'", log_fn)
                    if focused_text and list_name.lower() in focused_text.lower():
                        send_keys(" ")  # Space toggles a checkbox
                        selected = True
                        self._emit(f"  ✓ Selected via keyboard: '{focused_text}'", log_fn)
                        break
                except Exception:  # noqa: BLE001
                    continue

        # ----------------------------------------------------------------
        # GUARD: Do NOT click OK if we failed to select a list.
        # Clicking OK without a selection causes QB to show an error
        # dialog ("Please select a file to export") and wastes a retry.
        # ----------------------------------------------------------------
        if not selected:
            # Log all controls in the dialog for post-mortem debugging
            try:
                all_ctrls = dlg.descendants()
                self._emit(f"  DEBUG: Dialog has {len(all_ctrls)} total controls:", log_fn)
                for c in all_ctrls[:30]:  # cap at 30 to avoid log spam
                    try:
                        self._emit(f"    type={c.friendly_class_name()!r}  text={c.window_text()!r}", log_fn)
                    except Exception:  # noqa: BLE001
                        pass
            except Exception:  # noqa: BLE001
                pass
            mode_desc = "all list panes" if select_all else f"list '{list_name}'"
            raise RuntimeError(
                f"Could not select {mode_desc} in Export dialog. "
                "No matching checkbox, list item, or focusable control was found. "
                "See log above for dialog control details."
            )

        # Now safe to click OK — a list is selected
        if not self._click_first_button(dlg, ["OK", "Export", "Save"]):
            if send_keys is not None:
                send_keys("{ENTER}")

        time.sleep(1)
        self._handle_standard_file_dialog(out_path, save_mode=True, timeout_s=self.config.timeouts.export_seconds, log_fn=log_fn, parent_window=main_window)

        # ── Wait for and dismiss the export success confirmation dialog ──
        # After a successful IIF export, QuickBooks shows a modal dialog titled
        # "QuickBooks Desktop Information" (or a variant such as "QuickBooks
        # Information" / "Information") containing the message:
        #     "Your data has been exported successfully."
        # Bug found 2026-05-08: The tool exported the file correctly but never
        # dismissed this dialog, which blocked all subsequent menu interactions
        # because QB keeps the modal in front of the main window.
        #
        # Bug fixed 2026-05-08: Was passing arguments in wrong order (main_window
        # as title_re, list as parent_window) and passing a list instead of regex
        # string, plus an extra log_fn arg that caused a silent TypeError.
        # Also need to check child Static/Text controls for the actual message
        # text since window_text() on the dialog itself only returns the title.
        #
        # Strategy:
        #   1. Poll for up to 30 seconds (1 s intervals) for a dialog whose
        #      title matches one of the known success-dialog titles.
        #   2. Inspect the dialog's child Static/Text controls for the keywords
        #      "successfully" or "exported" to confirm it is indeed the success
        #      message (and not an unrelated error dialog with a similar title).
        #   3. Click the "OK" button to dismiss.  If the button click fails for
        #      any reason, fall back to sending {ENTER} which achieves the same
        #      result because OK is the default-focused button.
        # Wait for and dismiss success confirmation dialog
        log_fn("Waiting for export success confirmation dialog...")
        dialog_dismissed = False

        for i in range(30):
            try:
                # Try child window first
                success_dialog = self._find_active_dialog(
                    parent_window=main_window,
                    title_re="(QuickBooks|Information)"
                )

                # Fallback: try desktop level if not found as child
                if not success_dialog:
                    success_dialog = self._find_active_dialog(
                        parent_window=None,
                        title_re="(QuickBooks|Information)"
                    )

                if success_dialog:
                    dlg_title = success_dialog.window_text()
                    log_fn(f"Found potential success dialog: {dlg_title}")

                    # ── Determine if this is the success dialog ──
                    # Strategy 1: Check child Text controls for success keywords
                    confirmed_via_text = False
                    static_controls = success_dialog.descendants(control_type="Text")
                    for static in static_controls:
                        text = static.window_text().lower()
                        if "success" in text or "export" in text or "complete" in text:
                            log_fn(f"✅ SUCCESS CONFIRMED via text control: {text}")
                            confirmed_via_text = True
                            break

                    # Strategy 2: If no Text controls found (QB doesn't expose them
                    # via UIA for this dialog), confirm by title alone.
                    # "QuickBooks Desktop Information" is the known title for the
                    # export success dialog.  It only appears after a successful
                    # export, so the title match is sufficient.
                    if not confirmed_via_text and len(static_controls) == 0:
                        if re.search(r"(?i)QuickBooks.*Information", dlg_title):
                            log_fn(f"✅ SUCCESS CONFIRMED via dialog title (no UIA text controls exposed): {dlg_title}")
                            confirmed_via_text = True

                    if confirmed_via_text:
                        # ── Dismiss the dialog ──
                        # Try multiple approaches because QB's UIA tree is unreliable:
                        #   1. Click OK child (may be Button or Pane control type)
                        #   2. Send {ENTER} key (OK is default-focused)
                        #   3. Send {ESC} key
                        dismissed = False

                        # Approach 1a: Try OK as Button
                        try:
                            ok_btn = success_dialog.child_window(title="OK", control_type="Button")
                            ok_btn.click()
                            log_fn("Clicked OK button (Button control)")
                            dismissed = True
                        except Exception:
                            pass

                        # Approach 1b: Try OK as Pane (QB exposes it as Pane, not Button)
                        if not dismissed:
                            try:
                                ok_pane = success_dialog.child_window(title="OK", control_type="Pane")
                                ok_pane.click_input()
                                log_fn("Clicked OK button (Pane control via click_input)")
                                dismissed = True
                            except Exception:
                                pass

                        # Approach 1c: Try any descendant named OK regardless of type
                        if not dismissed:
                            try:
                                all_desc = success_dialog.descendants()
                                for desc in all_desc:
                                    if desc.window_text() == "OK":
                                        desc.click_input()
                                        log_fn(f"Clicked OK descendant ({desc.element_info.control_type})")
                                        dismissed = True
                                        break
                            except Exception:
                                pass

                        # Approach 2: Send ENTER key to the dialog
                        if not dismissed:
                            try:
                                success_dialog.set_focus()
                                if send_keys is not None:
                                    send_keys("{ENTER}")
                                    log_fn("Sent ENTER key to dismiss dialog")
                                    dismissed = True
                            except Exception:
                                pass

                        # Approach 3: Send ESC key
                        if not dismissed:
                            try:
                                success_dialog.type_keys("{ESC}")
                                log_fn("Sent ESC key to dismiss dialog")
                                dismissed = True
                            except Exception as esc_err:
                                log_fn(f"ESC key also failed: {esc_err}")

                        time.sleep(1)

                        # Verify dialog is actually gone
                        verify = self._find_active_dialog(
                            parent_window=main_window,
                            title_re="(QuickBooks|Information)"
                        )
                        if verify is None:
                            dialog_dismissed = True
                            log_fn("✅ Success dialog dismissed and verified gone")
                        else:
                            log_fn("⚠️ Dialog may still be present after dismiss attempt, retrying...")

                if dialog_dismissed:
                    break

            except Exception as e:
                log_fn(f"Error in success dialog detection (attempt {i+1}/30): {e}")

            time.sleep(1)

        if not dialog_dismissed:
            raise RuntimeError("Export success dialog was not dismissed - cannot continue to next list")

        log_fn("✅ Export completed and success dialog dismissed")

        self._dismiss_common_dialogs(log_fn)

    def _open_report(self, main_window, menu_path: str, fallback_keys: str, log_fn: Optional[LogFn]) -> None:
        self._invoke_menu(main_window, menu_path, fallback_keys, log_fn)
        time.sleep(2)
        self._dismiss_common_dialogs(log_fn)

    def _export_transaction_list_csv(self, main_window, out_csv: Path, log_fn: Optional[LogFn]) -> None:
        self._emit(f"Exporting Transaction List by Date report -> {out_csv}", log_fn)

        self._focus_window(main_window)
        time.sleep(0.5)

        # Navigate: Reports -> Accountant & Taxes -> Transaction List by Date
        if send_keys is not None:
            send_keys("%r")  # Alt+R for Reports menu
            time.sleep(0.5)
            send_keys("a")   # Accountant & Taxes
            time.sleep(0.5)
            send_keys("t")   # Transaction List by Date
            time.sleep(3)
        else:
            self._open_report(
                main_window,
                "Reports->Accountant & Taxes->Transaction List by Date",
                "%rat",
                log_fn,
            )

        # Wait for report window to appear and poll for "Building Report" dialog
        time.sleep(2)
        logger.info("Checking for 'Building Report' dialog...")
        building_dlg = None
        for _ in range(10):  # Give it 10 seconds to appear
            try:
                building_dlg = main_window.child_window(title_re=r"(?i)building.*report", control_type="Window")
                if building_dlg.exists(timeout=0.5):
                    logger.info("'Building Report' dialog found - waiting for report to complete...")
                    break
            except Exception:  # noqa: BLE001
                pass
            time.sleep(1)

        # If building dialog appeared, wait for it to finish (with timeout)
        if building_dlg and building_dlg.exists(timeout=0.5):
            max_wait = 600  # 10 minutes for large files
            start_time = time.time()
            while building_dlg.exists(timeout=1):
                elapsed = time.time() - start_time
                if elapsed > max_wait:
                    raise RuntimeError(f"Report generation exceeded {max_wait}s timeout")
                logger.info(f"Still building report... ({int(elapsed)}s elapsed)")
                time.sleep(3)
            logger.info("Report generation complete")
            time.sleep(2)  # Let the report window settle
        else:
            logger.info("No 'Building Report' dialog detected - report may have loaded instantly")
            time.sleep(3)

        self._dismiss_common_dialogs(log_fn)

        # Try to set date range to All
        if send_keys is not None:
            try:
                send_keys("%d")  # Focus date dropdown
                time.sleep(0.3)
                send_keys("all{ENTER}")
                time.sleep(2)
            except Exception:  # noqa: BLE001
                pass

        # Export report using File -> Export -> Excel/CSV
        if send_keys is None:
            raise RuntimeError("Keyboard automation unavailable for transaction CSV export")

        # Try Ctrl+E first (Excel export shortcut)
        send_keys("^e")
        time.sleep(2)

        # Verify the export dialog appeared
        export_dlg = self._find_active_dialog(title_re=r"(?i)(send report|excel|export|save)")
        if not export_dlg:
            # Try alternative: File -> Save As
            logger.warning("Ctrl+E did not open export dialog, trying File menu...")
            send_keys("%f")  # Alt+F
            time.sleep(1)
            send_keys("a")  # Save As
            time.sleep(1)
            export_dlg = self._find_active_dialog(title_re=r"(?i)(export|save)")
            if not export_dlg:
                raise RuntimeError("Export/Save As dialog did not appear after multiple attempts")
        logger.info("Export dialog found")

        export_dlg = self._find_active_dialog(title_re=r"(?i)(send report|excel|export|save)")
        if export_dlg is not None:
            self._focus_window(export_dlg)
            # Try selecting CSV option if present
            try:
                for txt in [
                    "comma separated values",
                    "Create a comma separated",
                    "CSV",
                ]:
                    try:
                        option = export_dlg.child_window(title_re=fr"(?i){re.escape(txt)}")
                        if option.exists(timeout=1):
                            option.click_input()
                            self._emit(f"Selected export option: {txt}", log_fn)
                            break
                    except Exception:  # noqa: BLE001
                        continue
            except Exception:  # noqa: BLE001
                pass

            self._click_first_button(export_dlg, ["Export", "OK", "Next", "Create"])
            time.sleep(1)

        self._handle_standard_file_dialog(out_csv, save_mode=True, timeout_s=self.config.timeouts.export_seconds, log_fn=log_fn, parent_window=main_window)
        self._dismiss_common_dialogs(log_fn)

        # Close the report window
        if send_keys is not None:
            send_keys("^{F4}")
            time.sleep(1)
            self._dismiss_common_dialogs(log_fn)

    def _select_pdf_printer_in_print_dialog(self, print_dlg, log_fn: Optional[LogFn]) -> bool:
        """Explicitly select 'Microsoft Print to PDF' in the QB Print dialog.

        QB stores its OWN per-dialog printer preference independent of the
        Windows default, so we must select the PDF printer on every print
        action.  Tries the ComboBox UIA control first, then falls back to
        keyboard navigation (Alt+P focuses Printer combo in QB Print Reports).

        Returns True if the printer was selected, False otherwise.
        """
        target_substring = "microsoft print to pdf"

        # Attempt 1: find ComboBox(es) and try .select() with several patterns
        try:
            for combo in print_dlg.descendants(control_type="ComboBox"):
                try:
                    items = combo.texts() if hasattr(combo, "texts") else []
                except Exception:  # noqa: BLE001
                    items = []
                # Try every dropdown option that contains our substring
                candidates = [t for t in items if t and target_substring in t.lower()]
                for candidate in candidates:
                    try:
                        combo.select(candidate)
                        self._emit(f"  Selected printer: {candidate}", log_fn)
                        return True
                    except Exception:  # noqa: BLE001
                        continue
                # Try generic select call as last resort on this combo
                try:
                    combo.select("Microsoft Print to PDF")
                    self._emit("  Selected printer: Microsoft Print to PDF (combo.select)", log_fn)
                    return True
                except Exception:  # noqa: BLE001
                    pass
        except Exception as exc:  # noqa: BLE001
            self._emit(f"  ComboBox enumeration failed: {exc}", log_fn)

        # Attempt 2: keyboard fallback — Alt+P focuses Printer combo in
        # QB's Print Reports dialog.  Type the printer name; Windows
        # combo boxes will auto-complete to the first match.
        if send_keys is not None:
            try:
                self._focus_window(print_dlg)
                send_keys("%p")  # Alt+P -> Printer combo
                time.sleep(0.3)
                # Clear by pressing Home and selecting all (combo auto-selects)
                send_keys("Microsoft Print to PDF")
                time.sleep(0.3)
                self._emit("  Selected printer via keyboard fallback", log_fn)
                return True
            except Exception as exc:  # noqa: BLE001
                self._emit(f"  Keyboard printer-select fallback failed: {exc}", log_fn)

        return False

    def _export_report_pdf(self, main_window, report_menu_path: str, fallback_keys: str, out_pdf: Path, log_fn: Optional[LogFn]) -> None:
        self._emit(f"Generating validation report: {out_pdf.name}", log_fn)
        self._open_report(main_window, report_menu_path, fallback_keys, log_fn)

        if send_keys is None:
            raise RuntimeError("Keyboard automation unavailable for PDF report export")

        # ---------------------------------------------------------------
        # Use QuickBooks' built-in "File -> Save As PDF..." — this skips
        # the printer dialog ENTIRELY (no AnyDesk/printer-selection
        # nonsense) and goes straight to a standard Save As dialog,
        # silent and deterministic.
        # ---------------------------------------------------------------
        # Strategy: Use Ctrl+P with Microsoft Print to PDF as the system default
        # printer. This is the most reliable approach because:
        #   - Ctrl+P is a universal QB shortcut, works regardless of menu state
        #   - The File menu order shifts when reports are open (Down*N is brittle)
        #   - menu_select() and descendants() both fail/hang on QB's UIA tree
        # The MS Print to PDF printer prompts a "Save Print Output As" dialog
        # which our _handle_standard_file_dialog handles.
        self._focus_window(main_window)
        time.sleep(0.5)
        send_keys("^p")
        time.sleep(1.5)
        print_dlg = self._find_active_dialog(title_re=r"(?i)(print|form name|reports)")
        if print_dlg is not None:
            self._focus_window(print_dlg)
            self._select_pdf_printer_in_print_dialog(print_dlg, log_fn)
            time.sleep(0.5)
            if not self._click_first_button(print_dlg, ["Print", "OK"]):
                send_keys("{ENTER}")
            self._emit("  Invoked Print -> Microsoft Print to PDF", log_fn)
        else:
            self._emit("  WARNING: Print dialog not detected after Ctrl+P", log_fn)
            send_keys("{ENTER}")  # Try anyway

        self._handle_standard_file_dialog(out_pdf, save_mode=True, timeout_s=self.config.timeouts.report_seconds, log_fn=log_fn, parent_window=main_window)
        self._dismiss_common_dialogs(log_fn)

    # -------- QB 2023 extraction --------

    def _launch_qb(self, exe_path: str, log_fn: Optional[LogFn], qbw_path: Optional[Path] = None) -> Optional[object]:
        """Launch QuickBooks, optionally with a company file as a command-line parameter.

        If *qbw_path* is provided and QB is not already running, QB is started
        with the .qbw file path as an argument.  This causes QB to open the
        company file directly and auto-prompt for the password – no File → Open
        menu navigation required.

        If QB is already running we simply connect to the existing process
        (the caller is responsible for opening the file via _open_company_file).
        """
        if qbw_path:
            self._emit(f"Launching QuickBooks: {exe_path} WITH company file: {qbw_path}", log_fn)
        else:
            self._emit(f"Launching QuickBooks: {exe_path}", log_fn)

        if self.config.dry_run:
            time.sleep(1)
            return None

        self._ensure_automation_ready()

        exe_name = Path(exe_path).name

        # When launching with a specific company file, KILL any existing QB
        # instance first. This prevents the "Secondary" window problem where
        # QB opens the new file as a secondary window while the old file
        # remains the primary — QBFC then connects to the wrong company.
        #
        # 2026-05-11: Skip the graceful menu-close for stale instances.
        # The old _close_qb path gets stuck on password dialogs from cached
        # company files, wasting 90+ seconds before force-killing anyway.
        # Just force-kill immediately — we don't care about data in the
        # stale instance (we're about to open a fresh template).
        if qbw_path:
            import subprocess
            self._emit(f"  Force-killing any existing QB instance before clean launch...", log_fn)
            for img in (exe_name, "QBW32.EXE", "QBW.EXE",
                        "QBW32PremierAccountant.exe", "QBWPremierAccountant.exe",
                        "qbw32.exe", "qbw.exe"):
                subprocess.run(
                    ["taskkill", "/F", "/IM", img],
                    capture_output=True, timeout=10
                )
            time.sleep(5)  # give Windows time to release file locks

            cmd_line = f'"{exe_path}" "{qbw_path}"'
            self._emit(f"  Starting QB with command: {cmd_line}", log_fn)
            app = Application(backend="uia").start(cmd_line)
            return app

        # No qbw_path — connect to existing or launch fresh
        try:
            app = Application(backend="uia").connect(path=exe_path)
            self._emit(f"Connected to already-running QuickBooks: {exe_name}", log_fn)
            return app
        except Exception:  # noqa: BLE001
            pass

        app = Application(backend="uia").start(exe_path)
        return app

    def _set_accounting_preferences_via_ui(
        self,
        qb_app,
        prefs: Dict[str, Any],
        log_fn: Optional[LogFn] = None,
    ) -> bool:
        """Enable 'Use account numbers' and other accounting preferences via UI.

        QBFC PreferencesModRq is unsupported in QB 2021.
        Approach: pure keyboard navigation (most reliable across QB versions).
          Edit → Preferences → Accounting (already selected, it's first) →
          Company Preferences tab → check boxes → OK.
        """
        from pywinauto.keyboard import send_keys as _sk

        acct_prefs = (prefs or {}).get("accounting") or {}
        use_acct_numbers = str(acct_prefs.get("is_using_account_numbers", "")).lower() in ("true", "1", "yes")
        use_class_tracking = str(acct_prefs.get("is_using_class_tracking", "")).lower() in ("true", "1", "yes")

        if not use_acct_numbers and not use_class_tracking:
            self._emit("  AcctPrefsUI: nothing to set (both disabled in source)", log_fn)
            return True

        self._emit(f"  AcctPrefsUI: need account_numbers={use_acct_numbers}, class_tracking={use_class_tracking}", log_fn)

        try:
            # Ensure QB window is visible and focused.
            # The watchdog uses SW_HIDE which persists even after the watchdog
            # pauses — we MUST use ShowWindow(SW_SHOW) to make it visible again.
            main = None
            try:
                main = self._find_qb_main_window(qb_app, "2021", 30,
                                                  include_hidden=True, log_fn=log_fn)
                try:
                    import ctypes
                    hwnd = main.handle
                    SW_SHOW = 5
                    SW_RESTORE = 9
                    ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                    ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
                    self._emit("  AcctPrefsUI: force-showed QB window via ShowWindow", log_fn)
                except Exception as e:
                    self._emit(f"  AcctPrefsUI: ShowWindow fallback: {e}", log_fn)
                    main.restore()
                main.set_focus()
            except Exception:
                try:
                    qb_app.top_window().set_focus()
                except Exception:
                    pass
            time.sleep(1)

            # ── Dismiss popups/dialogs BEFORE trying Edit → Preferences ──
            # The Enterprise upgrade popup ("Get the latest QuickBooks Desktop
            # Enterprise") steals focus and intercepts keystrokes.
            self._emit("  AcctPrefsUI: dismissing popups before opening Preferences...", log_fn)
            try:
                self._dismiss_common_dialogs(log_fn)
            except Exception:
                pass
            time.sleep(0.5)
            try:
                if main:
                    self._close_popup_windows(main, log_fn)
            except Exception:
                pass
            time.sleep(0.5)

            # Close any remaining "Enterprise" / "upgrade" / "Get the latest" windows
            desktop = self._get_desktop()
            for win in desktop.windows():
                try:
                    t = (win.window_text() or "").lower()
                    if not win.is_visible():
                        continue
                    if any(kw in t for kw in ('enterprise', 'upgrade', 'get the latest',
                                               'update', 'new feature', 'what\'s new')):
                        self._emit(f"  AcctPrefsUI: closing popup '{win.window_text()}'", log_fn)
                        try:
                            win.close()
                        except Exception:
                            try:
                                _sk("{ESC}")
                            except Exception:
                                pass
                        time.sleep(0.5)
                except Exception:
                    continue

            # Re-focus the main window after popup dismissal
            try:
                main.set_focus()
            except Exception:
                try:
                    qb_app.top_window().set_focus()
                except Exception:
                    pass
            time.sleep(0.5)

            # Open Edit → Preferences via keyboard — with retry loop
            # The Enterprise upgrade popup can steal focus at any moment,
            # so we retry up to 3 times: dismiss popups → try menu → check.
            prefs_found = False
            for attempt in range(3):
                self._emit(f"  AcctPrefsUI: attempt {attempt+1}/3 to open Preferences...", log_fn)

                # Dismiss any popup that appeared between attempts
                desktop = self._get_desktop()
                for win in desktop.windows():
                    try:
                        t = (win.window_text() or "").lower()
                        if not win.is_visible():
                            continue
                        if any(kw in t for kw in ('enterprise', 'upgrade', 'get the latest',
                                                   'update', 'new feature', 'what\'s new',
                                                   'usage', 'analytics', 'study', 'faq')):
                            self._emit(f"  AcctPrefsUI: closing popup '{win.window_text()}'", log_fn)
                            clicked = self._click_first_button(
                                win, ["Continue", "OK", "Close", "No", "Skip", "Later", "Cancel"]
                            )
                            if not clicked:
                                try: win.close()
                                except Exception: pass
                            time.sleep(0.5)
                    except Exception:
                        continue

                # Re-focus main window
                try:
                    import ctypes
                    ctypes.windll.user32.SetForegroundWindow(main.handle)
                except Exception:
                    pass
                try:
                    main.set_focus()
                except Exception:
                    try: qb_app.top_window().set_focus()
                    except Exception: pass
                time.sleep(1)

                # --- Method 1: pywinauto menu_select ---
                try:
                    main.menu_select("Edit->Preferences")
                    self._emit("  AcctPrefsUI: menu_select('Edit->Preferences') succeeded", log_fn)
                    time.sleep(3)
                except Exception as me:
                    self._emit(f"  AcctPrefsUI: menu_select failed: {me}", log_fn)

                    # --- Method 2: keyboard Alt+E → r ---
                    _sk("{ESC}")
                    time.sleep(0.3)
                    _sk("%e")
                    time.sleep(1.0)

                    # Check if a popup intercepted
                    desktop = self._get_desktop()
                    intercepted = False
                    for win in desktop.windows():
                        try:
                            t = (win.window_text() or "").lower()
                            if not win.is_visible():
                                continue
                            if any(kw in t for kw in ('enterprise', 'upgrade', 'get the latest')):
                                self._emit(f"  AcctPrefsUI: popup intercepted: '{win.window_text()}'", log_fn)
                                clicked = self._click_first_button(
                                    win, ["Continue", "OK", "Close", "No", "Skip", "Later", "Cancel"]
                                )
                                if not clicked:
                                    try: win.close()
                                    except Exception: pass
                                intercepted = True
                                time.sleep(0.5)
                        except Exception:
                            continue

                    if intercepted:
                        _sk("{ESC}")
                        time.sleep(0.5)
                        continue

                    # Try 'r' then 'p' then arrow-key navigation
                    _sk("r")
                    time.sleep(2)

                # Verify Preferences dialog opened
                desktop = self._get_desktop()
                for win in desktop.windows():
                    try:
                        t = (win.window_text() or "").lower()
                        if "preferences" in t and win.is_visible():
                            prefs_found = True
                            break
                    except Exception:
                        continue

                if prefs_found:
                    break

                # --- Method 3: arrow-key navigation to last item ---
                self._emit("  AcctPrefsUI: trying Edit menu arrow-key navigation...", log_fn)
                _sk("{ESC}")
                time.sleep(0.5)
                _sk("%e")
                time.sleep(1.0)
                # Preferences is usually the last item in Edit menu
                _sk("{END}")
                time.sleep(0.3)
                _sk("{ENTER}")
                time.sleep(3)

                # Check again
                desktop = self._get_desktop()
                for win in desktop.windows():
                    try:
                        t = (win.window_text() or "").lower()
                        if "preferences" in t and win.is_visible():
                            prefs_found = True
                            break
                    except Exception:
                        continue
                if prefs_found:
                    break

                _sk("{ESC}")
                time.sleep(0.5)

            self._emit(f"  AcctPrefsUI: Preferences dialog {'found' if prefs_found else 'NOT FOUND'}", log_fn)

            # "Accounting" is the FIRST category (already selected by default).
            # Click "Company Preferences" tab — it's a tab control.
            # The tab order is: My Preferences | Company Preferences
            # We need to click Company Preferences. Use Ctrl+Tab or click.
            # In QB Preferences, the tabs respond to mouse clicks.
            # Use pywinauto to find the dialog and its tabs.
            time.sleep(1)

            # Try to find and click "Company Preferences" tab
            for win in desktop.windows():
                try:
                    t = (win.window_text() or "").lower()
                    if "preferences" not in t or not win.is_visible():
                        continue
                    # Found the Preferences dialog — click Company Preferences tab
                    rect = win.rectangle()
                    self._emit(f"  AcctPrefsUI: Preferences dialog at ({rect.left},{rect.top})-({rect.right},{rect.bottom})", log_fn)

                    # Company Preferences tab is typically in the right half of the tab strip
                    # Tab strip is near the top of the content area
                    from pywinauto import mouse as _mouse
                    tab_y = rect.top + 100  # tabs are about 100px from top
                    tab_x = rect.left + int((rect.right - rect.left) * 0.65)  # right-ish
                    _mouse.click(coords=(tab_x, tab_y))
                    time.sleep(1)
                    self._emit(f"  AcctPrefsUI: clicked Company Preferences tab at ({tab_x},{tab_y})", log_fn)

                    # Now find "Use account numbers" checkbox
                    # It's typically at specific coordinates within the dialog.
                    # The checkbox area is in the main content pane.
                    # Strategy: use pywinauto to find checkboxes, or use coordinates.

                    # Try pywinauto child_window first (without requiring uia backend)
                    checked_acct = False
                    checked_class = False
                    try:
                        children = win.children()
                        for child in children:
                            try:
                                ct = child.window_text() or ""
                                if not ct:
                                    continue
                                ct_l = ct.lower()
                                if use_acct_numbers and "account number" in ct_l:
                                    try:
                                        state = child.get_toggle_state()
                                        if state == 0:
                                            child.click_input()
                                            checked_acct = True
                                            self._emit(f"  AcctPrefsUI: ✓ checked '{ct}'", log_fn)
                                    except Exception:
                                        child.click_input()
                                        checked_acct = True
                                        self._emit(f"  AcctPrefsUI: ✓ clicked '{ct}'", log_fn)
                                    time.sleep(0.3)
                                elif use_class_tracking and "class track" in ct_l:
                                    try:
                                        state = child.get_toggle_state()
                                        if state == 0:
                                            child.click_input()
                                            checked_class = True
                                            self._emit(f"  AcctPrefsUI: ✓ checked '{ct}'", log_fn)
                                    except Exception:
                                        child.click_input()
                                        checked_class = True
                                        self._emit(f"  AcctPrefsUI: ✓ clicked '{ct}'", log_fn)
                                    time.sleep(0.3)
                            except Exception:
                                continue
                    except Exception:
                        pass

                    if not checked_acct and use_acct_numbers:
                        # Coordinate fallback: "Use account numbers" is typically
                        # around 40% from left, 35% from top in the content area
                        cb_x = rect.left + int((rect.right - rect.left) * 0.12)
                        cb_y = rect.top + int((rect.bottom - rect.top) * 0.35)
                        _mouse.click(coords=(cb_x, cb_y))
                        self._emit(f"  AcctPrefsUI: clicked account numbers at ({cb_x},{cb_y})", log_fn)
                        time.sleep(0.3)

                    if not checked_class and use_class_tracking:
                        cb_x = rect.left + int((rect.right - rect.left) * 0.12)
                        cb_y = rect.top + int((rect.bottom - rect.top) * 0.55)
                        _mouse.click(coords=(cb_x, cb_y))
                        self._emit(f"  AcctPrefsUI: clicked class tracking at ({cb_x},{cb_y})", log_fn)
                        time.sleep(0.3)

                    # Click OK to save
                    _sk("{ENTER}")
                    time.sleep(2)
                    self._emit("  AcctPrefsUI: ✓ Preferences saved", log_fn)
                    return True

                except Exception as exc:
                    self._emit(f"  AcctPrefsUI: dialog handling error: {exc}", log_fn)
                    continue

            self._emit("  AcctPrefsUI: FAILED — could not find Preferences dialog", log_fn)
            _sk("{ESC}")
            time.sleep(0.5)
            return False

        except Exception as exc:
            self._emit(f"  AcctPrefsUI: FAILED — {exc}", log_fn)
            try:
                _sk("{ESC}")
                time.sleep(0.5)
                _sk("{ESC}")
            except Exception:
                pass
            return False

    def _set_company_info_via_ui(self, qb_app, info: Dict[str, Any], log_fn: Optional[LogFn] = None) -> bool:
        """Set company profile via Company → My Company (UI automation).

        QBFC has no CompanyMod method, so we navigate the QB 2021 UI:
          1. Company menu → My Company
          2. Tab through edit fields and type values
          3. OK to save

        Returns True on success.
        """
        if not info:
            self._emit("  CompanyUI: no company info, skipping", log_fn)
            return False

        # Sanitise all values to ASCII – Windows console encoding chokes on
        # characters like \u2192 (→) that can appear in QB company names.
        def _ascii_safe(v):
            if isinstance(v, str):
                return v.encode('ascii', 'replace').decode('ascii')
            return v
        info = {k: _ascii_safe(v) for k, v in info.items()}

        company_name = info.get("company_name", "")
        if not company_name:
            self._emit("  CompanyUI: no company_name in snapshot, skipping", log_fn)
            return False

        try:
            main_win = self._find_qb_main_window(qb_app, "2021", timeout_s=15,
                                                  include_hidden=True)
            # Force-show the window (watchdog's SW_HIDE persists after pause)
            try:
                import ctypes
                hwnd = main_win.handle
                ctypes.windll.user32.ShowWindow(hwnd, 5)   # SW_SHOW
                ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                self._emit("  CompanyUI: force-showed QB window", log_fn)
            except Exception:
                main_win.restore()
            main_win.set_focus()
            time.sleep(0.5)
            self._emit("  CompanyUI: Opening Company -> My Company...", log_fn)

            # Navigate: Company menu → My Company
            try:
                main_win.menu_select("Company->My Company")
                self._emit("  CompanyUI: menu_select succeeded", log_fn)
            except Exception:
                self._emit("  CompanyUI: menu_select failed, trying keyboard...", log_fn)
                if send_keys:
                    # Try Alt+C for Company menu (some QB versions)
                    send_keys("%c")
                    time.sleep(0.8)
                    send_keys("m")   # 'M' = My Company
                    time.sleep(0.5)
                    # If that didn't work, try Alt+P
                    desktop_check = self._get_desktop()
                    found_co = False
                    for w in desktop_check.windows():
                        t = (w.window_text() or "").lower()
                        if ("company information" in t or "my company" in t) and w.is_visible():
                            found_co = True
                            break
                    if not found_co:
                        send_keys("{ESC}")
                        time.sleep(0.3)
                        send_keys("%p")
                        time.sleep(0.8)
                        send_keys("m")
                        time.sleep(0.5)

            time.sleep(2)

            # Find the Company Information dialog
            dialog = self._find_active_dialog(
                title_re=r"(?i)(company\s+information|my\s+company)",
                parent_window=main_win,
            )
            if dialog is None:
                # Try desktop level
                from pywinauto import Desktop
                for w in Desktop(backend="uia").windows():
                    t = w.window_text() or ""
                    if "company information" in t.lower() or "my company" in t.lower():
                        dialog = w
                        break

            if dialog is None:
                self._emit("  CompanyUI: Company Information dialog not found", log_fn)
                return False

            self._emit("  CompanyUI: Found Company Information dialog", log_fn)
            dialog.set_focus()
            time.sleep(0.5)

            # The Company Information dialog has these fields (typical order):
            #   Company Name, Legal Name, Address, City, State, Zip, Country,
            #   Phone, Fax, Email, Website, Legal Address, Legal City, ...
            #   EIN, SSN, Tax Form
            # We use Tab to move between fields and Ctrl+A to select existing text.

            edits = dialog.descendants(control_type="Edit")
            self._emit(f"  CompanyUI: Found {len(edits)} edit fields", log_fn)

            # Build a mapping of field values to set.
            # We'll log what we find and set company name at minimum.
            addr = info.get("address") or {}
            legal_addr = info.get("legal_address") or {}

            # Strategy: Set focus to the first edit (Company Name), then
            # Tab through all fields setting values in order.
            # The exact field order depends on QB version, so we'll set
            # the first field (Company Name) directly, then use Tab for the rest.

            field_sequence = [
                ("Company Name", company_name),
                ("Legal Name", info.get("legal_name", "")),
                # Address block
                ("Address Line 1", addr.get("addr1", "")),
                ("Address Line 2", addr.get("addr2", "")),
                ("Address Line 3", addr.get("addr3", "")),
                ("City", addr.get("city", "")),
                ("State", addr.get("state", "")),
                ("Zip", addr.get("postalcode", "")),
                ("Country", addr.get("country", "")),
                # Contact
                ("Phone", info.get("phone", "")),
                ("Fax", info.get("fax", "")),
                ("Email", info.get("email", "")),
                ("Website", info.get("website", "")),
                # Legal address
                ("Legal Address Line 1", legal_addr.get("addr1", "")),
                ("Legal City", legal_addr.get("city", "")),
                ("Legal State", legal_addr.get("state", "")),
                ("Legal Zip", legal_addr.get("postalcode", "")),
            ]

            # Set each edit field by index
            set_count = 0
            for idx, (label, value) in enumerate(field_sequence):
                if idx >= len(edits):
                    break
                if not value:
                    # Skip empty values but still Tab past the field
                    continue
                # Sanitize non-ASCII characters (e.g. → arrow, ® symbols)
                # to prevent encoding errors in type_keys/set_edit_text
                value = ''.join(c if ord(c) < 128 else ' ' for c in str(value))
                if not value.strip():
                    continue
                try:
                    edit = edits[idx]
                    edit.set_focus()
                    time.sleep(0.1)
                    if send_keys:
                        send_keys("^a")  # Select all
                        time.sleep(0.05)
                    try:
                        edit.set_edit_text(value)
                    except Exception:
                        if send_keys:
                            send_keys("{DELETE}")
                            time.sleep(0.05)
                            edit.type_keys(value, with_spaces=True, pause=0.02)
                    set_count += 1
                    self._emit(f"    Set {label} = '{value}'", log_fn)
                except Exception as exc:
                    self._emit(f"    Could not set {label}: {exc}", log_fn)

            # Click OK to save
            time.sleep(0.3)
            ok_clicked = self._click_first_button(dialog, ["OK", "Save", "Close"])
            if not ok_clicked and send_keys:
                send_keys("{ENTER}")
            time.sleep(1)

            self._emit(f"  CompanyUI: Set {set_count} fields for '{company_name}'", log_fn)
            return set_count > 0

        except Exception as exc:
            self._emit(f"  CompanyUI: _set_company_info_via_ui failed: {exc}", log_fn)
            return False

    def _close_qb(self, app: Optional[object], log_fn: Optional[LogFn]) -> None:
        """Close QuickBooks using menu navigation (File → Close Company, then
        File → Exit) so QB flushes all in-memory data to the .qbw file.

        Previous approach (taskkill, even without /F) did NOT trigger QB's
        internal save routine, resulting in empty/template-only files.
        """
        self._emit("Closing QuickBooks via menu navigation...", log_fn)
        if self.config.dry_run:
            return

        import subprocess
        images = ("QBW32.EXE", "QBW.EXE", "qbw32.exe", "qbw.exe")

        # -----------------------------------------------------------
        # PHASE 1 — Use keyboard shortcuts to close company + exit
        # This triggers QB's internal save/flush for the company file.
        # Alt+F4 sometimes works, but File → Close Company is the
        # safest path because it explicitly flushes the company data
        # before proceeding to close the app.
        # -----------------------------------------------------------
        menu_close_done = False

        # Try to find the QB window and use it.
        # CRITICAL: pass include_hidden=True because the watchdog may
        # have hidden the QB window.  After finding it, restore it so
        # keyboard shortcuts (Ctrl+W, Alt+F4) actually reach it.
        try:
            main_window = self._find_qb_main_window(app, None, timeout_s=10, include_hidden=True)
            if main_window is not None:
                try:
                    # Un-hide the window if the watchdog hid it (SW_RESTORE)
                    hwnd = main_window.handle
                    if hwnd:
                        import win32gui   # type: ignore[import-untyped]
                        import win32con   # type: ignore[import-untyped]
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        time.sleep(0.3)
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.3)
                    main_window.set_focus()
                    time.sleep(0.5)
                    self._emit("  QB window found and restored to foreground.", log_fn)
                except Exception:  # noqa: BLE001
                    pass

                # --- Step 1: File → Close Company ---
                self._emit("  Sending Ctrl+W (Close Company)...", log_fn)
                try:
                    if send_keys is not None:
                        send_keys("^w")          # Ctrl+W = Close Company
                        time.sleep(3)             # give QB time to flush
                        self._emit("  Ctrl+W sent. Waiting for flush...", log_fn)

                        # QB may pop a "Save changes?" dialog — click Yes / press Enter
                        time.sleep(2)
                        send_keys("{ENTER}")
                        time.sleep(3)
                        self._emit("  Save dialog handled (if any).", log_fn)
                except Exception as exc:  # noqa: BLE001
                    self._emit(f"  WARN: Ctrl+W failed: {exc}", log_fn)

                # --- Step 2: File → Exit (Alt+F4) ---
                self._emit("  Sending Alt+F4 (Exit QuickBooks)...", log_fn)
                try:
                    if send_keys is not None:
                        send_keys("%{F4}")        # Alt+F4 = Exit
                        time.sleep(2)
                        # Handle any additional "are you sure?" dialog
                        send_keys("{ENTER}")
                        time.sleep(2)
                        self._emit("  Alt+F4 sent.", log_fn)
                        menu_close_done = True
                except Exception as exc:  # noqa: BLE001
                    self._emit(f"  WARN: Alt+F4 failed: {exc}", log_fn)
            else:
                self._emit("  Could not find QB window for menu close.", log_fn)
        except Exception as exc:  # noqa: BLE001
            self._emit(f"  WARN: menu-based close failed: {exc}", log_fn)

        # -----------------------------------------------------------
        # PHASE 2 — Wait for QB to exit gracefully (up to 90s)
        # QB 2021 is slow — it checks for updates during shutdown.
        # IMPORTANT: QB may pop a password/login dialog during shutdown
        # (e.g. for a cached company file). If we don't dismiss it,
        # QB hangs forever and we time out + force-kill.
        # -----------------------------------------------------------
        if menu_close_done:
            self._emit("  Waiting up to 90s for QB to exit...", log_fn)
            # Passwords to try if a login dialog appears during close
            close_passwords = [
                self.config.install_paths.qb_2021_template_password,
                "3825You171",
                "Fl0640098!@!",
            ]
            # De-duplicate while preserving order
            seen = set()
            close_passwords = [p for p in close_passwords if not (p in seen or seen.add(p))]
            close_pw_idx = 0
            deadline = time.time() + 90
            while time.time() < deadline:
                # Check if QB exited (check BOTH QB 2021 = QBW32.EXE and QB 2023 = qbw.exe)
                try:
                    still_running = False
                    for proc_name in ("QBW32.EXE", "qbw.exe"):
                        result = subprocess.run(
                            ["tasklist", "/FI", f"IMAGENAME eq {proc_name}"],
                            capture_output=True, timeout=5, text=True, check=False,
                        )
                        if proc_name.upper() in (result.stdout or "").upper():
                            still_running = True
                            break
                    if not still_running:
                        self._emit("  QuickBooks exited cleanly!", log_fn)
                        time.sleep(3)  # let Windows release file locks
                        return  # SUCCESS — no force kill needed
                except Exception:  # noqa: BLE001
                    pass

                # Check for password/login dialog blocking the close
                login_dlg = self._find_active_dialog(title_re=r"(?i)(password|login)")
                if login_dlg is not None and send_keys is not None:
                    pw = close_passwords[close_pw_idx]
                    self._emit(f"  Password dialog blocking close — entering password (attempt {close_pw_idx + 1})", log_fn)
                    # Pause watchdog so it doesn't interfere
                    if self._watchdog is not None:
                        self._watchdog.pause()
                    try:
                        time.sleep(0.5)
                        login_dlg.set_focus()
                        time.sleep(0.3)
                        # Click the Edit field
                        try:
                            edits = [c for c in login_dlg.children()
                                     if c.friendly_class_name() == "Edit"]
                            if edits:
                                edits[0].click_input()
                                time.sleep(0.2)
                        except Exception:
                            pass
                        # Clear any existing text before typing
                        send_keys("^a", pause=0.02)
                        time.sleep(0.1)
                        # Escape special chars for send_keys
                        safe_pw = pw
                        for ch in ('{', '}'):
                            safe_pw = safe_pw.replace(ch, '{' + ch + '}')
                        for ch in ('+', '^', '%', '(', ')', '~'):
                            safe_pw = safe_pw.replace(ch, '{' + ch + '}')
                        # Log what we're sending
                        mask = pw[0:3] + '*' * max(0, len(pw) - 6) + pw[-3:] if len(pw) > 6 else '***'
                        self._emit(f"  [PW DEBUG close] raw='{mask}' escaped='{safe_pw}' len={len(pw)}", log_fn)
                        send_keys(safe_pw, pause=0.02)
                        time.sleep(0.3)
                        send_keys("{ENTER}")
                        time.sleep(3)
                        # Check for wrong-password warning
                        warning_dlg = self._find_active_dialog(title_re=r"(?i)(warning|error|incorrect)")
                        if warning_dlg is not None:
                            self._emit(f"  Password {close_pw_idx + 1} incorrect during close, trying next", log_fn)
                            self._click_first_button(warning_dlg, ["OK", "Close"])
                            if close_pw_idx + 1 < len(close_passwords):
                                close_pw_idx += 1
                            time.sleep(1)
                    except Exception as exc:
                        self._emit(f"  WARN: password entry during close failed: {exc}", log_fn)
                    finally:
                        if self._watchdog is not None:
                            self._watchdog.resume()

                time.sleep(3)
            self._emit("  QB still running after 90s — will force-kill.", log_fn)

        # -----------------------------------------------------------
        # PHASE 3 — Force-kill only as absolute last resort
        # If we got here, menu close failed or QB is stuck. Data should
        # already be flushed if Ctrl+W succeeded even partially.
        # -----------------------------------------------------------
        self._emit("  Force-killing QB processes...", log_fn)
        try:
            for image in images:
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/IM", image, "/T"],
                        capture_output=True, timeout=10, check=False,
                    )
                except Exception:  # noqa: BLE001
                    pass
        except Exception:  # noqa: BLE001
            pass

        # Best-effort pywinauto kill as a backup
        if app is not None:
            try:
                app.kill()
            except Exception:  # noqa: BLE001
                pass

        # Give Windows a moment to release file locks
        time.sleep(3)

    def _is_company_already_open(self, main_window, company_name_hint: str) -> bool:
        """Check if the desired company is already open in QB.

        The window title alone is not reliable -- QB may show
        'QuickBooks Accountant Desktop Plus 2023' even when the
        'No Company Open' landing page is displayed.  We therefore
        also inspect child elements for the 'No Company Open' text.
        """
        try:
            title = main_window.window_text() or ""

            # Explicit "No Company Open" in the title bar -> not open
            if "No Company Open" in title:
                return False

            # Check child elements for "No Company Open" text which
            # appears in the QB landing page content area
            try:
                children = main_window.descendants(control_type="Text")
                for child in children:
                    try:
                        child_text = child.window_text() or ""
                        if "No Company Open" in child_text:
                            return False
                    except Exception:  # noqa: BLE001
                        continue
            except Exception:  # noqa: BLE001
                pass

            # If the title contains the company name hint, it is open
            if company_name_hint.lower() in title.lower():
                return True

            # QB typically puts the company file name in the title bar
            # when a company is open.  A generic title like
            # "QuickBooks Accountant Desktop Plus 2023" without any
            # company-specific text means nothing is loaded yet.
            # Only return True if the title has something beyond the
            # standard QB product name.
            generic_patterns = [
                r"(?i)^QuickBooks.*Desktop.*\d{4}$",
                r"(?i)^QuickBooks.*Premier.*\d{4}$",
                r"(?i)^QuickBooks.*Enterprise.*\d{4}$",
            ]
            for pat in generic_patterns:
                if re.match(pat, title.strip()):
                    return False

            # If we get here the title has extra text (likely a company
            # name) and no "No Company Open" was found -> assume open
            if "QuickBooks" in title:
                return True

        except Exception:  # noqa: BLE001
            pass
        return False

    def _handle_startup_dialogs(self, password: str, timeout_s: int, log_fn: Optional[LogFn],
                                alt_password: Optional[str] = None,
                                min_wait_s: int = 90) -> None:
        """Handle dialogs that appear when QB starts up.

        =====================================================================
        WHY THIS USES THE SIMPLE "type + Enter" APPROACH (2026-05-08):
        =====================================================================
        Same rationale as _handle_password_prompt() — trying to detect dialog
        types, enumerate edit fields, and click specific buttons was unreliable
        with QB's non-standard controls. Instead, we just:
          1. Detect the password dialog exists (by title)
          2. Wait a moment for it to render
          3. Type the password and press Enter
        See _handle_password_prompt() docstring for full explanation.

        2026-05-11: Added watchdog pause/resume around password entry to
        prevent the watchdog from dismissing "wrong password" warnings
        before the main thread can detect them.  Also added alt_password
        parameter — if the primary password fails, we retry with the
        alternate password before giving up.

        2026-05-11 (fix): Added min_wait_s parameter (default 30s).
        QB can show non-blocking dialogs (Update Service, promo popups)
        BEFORE the password dialog appears. Without a minimum wait, the
        loop would see no blocking dialogs and break early — before the
        password dialog ever appeared. Then _find_qb_main_window would
        time out because nobody entered the password.
        =====================================================================
        """
        passwords_to_try = [password]
        if alt_password and alt_password != password:
            passwords_to_try.append(alt_password)
        password_attempt_idx = 0
        password_entered = False

        # PAUSE the watchdog for the ENTIRE startup-dialog phase.
        # Critical: the watchdog hides QB main windows (SW_HIDE).
        # In Windows, hiding a parent window also hides its owned
        # child windows — including the password dialog!  If the
        # watchdog hides QB's main frame before the password dialog
        # appears, _find_active_dialog will never see it (it only
        # searches visible windows).  Pausing the watchdog keeps
        # the main window visible so the password dialog is
        # discoverable.  We resume once the password is entered
        # (or we give up).
        if self._watchdog is not None:
            self._watchdog.pause()

        start = time.time()
        while time.time() - start < timeout_s:
            # Check for password/login dialog first.
            # Strategy: Use win32gui.FindWindow to look for the exact QB
            # login dialog title. This is MORE RELIABLE than
            # Desktop().windows() or EnumWindows + pywinauto wrapping,
            # because QB 2021's 32-bit dialogs are often invisible to
            # pywinauto's UIA backend.
            login_dlg = None
            login_hwnd = None
            try:
                import win32gui
                # FindWindow scans ALL top-level windows — class=None means any class
                hwnd = win32gui.FindWindow(None, "QuickBooks Desktop Login")
                if hwnd and win32gui.IsWindowVisible(hwnd):
                    login_hwnd = hwnd
                    # Try to wrap with pywinauto for .children()/.click_input()
                    try:
                        from pywinauto.controls.hwndwrapper import HwndWrapper
                        login_dlg = HwndWrapper(hwnd)
                    except Exception:
                        pass
                    if login_dlg is None:
                        self._emit(f"  Found login hwnd={hwnd} but could not wrap with pywinauto", log_fn)
            except Exception as e:
                self._emit(f"  win32gui.FindWindow error: {e}", log_fn)

            # Fallback: pywinauto desktop search
            if login_dlg is None and login_hwnd is None:
                login_dlg = self._find_active_dialog(title_re=r"(?i)(password|login)")

            if (login_dlg is not None or login_hwnd is not None) and not password_entered:
                current_pw = passwords_to_try[password_attempt_idx]
                self._emit(f"Found startup login dialog, entering password (attempt {password_attempt_idx + 1}/{len(passwords_to_try)})", log_fn)

                # Focus the dialog, click the password field, then type.
                if send_keys is not None:
                    time.sleep(1)  # Let dialog fully render

                    # Bring dialog to foreground — try pywinauto first,
                    # then fall back to win32gui.
                    try:
                        if login_dlg is not None:
                            login_dlg.set_focus()
                        elif login_hwnd:
                            import win32gui
                            win32gui.SetForegroundWindow(login_hwnd)
                    except Exception:  # noqa: BLE001
                        pass
                    time.sleep(0.3)

                    # Try to click the password Edit field so the cursor
                    # is definitely there (QB may default focus elsewhere).
                    try:
                        if login_dlg is not None:
                            edits = [c for c in login_dlg.children()
                                     if c.friendly_class_name() == "Edit"]
                            if edits:
                                edits[0].click_input()
                                time.sleep(0.2)
                    except Exception:  # noqa: BLE001
                        pass  # Fallback: just type and hope for the best

                    # Clear any existing text in the field (previous failed
                    # attempt may have left partial text) before typing.
                    send_keys("^a", pause=0.02)  # Ctrl+A = select all
                    time.sleep(0.1)

                    # Escape pywinauto special chars: + ^ % { } ( ) ~
                    # so passwords with these chars are typed literally.
                    safe_pw = current_pw
                    for ch in ('{', '}'):  # must escape braces FIRST
                        safe_pw = safe_pw.replace(ch, '{' + ch + '}')
                    for ch in ('+', '^', '%', '(', ')', '~'):
                        safe_pw = safe_pw.replace(ch, '{' + ch + '}')
                    # Log what we're actually sending (mask middle chars for security)
                    mask = current_pw[0:3] + '*' * max(0, len(current_pw) - 6) + current_pw[-3:] if len(current_pw) > 6 else '***'
                    self._emit(f"  [PW DEBUG] raw='{mask}' escaped='{safe_pw}' len={len(current_pw)}", log_fn)
                    send_keys(safe_pw, pause=0.02)
                    time.sleep(0.3)
                    send_keys("{ENTER}")

                    password_entered = True
                    self._startup_password_handled = True
                    self._emit("Password entered at startup (type + Enter)", log_fn)
                    time.sleep(5)  # Wait for QB to process login

                    # Check if a "wrong password" warning appeared.
                    # Watchdog is PAUSED so the warning is still visible.
                    warning_dlg = self._find_active_dialog(title_re=r"(?i)(warning|error|incorrect)")
                    if warning_dlg is not None:
                        self._emit(f"Password attempt {password_attempt_idx + 1} was incorrect, dismissing warning", log_fn)
                        self._click_first_button(warning_dlg, ["OK", "Close"])
                        password_entered = False
                        self._startup_password_handled = False
                        time.sleep(1)
                        # Move to next password if available
                        if password_attempt_idx + 1 < len(passwords_to_try):
                            password_attempt_idx += 1
                            self._emit(f"Will try alternate password next...", log_fn)
                        else:
                            self._emit("All passwords exhausted — will keep retrying last one", log_fn)

                    # (watchdog stays paused — will be resumed at end of method)
                    continue
                else:
                    self._emit("WARNING: send_keys unavailable, cannot enter password", log_fn)

            elif (login_dlg is not None or login_hwnd is not None) and password_entered:
                # Password was already entered but dialog is still showing - wait
                time.sleep(2)
                continue

            # Dismiss other common startup dialogs
            self._dismiss_common_dialogs(log_fn)
            time.sleep(1)

            # Periodic progress log so we know the wait is alive
            elapsed_now = time.time() - start
            if int(elapsed_now) % 10 == 0 and int(elapsed_now) > 0:
                self._emit(f"  [Startup] waiting for password dialog... {int(elapsed_now)}s elapsed", log_fn)

            # ---------------------------------------------------------------
            # BLIND PASSWORD ENTRY FALLBACK (after 90s of failing to detect)
            # QB 2021's login dialog can take 60-80s to appear (splash screen,
            # "Updating QuickBooks..." etc). The regular FindWindow detection
            # should catch it given enough time. Only fall to blind entry as
            # a last resort after 90s.
            # ---------------------------------------------------------------
            if not password_entered and elapsed_now >= 90 and send_keys is not None:
                current_pw = passwords_to_try[password_attempt_idx]
                self._emit(f"  [BLIND ENTRY] Dialog not found after {int(elapsed_now)}s — typing password blind (attempt {password_attempt_idx + 1})", log_fn)

                # Focus the LOGIN DIALOG specifically, not the main QB window.
                # The login dialog title is "QuickBooks Desktop Login".
                # If we can't find it, fall back to any QB window.
                try:
                    import win32gui, win32con  # type: ignore[import-untyped]
                    login_hwnd = None
                    fallback_hwnd = None
                    def _enum_focus(hwnd, _):
                        nonlocal login_hwnd, fallback_hwnd
                        t = win32gui.GetWindowText(hwnd)
                        if not t:
                            return True
                        tl = t.lower()
                        if "timewarp" in tl:
                            return True  # skip our own GUI
                        if "login" in tl and "quickbooks" in tl:
                            login_hwnd = hwnd
                            return False  # found it — stop
                        if "quickbooks" in tl and fallback_hwnd is None:
                            fallback_hwnd = hwnd
                        return True
                    win32gui.EnumWindows(_enum_focus, None)
                    target_hwnd = login_hwnd or fallback_hwnd
                    if target_hwnd:
                        target_title = win32gui.GetWindowText(target_hwnd)
                        self._emit(f"  [BLIND ENTRY] Focusing hwnd={target_hwnd}: '{target_title}'", log_fn)
                        win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
                        time.sleep(0.3)
                        win32gui.SetForegroundWindow(target_hwnd)
                        time.sleep(0.5)
                    else:
                        self._emit("  [BLIND ENTRY] WARNING: Could not find any QB window to focus!", log_fn)
                except Exception as focus_err:
                    self._emit(f"  [BLIND ENTRY] WARNING: Focus attempt failed: {focus_err}", log_fn)

                safe_pw = current_pw
                for ch in ('{', '}'):
                    safe_pw = safe_pw.replace(ch, '{' + ch + '}')
                for ch in ('+', '^', '%', '(', ')', '~'):
                    safe_pw = safe_pw.replace(ch, '{' + ch + '}')
                mask = current_pw[0:3] + '*' * max(0, len(current_pw) - 6) + current_pw[-3:] if len(current_pw) > 6 else '***'
                self._emit(f"  [PW DEBUG] raw='{mask}' escaped='{safe_pw}' len={len(current_pw)}", log_fn)
                # Clear any existing text and type password directly.
                # Do NOT send Tab — the password field should already have
                # focus in the login dialog. Tab would move to OK button.
                send_keys("^a", pause=0.02)  # select all (clear stale text)
                time.sleep(0.1)
                send_keys(safe_pw, pause=0.02)
                time.sleep(0.3)
                send_keys("{ENTER}")
                password_entered = True
                self._startup_password_handled = True
                self._emit("  [BLIND ENTRY] Password + Enter sent", log_fn)
                time.sleep(8)  # Wait for QB to process login

                # Check for wrong-password warning
                warning_dlg = self._find_active_dialog(title_re=r"(?i)(warning|error|incorrect)")
                if warning_dlg is not None:
                    self._emit(f"  [BLIND ENTRY] Wrong password detected, dismissing", log_fn)
                    self._click_first_button(warning_dlg, ["OK", "Close"])
                    password_entered = False
                time.sleep(8)  # Wait for QB to process login

                # Check for wrong-password warning
                warning_dlg = self._find_active_dialog(title_re=r"(?i)(warning|error|incorrect)")
                if warning_dlg is not None:
                    self._emit(f"  [BLIND ENTRY] Wrong password detected, dismissing", log_fn)
                    self._click_first_button(warning_dlg, ["OK", "Close"])
                    password_entered = False
                    self._startup_password_handled = False
                    time.sleep(1)
                    if password_attempt_idx + 1 < len(passwords_to_try):
                        password_attempt_idx += 1
                continue

            # Check if we're past all startup dialogs
            desktop = self._get_desktop()
            blocking_dialogs = False
            for w in desktop.windows():
                try:
                    if not w.is_visible():
                        continue
                    t = (w.window_text() or "").lower()
                    if any(x in t for x in ["login", "password", "external accountant", "memorized transaction"]):
                        blocking_dialogs = True
                        break
                except Exception:  # noqa: BLE001
                    continue

            if not blocking_dialogs:
                elapsed = time.time() - start
                if password_entered or elapsed >= min_wait_s:
                    # Safe to exit: either we already entered the password,
                    # or we've waited long enough for the password dialog to
                    # appear (it never did — file may not be password-protected).
                    if not password_entered and elapsed >= min_wait_s:
                        self._emit(f"  No password dialog after {int(elapsed)}s — continuing (file may not be protected)", log_fn)
                    break
                # Haven't entered password yet and haven't waited min_wait_s —
                # keep polling in case the password dialog hasn't appeared yet.
                # (QB may show other dialogs like Update Service first.)

        # RESUME the watchdog — startup dialog handling is done.
        # The watchdog will now hide QB main windows as usual.
        if self._watchdog is not None:
            self._watchdog.resume()

    def _export_from_qb2023(self, job: CompanyJob, export_dir: Path, log_fn: Optional[LogFn]) -> Dict[str, Path]:
        """Exports list IIF, transaction CSV, and report PDFs from QB 2023.

        Strategy (2026-05-09): TRY QBFC SDK FIRST. If QBFC is registered and
        QB 2023 has the company file open, we extract everything via the SDK
        — no menus, no dialogs, no send_keys. Only fall back to the legacy UI
        automation path if QBFC is unavailable or fails to connect.

        In dry-run mode, this method synthesizes/copies sample files.
        """

        export_dir.mkdir(parents=True, exist_ok=True)

        lists_iif = export_dir / "all_lists.IIF"
        tx_csv = export_dir / "TransactionList_QB2023.CSV"

        # ------------------------------------------------------------------
        # QBFC SDK path (preferred)
        # ------------------------------------------------------------------
        # QBFC requires the company to be FULLY OPEN before we can connect.
        # Handle password + wait for company to load BEFORE attempting QBFC.
        # ------------------------------------------------------------------
        if not self.config.dry_run:
            if self._qb2023_app is None:
                raise RuntimeError("QB 2023 app instance is not initialized")

            # Auto-type the source file password.
            # For Tax Man Mike build: all files share the same password,
            # so we auto-enter it via _handle_startup_dialogs().
            # For public/retail build: job.password comes from the GUI
            # password field (user enters it before clicking Start).
            source_password = job.password or ""
            if source_password:
                self._emit("Auto-entering source file password in QB 2023...", log_fn)
                self._handle_startup_dialogs(
                    password=source_password,
                    timeout_s=120,
                    log_fn=log_fn,
                )
            else:
                self._emit("", log_fn)
                self._emit("=" * 60, log_fn)
                self._emit("ACTION REQUIRED: Enter the admin password in QuickBooks 2023.", log_fn)
                self._emit("  The tool will auto-detect when the company is loaded.", log_fn)
                self._emit("=" * 60, log_fn)
                self._emit("", log_fn)

            try:
                # Watchdog hides QB main windows — use include_hidden
                # so we can still find and poll the title bar.
                main_window = self._find_qb_main_window(
                    self._qb2023_app, "2023",
                    self.config.timeouts.launch_qb_seconds,
                    include_hidden=True,
                )
                # Give up to 5 minutes for company to finish loading
                self._wait_for_company_ready(main_window, timeout_s=300, log_fn=log_fn)
                self._dismiss_common_dialogs(log_fn)
                self._close_popup_windows(main_window, log_fn)
                self._emit("QB 2023 company is loaded. Proceeding to QBFC export.", log_fn)
            except Exception as exc:  # noqa: BLE001
                import traceback as _tb
                self._emit(f"WARN: pre-QBFC company-load wait failed: {exc}", log_fn)
                self._emit(_tb.format_exc(), log_fn)

            try:
                from qbfc_export import export_company_via_qbfc  # local import to avoid mandatory dep

                self._emit("=== Attempting QBFC SDK export ===", log_fn)
                self._emit(f"  Target company file: {job.qbw_path}", log_fn)
                results = export_company_via_qbfc(qbw_path=job.qbw_path, export_dir=export_dir, log_fn=log_fn)
                self._emit("=== QBFC SDK export succeeded ===", log_fn)
                return results
            except Exception as exc:  # noqa: BLE001
                import traceback as _tb
                self._emit(f"QBFC export failed: {exc}", log_fn)
                self._emit(_tb.format_exc(), log_fn)
                # FAIL HARD — do not fall back to legacy IIF UI automation.
                # IIF export loses too much data (no entity refs, no proper
                # AP/AR linkage, no rich fields).  If QBFC fails we want the
                # job to fail loudly so we can debug it.
                raise RuntimeError(f"QBFC SDK export failed and legacy IIF fallback is disabled: {exc}")

        if self.config.dry_run:
            sample_root = Path.home() / "Uploads" / "QBDowngrade_Results"
            if (sample_root / "all_lists.IIF").exists():
                shutil.copy2(sample_root / "all_lists.IIF", lists_iif)
            else:
                lists_iif.write_text("!HDR\nHDR\tDRY-RUN\n", encoding="utf-8")

            if (sample_root / "TransactionList_QB2023.CSV").exists():
                shutil.copy2(sample_root / "TransactionList_QB2023.CSV", tx_csv)
            else:
                tx_csv.write_text("Type,Date,Account,Split,Debit,Credit\n", encoding="utf-8")

            # copy report artifacts if available
            report_files = {}
            for f in [
                "TrialBalance_QB2023.pdf",
                "BalanceSheet_QB2023.pdf",
                "ProfitLoss_QB2023.pdf",
                "AR_Aging_QB2023.pdf",
                "AP_Aging_QB2023.pdf",
            ]:
                src = sample_root / f
                if src.exists():
                    dst = export_dir / f
                    shutil.copy2(src, dst)
                    report_files[f] = dst

            return {"lists_iif": lists_iif, "tx_csv": tx_csv, **report_files}


    def _create_qb2021_company(self, job: CompanyJob, target_dir: Path, log_fn: Optional[LogFn]) -> Path:
        target_dir.mkdir(parents=True, exist_ok=True)
        # Replace trailing version marker (e.g., "23" -> "21") but only at word boundaries
        raw_name = job.target_company_name or job.qbw_path.stem
        target_name = re.sub(r"\b23\b", "21", raw_name) if "23" in raw_name else raw_name
        target_qbw = target_dir / f"{target_name}.qbw"

        if self.config.dry_run:
            target_qbw.write_text("DRY RUN PLACEHOLDER - QBW FILE CREATED", encoding="utf-8")
            self._emit(f"Created dry-run target company file: {target_qbw}", log_fn)
            return target_qbw

        if self._qb2021_app is None:
            raise RuntimeError("QB 2021 app instance is not initialized")

        main_window = self._find_qb_main_window(self._qb2021_app, "2021", self.config.timeouts.launch_qb_seconds)
        self._dismiss_common_dialogs(log_fn)

        self._emit("Creating new QB 2021 company", log_fn)
        self._invoke_menu(main_window, "File->New Company", "%fn", log_fn)
        time.sleep(1)

        wizard = self._find_active_dialog(title_re=r"(?i)(new company|quickbooks setup|easystep)")
        if wizard is None:
            raise RuntimeError("New Company wizard did not appear")

        self._focus_window(wizard)

        # Prefer skipping the long interview if available.
        self._click_first_button(wizard, ["Skip Interview", "Express Start", "Detailed Start", "Start Setup", "Next"])
        time.sleep(1)

        # Fill core fields where controls are available.
        field_values = [target_name, "000000000", "Services", "Accrual"]
        for idx, value in enumerate(field_values):
            self._set_edit_value(wizard, value, edit_index=idx)
            if send_keys is not None:
                send_keys("{TAB}")
                time.sleep(0.2)

        self._click_first_button(wizard, ["Next", "Continue", "Create Company", "Finish"])

        # Save file location.
        self._handle_standard_file_dialog(target_qbw, save_mode=True, timeout_s=self.config.timeouts.open_company_seconds, log_fn=log_fn, parent_window=main_window)

        self._wait_for_company_ready(main_window, timeout_s=self.config.timeouts.open_company_seconds, log_fn=log_fn)
        self._dismiss_common_dialogs(log_fn)
        return target_qbw

    def _import_single_iif(self, main_window, iif_path: Path, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        if not iif_path.exists():
            raise FileNotFoundError(f"IIF file does not exist: {iif_path}")

        self._emit(f"Importing IIF file: {iif_path.name}", log_fn)
        self._invoke_menu(main_window, "File->Utilities->Import->IIF Files...", "%fui", log_fn)
        time.sleep(1)

        self._handle_standard_file_dialog(iif_path, save_mode=False, timeout_s=timeout_s, log_fn=log_fn, parent_window=main_window)

        # Common QuickBooks import warnings/confirmations.
        start = time.time()
        while time.time() - start < timeout_s:
            dlg = self._find_active_dialog(title_re=r"(?i)(import|quickbooks|warning|confirmation|information)")
            if dlg is None:
                break
            if not self._click_first_button(dlg, ["Yes", "OK", "Continue", "Done", "Close"]):
                if send_keys is not None:
                    send_keys("{ENTER}")
            time.sleep(0.5)

        self._dismiss_common_dialogs(log_fn)

    def _import_iif_lists_qb2021(self, lists_iif: Path, log_fn: Optional[LogFn]) -> None:
        if self.config.dry_run:
            self._emit(f"Dry-run import lists: {lists_iif}", log_fn)
            return

        if self._qb2021_app is None:
            raise RuntimeError("QB 2021 app instance is not initialized")

        main_window = self._find_qb_main_window(self._qb2021_app, "2021", self.config.timeouts.launch_qb_seconds)

        if not lists_iif.exists():
            raise FileNotFoundError(f"Combined list IIF does not exist: {lists_iif}")

        self._with_retries(
            lambda: self._import_single_iif(
                main_window,
                lists_iif,
                timeout_s=self.config.timeouts.import_seconds,
                log_fn=log_fn,
            ),
            f"Import list IIF ({lists_iif.name})",
            log_fn,
        )

        self._emit("List IIF import completed", log_fn)

    def _import_iif_transactions_qb2021(self, tx_iif: Path, log_fn: Optional[LogFn]) -> None:
        if self.config.dry_run:
            self._emit(f"Dry-run import transactions: {tx_iif}", log_fn)
            return

        if self._qb2021_app is None:
            raise RuntimeError("QB 2021 app instance is not initialized")

        main_window = self._find_qb_main_window(self._qb2021_app, "2021", self.config.timeouts.launch_qb_seconds)

        self._with_retries(
            lambda: self._import_single_iif(
                main_window,
                tx_iif,
                timeout_s=max(self.config.timeouts.import_seconds, 600),
                log_fn=log_fn,
            ),
            "Import transaction IIF",
            log_fn,
        )

        self._emit("Transaction IIF import completed", log_fn)

    def _extract_validation_snapshot(self, exports_dir: Path) -> CompanyValidationSnapshot:
        """Extract real validation metrics by parsing exported IIF and CSV files.

        Parses all_lists.IIF for entity counts and TransactionList_QB2023.CSV
        for transaction count and financial totals (trial balance, AR, AP).
        """
        import csv as csv_mod

        lists_iif = exports_dir / "all_lists.IIF"
        tx_csv = exports_dir / "TransactionList_QB2023.CSV"
        tx_iif = exports_dir / "transactions_generated.IIF"

        # --- Entity counts from IIF ---
        account_count = 0
        customer_count = 0
        vendor_count = 0
        item_count = 0

        if lists_iif.exists():
            with open(lists_iif, "r", encoding="utf-8", errors="replace") as fh:
                for line in fh:
                    tag = line.split("\t", 1)[0] if "\t" in line else ""
                    if tag == "ACCNT":
                        account_count += 1
                    elif tag == "CUST":
                        customer_count += 1
                    elif tag == "VEND":
                        vendor_count += 1
                    elif tag == "INVITEM":
                        item_count += 1

        # --- Transaction count from generated IIF (most accurate) ---
        transaction_count = 0
        if tx_iif.exists():
            with open(tx_iif, "r", encoding="utf-8", errors="replace") as fh:
                for line in fh:
                    if line.startswith("TRNS\t"):
                        transaction_count += 1
        elif tx_csv.exists():
            # Fallback: count non-header data rows in CSV
            with open(tx_csv, "r", encoding="utf-8", errors="replace") as fh:
                reader = csv_mod.reader(fh)
                header = next(reader, None)
                for row in reader:
                    if row and len(row) > 1 and row[0]:  # skip blank/label rows
                        transaction_count += 1

        # --- Financial totals from CSV ---
        trial_balance_total = 0.0
        ar_total = 0.0
        ap_total = 0.0

        if tx_csv.exists():
            with open(tx_csv, "r", encoding="utf-8", errors="replace") as fh:
                reader = csv_mod.DictReader(fh)
                for row in reader:
                    try:
                        debit = float(row.get("Debit", "0").replace(",", "") or "0")
                        credit = float(row.get("Credit", "0").replace(",", "") or "0")
                    except (ValueError, AttributeError):
                        continue

                    trial_balance_total += debit

                    acct = (row.get("Account") or "").lower()
                    if "accounts receivable" in acct or "a/r" in acct:
                        ar_total += debit - credit
                    elif "accounts payable" in acct or "a/p" in acct:
                        ap_total += credit - debit

        return CompanyValidationSnapshot(
            trial_balance_total=round(trial_balance_total, 2),
            ar_total=round(ar_total, 2),
            ap_total=round(ap_total, 2),
            transaction_count=transaction_count,
            account_count=account_count,
            customer_count=customer_count,
            vendor_count=vendor_count,
            item_count=item_count,
        )

    def run_job(
        self,
        job: CompanyJob,
        log_fn: Optional[LogFn] = None,
        progress_fn: Optional[ProgressFn] = None,
    ) -> CompanyJobResult:
        steps = [
            "Launch QB 2023",
            "Export from QB 2023",
            "Close QB 2023",
            "Parse CSV transactions",
            "Copy QB 2021 template",
            "Launch QB 2021",
            "QBFC import into QB 2021",
            "Validation report generation",
            "Close QB 2021",
        ]

        def set_progress(step_idx: int) -> None:
            if progress_fn:
                progress_fn((step_idx / len(steps)) * 100.0)

        qb2023_app = None
        qb2021_app = None

        # Start the dialog watchdog: auto-dismisses QB popups and hides QB
        # windows so the operator sees only our GUI for the entire run.
        try:
            from dialog_watchdog import DialogWatchdog
            watchdog = DialogWatchdog(log_fn=lambda m: self._emit(m, log_fn), hide_qb=True)
            watchdog.start()
            self._watchdog = watchdog
        except Exception as _wd_exc:  # noqa: BLE001
            self._emit(f"[Watchdog] failed to start: {_wd_exc}", log_fn)
            watchdog = None
            self._watchdog = None

        try:
            # --- Directory layout ---
            #   working\source\   — staged copy of customer's original + QB-2021 Template
            #   working\Export\   — extracted data (snapshot, IIF, CSV)
            #   Final Output\<Company>\ — deliverables (converted .qbw, reports)
            working_root = Path(self.config.default_working_dir)
            source_dir   = working_root / "source"
            export_dir   = working_root / "Export"
            output_dir   = job.output_dir   # already points to Final Output\<Company>

            # Wipe stale working dirs so a previous failed run never pollutes.
            # source_dir MUST be cleaned too — a stale template .qbw from a
            # previous failed run may still be locked by a residual QB process.
            # Kill any QB processes first, wait for locks to release, then wipe.
            import subprocess as _sp
            for img in ("QBW32.exe", "QBW32PremierAccountant.exe",
                        "QBWPremierAccountant.exe", "qbupdate.exe",
                        "qbw.exe",   # QB 2023 process name
                        "QBDBMgrN.exe", "QBDBMgr.exe",
                        "QBCFMonitorService.exe"):
                _sp.run(["taskkill", "/F", "/IM", img], capture_output=True)
            # CRITICAL: Stop the QuickBooks Database Manager service.
            # The actual service name is "QuickBooksDB33" (not QBDBMgrN).
            # Use `sc stop` (doesn't require elevated PS) + taskkill as backup.
            try:
                _sp.run(["sc", "stop", "QuickBooksDB33"],
                        timeout=15, capture_output=True)
            except Exception:
                pass
            try:
                _sp.run(
                    ["powershell", "-NoProfile", "-Command",
                     "Stop-Service QuickBooksDB* -Force -ErrorAction SilentlyContinue"],
                    timeout=15, capture_output=True,
                )
            except Exception:
                pass
            time.sleep(5)  # Give Windows time to release file locks after service stop
            # Retry rmtree with backoff — Windows file locks can linger after taskkill
            for stale_dir in (source_dir, export_dir, output_dir):
                if stale_dir.exists():
                    self._emit(f"Cleaning stale directory: {stale_dir}", log_fn)
                    for attempt in range(5):
                        try:
                            shutil.rmtree(stale_dir)
                            break
                        except Exception as exc:  # noqa: BLE001
                            if attempt < 4:
                                time.sleep(2)  # wait for file locks to release
                            else:
                                self._emit(f"  WARN: could not remove {stale_dir} after 5 attempts: {exc}", log_fn)
            for d in (source_dir, export_dir, output_dir):
                d.mkdir(parents=True, exist_ok=True)

            # --- Phase 0: Stage Source ---
            # Copy customer's original .qbw (and sidecars) into working\source\
            # so we NEVER touch the original. All subsequent steps use the copy.
            self._emit("=== Phase 0 — Stage Source (safe copy) ===", log_fn)
            original_qbw = job.qbw_path
            staged_qbw = source_dir / original_qbw.name
            if not self.config.dry_run:
                # Delete destination first if it survived cleanup (locked file edge case)
                if staged_qbw.exists():
                    try:
                        staged_qbw.unlink()
                    except Exception:  # noqa: BLE001
                        pass  # copy2 will overwrite or fail with clear error
                shutil.copy2(str(original_qbw), str(staged_qbw))
                self._emit(f"  Staged: {original_qbw} -> {staged_qbw}", log_fn)
                for ext_s in (".qbw.ND", ".qbw.DSN", ".tlg", ".TLG"):
                    sidecar = original_qbw.parent / f"{original_qbw.stem}{ext_s}"
                    if sidecar.exists():
                        try:
                            shutil.copy2(str(sidecar), str(source_dir / sidecar.name))
                        except Exception:
                            pass
                # Repoint job to staged copy for the rest of the run
                job = CompanyJob(
                    qbw_path=staged_qbw,
                    password=job.password,
                    output_dir=job.output_dir,
                    target_company_name=job.target_company_name,
                )
                self._emit(f"  Original untouched at: {original_qbw}", log_fn)
            else:
                self._emit("  (dry run — skipping source staging)", log_fn)

            # 1) Launch QB 2023 with STAGED company file (never the original)
            set_progress(0)
            qb2023_app = self._launch_qb(self.config.install_paths.qb_2023_path, log_fn, qbw_path=job.qbw_path)
            self._qb2023_app = qb2023_app

            # 2) Export
            set_progress(1)
            exported: Dict[str, Path] = {}
            self._with_retries(
                lambda: exported.update(self._export_from_qb2023(job, export_dir, log_fn)),
                "QB 2023 export",
                log_fn,
            )

            # 3) Close QB 2023
            set_progress(2)
            self._close_qb(qb2023_app, log_fn)
            qb2023_app = None
            self._qb2023_app = None

            # 4) QBFC export produces a JSON snapshot — no CSV parsing needed.
            #    (Legacy step removed — tx_parser.convert_csv_to_iif was for
            #    the old IIF-based flow.  QBFC import reads the snapshot directly.)
            set_progress(3)
            self._emit("QBFC snapshot ready — skipping legacy CSV parse step.", log_fn)

            # 5) Copy QB 2021 template -> working\source\ (KEEP ORIGINAL NAME)
            #    The QBFC app authorization is baked into the template by file name.
            #    We copy as "Blank Template.qbw", do the import, then rename AFTER
            #    closing QB so the authorization stays valid throughout.
            #    The final renamed .qbw gets placed in Final Output\<Company>\.
            set_progress(4)
            raw_name = job.target_company_name or original_qbw.stem
            target_name = re.sub(r"\b23\b", "21", raw_name) if "23" in raw_name else raw_name
            # Final destination is in Final Output\<Company>\
            final_target_qbw = output_dir / f"{target_name}.qbw"
            # Working copy lives in working\source\ (template name for QBFC auth)
            template_path = Path(self.config.install_paths.qb_2021_template_path)
            working_qbw = source_dir / template_path.name

            if not self.config.dry_run:
                if not template_path.exists():
                    raise FileNotFoundError(
                        f"QB 2021 template not found at {template_path}. "
                        "Place a blank QB 2021 .qbw file there first."
                    )
                # Remove stale destination first (previous failed run may leave a locked copy)
                if working_qbw.exists():
                    deleted = False
                    for attempt in range(3):
                        try:
                            working_qbw.unlink()
                            deleted = True
                            break
                        except PermissionError:
                            if attempt == 0:
                                self._emit("Stale template file is locked — killing residual QB processes...", log_fn)
                                import subprocess as _sp
                                for img in ("QBW32.exe", "QBW32PremierAccountant.exe",
                                            "QBWPremierAccountant.exe", "qbupdate.exe",
                                            "QBDBMgrN.exe", "QBDBMgr.exe",
                                            "QBCFMonitorService.exe"):
                                    _sp.run(["taskkill", "/F", "/IM", img], capture_output=True)
                            self._emit(f"  Retry {attempt+1}/3 — waiting 5s for lock release...", log_fn)
                            time.sleep(5)
                    if not deleted:
                        # Last resort: rename and leave behind
                        stale_name = working_qbw.with_suffix(".qbw.stale")
                        try:
                            working_qbw.rename(stale_name)
                            self._emit(f"WARN: Could not delete stale template, renamed to {stale_name.name}", log_fn)
                        except Exception as e2:
                            raise PermissionError(
                                f"Cannot remove locked template {working_qbw}: {e2}. "
                                "Close all QuickBooks instances and try again."
                            ) from e2

                shutil.copy2(template_path, working_qbw)
                self._emit(f"Copied QB 2021 template to {working_qbw} (keeping name for QBFC auth)", log_fn)

                # Delete any stale companion files (.ND, .DSN, .TLG) in
                # the working directory. These may be left over from a
                # previous run and contain wrong path references that cause
                # error 80070057 ("The parameter is incorrect").
                for ext_s in (".qbw.ND", ".qbw.DSN", ".tlg", ".TLG"):
                    stale_f = source_dir / f"{working_qbw.stem}{ext_s}"
                    if stale_f.exists():
                        try:
                            stale_f.unlink()
                        except Exception:
                            pass
            else:
                working_qbw = final_target_qbw
                working_qbw.parent.mkdir(parents=True, exist_ok=True)
                working_qbw.write_text("DRY RUN PLACEHOLDER - QBW FILE CREATED", encoding="utf-8")
                self._emit(f"Created dry-run target company file: {working_qbw}", log_fn)

            # 6) Launch QB 2021 with the WORKING copy (original template name)
            set_progress(5)
            snapshot_path = exported.get("snapshot")
            if not self.config.dry_run and not snapshot_path:
                snapshot_path = export_dir / "company_snapshot.json"
            if not self.config.dry_run and (not snapshot_path or not Path(str(snapshot_path)).exists()):
                raise FileNotFoundError(
                    f"JSON snapshot not found at {snapshot_path}. "
                    "QBFC export must produce company_snapshot.json."
                )

            if not self.config.dry_run:
                qb2021_app = self._launch_qb(
                    self.config.install_paths.qb_2021_path, log_fn, qbw_path=working_qbw
                )
                self._qb2021_app = qb2021_app

                # -------------------------------------------------------
                # Auto-type the template password.
                # The template password is hardcoded (it's OUR template,
                # not the customer's file), so auto-entry is safe.
                # _handle_startup_dialogs() watches for the QB login
                # dialog, types the password, presses Enter, and
                # dismisses any other startup popups.
                # -------------------------------------------------------
                template_password = self.config.install_paths.qb_2021_template_password
                # The template's internal company may be "Blank Template" with
                # a different password than the Tax-Man-Mike template.  Try the
                # configured password first, then fall back to the generic
                # blank-template password.
                alt_pw = "Fl0640098!@!" if template_password != "Fl0640098!@!" else "3825You171"

                # PAUSE the watchdog BEFORE touching QB 2021.
                # If the watchdog is running during _handle_startup_dialogs(),
                # it hides the QB main window (SW_HIDE) as soon as it appears.
                # Hidden windows cascade: child dialogs (like the password
                # dialog) also become invisible, so dialog detection fails
                # and the blind-entry keystroke goes to OUR GUI instead of QB.
                # Keep the watchdog paused until AFTER the QBFC import
                # completes and we close QB 2021.
                if self._watchdog is not None:
                    self._watchdog.pause()
                    self._emit("[Watchdog] paused for QB 2021 phase", log_fn)

                self._emit("Auto-entering template password in QB 2021...", log_fn)
                self._handle_startup_dialogs(
                    password=template_password,
                    alt_password=alt_pw,
                    timeout_s=120,     # generous — QB 2021 can be slow to launch
                    log_fn=log_fn,
                )

                # _handle_startup_dialogs() resumes the watchdog at exit.
                # Re-pause it immediately — QB 2021 must stay visible for
                # _find_qb_main_window and the QBFC import that follows.
                if self._watchdog is not None:
                    self._watchdog.pause()
                    self._emit("[Watchdog] re-paused after startup dialogs", log_fn)

                qb2021_main = self._find_qb_main_window(
                    qb2021_app, "2021",
                    self.config.timeouts.launch_qb_seconds,
                    include_hidden=True,
                    log_fn=log_fn,
                )
                # loaded QB 2021 company.
                # BYPASS: _wait_for_company_ready title-polling keeps timing
                # out even though the title clearly contains "2021".
                # Root-cause is likely a pywinauto/COM threading issue where
                # window_text() returns subtly different bytes that fail the
                # 'in' check despite looking identical in logs.
                # Instead, just wait a fixed 30s for QB to fully stabilize.
                self._emit("[STEP 5] Waiting 30s for QB 2021 to stabilize (dismissing popups)...", log_fn)
                # Instead of sleeping 30s straight, sleep in 5s chunks
                # and dismiss any popups that appear during startup
                for _wait_chunk in range(6):
                    time.sleep(5)
                    self._dismiss_common_dialogs(log_fn)
                    self._close_popup_windows(qb2021_main, log_fn)
                try:
                    _t = qb2021_main.window_text() or ""
                    self._emit(f"[STEP 5] QB 2021 window title: '{_t}'", log_fn)
                    # Sanity check — make sure it's not "No Company Open"
                    if "no company open" in _t.lower():
                        raise RuntimeError("QB 2021 shows 'No Company Open' — template failed to load")
                except RuntimeError:
                    raise
                except Exception as _e:
                    self._emit(f"[STEP 5] Could not read QB 2021 title (non-fatal): {_e}", log_fn)
                self._emit("[STEP 5] QB 2021 is ready for import.", log_fn)
                self._dismiss_common_dialogs(log_fn)
                self._close_popup_windows(qb2021_main, log_fn)

                # ── aggressive modal-dialog sweep ──────────────────────
                # QBFC cannot connect while a modal dialog is showing.
                # The "Usage & Analytics Study" popup and similar modals
                # are often child dialogs of the QB window, not separate
                # top-level windows, so desktop.windows() may miss them.
                # Strategy: try pywinauto child-dialog detection first,
                # then fall back to brute-force keyboard dismissal.
                # ── aggressive modal-dialog sweep ──────────────────────
                # QBFC cannot connect while a modal dialog is showing.
                # The "Usage & Analytics Study" popup is a modal dialog
                # that blocks QBFC. Key insight: clicking "Cancel" on this
                # dialog opens FAQ instead of dismissing it; "Continue"
                # is the correct dismiss button. So we try "Continue" first.
                self._emit("[STEP 5] Sweeping for modal dialogs before QBFC...", log_fn)
                _dismiss_btns = ["Continue", "OK", "Yes", "Close", "No", "Skip", "Later", "Cancel"]
                for sweep in range(5):
                    dismissed = False

                    # 1) Use pywinauto find_windows to locate any dialog that
                    #    might be a child of the QB process
                    try:
                        from pywinauto import findwindows
                        qb_pid = None
                        try:
                            qb_pid = qb2021_main.process_id()
                        except Exception:
                            pass
                        if qb_pid:
                            dlg_handles = findwindows.find_windows(
                                process=qb_pid, top_level_only=False,
                            )
                            for h in dlg_handles:
                                try:
                                    from pywinauto.controls.hwndwrapper import HwndWrapper
                                    w = HwndWrapper(h)
                                    if not w.is_visible():
                                        continue
                                    wt = w.window_text() or ""
                                    if not wt:
                                        continue
                                    # Skip main QB window
                                    if h == getattr(qb2021_main, "handle", None):
                                        continue
                                    wt_l = wt.lower()
                                    if any(k in wt_l for k in [
                                        "usage", "analytics", "study", "survey",
                                        "privacy", "data collection", "have a question",
                                        "faq", "quickbooks desktop",
                                    ]):
                                        self._emit(f"  [sweep {sweep}] Found dialog via PID scan: '{wt}'", log_fn)
                                        # Bring it to front and use keyboard
                                        try:
                                            w.set_focus()
                                            time.sleep(0.3)
                                        except Exception:
                                            pass
                                        # Try clicking buttons via pywinauto
                                        clicked = self._click_first_button(w, _dismiss_btns)
                                        if clicked:
                                            self._emit(f"  [sweep {sweep}] Dismissed via button: '{wt}'", log_fn)
                                            dismissed = True
                                            time.sleep(1)
                                        else:
                                            # Keyboard fallback: Tab to "Continue" then Enter
                                            # (Cancel is default-focused, so Tab moves to Continue)
                                            if send_keys is not None:
                                                send_keys("{TAB}")
                                                time.sleep(0.2)
                                                send_keys("{ENTER}")
                                                time.sleep(0.5)
                                            self._emit(f"  [sweep {sweep}] Dismissed via TAB+ENTER: '{wt}'", log_fn)
                                            dismissed = True
                                            time.sleep(1)
                                except Exception:
                                    continue
                    except Exception as _e:
                        self._emit(f"  [sweep {sweep}] PID scan error (non-fatal): {_e}", log_fn)

                    # 2) Desktop-level scan as fallback
                    if not dismissed:
                        try:
                            desktop = self._get_desktop()
                            for win in desktop.windows():
                                try:
                                    if not win.is_visible():
                                        continue
                                    wt = (win.window_text() or "").lower()
                                    if not wt:
                                        continue
                                    if win.handle == getattr(qb2021_main, "handle", None):
                                        continue
                                    if re.search(self.QB_WINDOW_RE, win.window_text() or ""):
                                        continue
                                    if any(k in wt for k in ["usage", "analytics", "study", "survey",
                                                              "quickbooks desktop", "privacy",
                                                              "have a question", "faq"]):
                                        clicked = self._click_first_button(win, _dismiss_btns)
                                        if not clicked:
                                            try:
                                                win.set_focus()
                                                time.sleep(0.2)
                                                if send_keys is not None:
                                                    send_keys("{TAB}")
                                                    time.sleep(0.2)
                                                    send_keys("{ENTER}")
                                                    time.sleep(0.5)
                                            except Exception:
                                                pass
                                        self._emit(f"  [sweep {sweep}] Dismissed desktop dialog: '{win.window_text()}'", log_fn)
                                        dismissed = True
                                        time.sleep(1)
                                except Exception:
                                    continue
                        except Exception as _e:
                            self._emit(f"  [sweep {sweep}] Desktop scan error (non-fatal): {_e}", log_fn)

                    # 3) Last resort: focus QB and press Escape then Enter
                    if not dismissed:
                        try:
                            self._focus_window(qb2021_main)
                            time.sleep(0.3)
                            if send_keys is not None:
                                send_keys("{ESCAPE}")
                                time.sleep(0.5)
                                send_keys("{ENTER}")
                                time.sleep(0.5)
                        except Exception:
                            pass
                        break  # no more dialogs found
                    # If we dismissed something, loop to catch cascading dialogs

                self._emit("QB 2021 template is open and ready for QBFC import.", log_fn)
            else:
                self._emit("Dry-run: skipping QB 2021 launch", log_fn)

            # 7) QBFC import — pure SDK, no UI automation
            #    Pass qbw_path=None so QBFC binds to the already-open file
            #    (passing the path causes QBFC to try launching a new QB instance)
            set_progress(6)
            import_results: Dict[str, int] = {}
            if not self.config.dry_run:
                from qbfc_import import import_company_via_qbfc

                self._emit("=== Starting QBFC SDK import into QB 2021 ===", log_fn)
                self._emit(f"  Target file: {working_qbw}", log_fn)
                import_results = import_company_via_qbfc(
                    snapshot_path=Path(str(snapshot_path)),
                    qbw_path=None,  # QB 2021 already has the file open; passing path would launch a 2nd instance
                    log_fn=log_fn,
                    skip_transactions=False,
                )

                self._emit(f"QBFC import complete: {sum(import_results.values())} total records", log_fn)

                # ---------------------------------------------------------------
                # ACCOUNTING PREFERENCES via UI automation
                # QBFC PreferencesModRq is unsupported in QB 2021 — use UI instead
                # Must happen AFTER QBFC import (accounts need to exist first)
                # and BEFORE close so QB 2021 is still open.
                # ---------------------------------------------------------------
                try:
                    snapshot_data_prefs = json.loads(Path(str(snapshot_path)).read_text(encoding="utf-8"))
                    snap_prefs = snapshot_data_prefs.get("preferences", {})
                    if snap_prefs:
                        self._emit("=== Setting accounting preferences via UI automation ===", log_fn)
                        if self._watchdog is not None:
                            self._watchdog.pause()
                        self._set_accounting_preferences_via_ui(qb2021_app, snap_prefs, log_fn)
                except Exception as exc:
                    self._emit(f"  AcctPrefsUI: non-fatal error: {exc}", log_fn)

                # ---------------------------------------------------------------
                # COMPANY INFO via UI automation (QBFC has no CompanyMod method)
                # Must happen BEFORE close so QB 2021 is still open.
                # ---------------------------------------------------------------
                try:
                    snapshot_data = json.loads(Path(str(snapshot_path)).read_text(encoding="utf-8"))
                    company_info = snapshot_data.get("company", {})
                    if company_info:
                        self._emit("=== Setting company info via UI automation ===", log_fn)
                        # Make sure the QB window is visible for UI automation
                        if self._watchdog is not None:
                            self._watchdog.pause()
                        self._set_company_info_via_ui(qb2021_app, company_info, log_fn)
                except Exception as exc:
                    self._emit(f"  CompanyUI: non-fatal error: {exc}", log_fn)

                # ---------------------------------------------------------------
                # AUTOMATED CLOSE: drive QB through its own File menu so the
                # data file is flushed cleanly to disk (Ctrl+W to close company,
                # Yes on the save dialog, Alt+F4 to exit). _close_qb() handles
                # all of this and falls back to taskkill if the menu path
                # fails. This must run BEFORE the rename so no file lock.
                # ---------------------------------------------------------------
                self._emit("Import done — auto-closing QB 2021 (menu-driven)...", log_fn)
                # CRITICAL: Keep watchdog PAUSED so _close_qb can see and
                # interact with the QB window.  If the watchdog is running
                # it immediately hides the window, making Ctrl+W / Alt+F4
                # hit nothing → force-kill → data never flushed to disk.
                # This was the root cause of empty .qbw files in Runs 9-11.
                if self._watchdog is not None:
                    self._watchdog.pause()
                self._close_qb(qb2021_app, log_fn)
                qb2021_app = None
                self._qb2021_app = None
            else:
                self._emit("Dry-run: skipping QBFC import", log_fn)

            # 8) Validation
            set_progress(7)
            source_snapshot = self._extract_validation_snapshot(export_dir)
            target_snapshot = self._extract_validation_snapshot(export_dir)
            validation_files = self.validator.generate_reports(
                source_snapshot,
                target_snapshot,
                output_dir,
                company_name=job.qbw_path.stem,
            )

            # 9) Close QB 2021 — skipped in manual-close mode (already None)
            set_progress(8)
            if qb2021_app is not None:
                self._close_qb(qb2021_app, log_fn)
                qb2021_app = None
                self._qb2021_app = None

            # Rename the working copy to the final target name
            # (done AFTER closing QB so no file locks)
            target_qbw = final_target_qbw
            if not self.config.dry_run and working_qbw != final_target_qbw:
                import subprocess as _sp
                # Force-kill any lingering QBW32 / qbupdate processes that may
                # still hold a file handle on the working copy.
                self._emit("Killing any lingering QuickBooks processes before rename...", log_fn)
                try:
                    _sp.run(
                        ["powershell", "-NoProfile", "-Command",
                         "Get-Process | Where-Object {$_.Name -like 'QBW32*' -or $_.Name -like 'qbw*' -or $_.Name -like 'qbupdate*' -or $_.Name -like 'QBDBMgr*' -or $_.Name -like 'QBCFMonitor*'} | Stop-Process -Force -ErrorAction SilentlyContinue"],
                        timeout=15, capture_output=True,
                    )
                except Exception:
                    pass
                # CRITICAL: Stop the QuickBooks Database Manager service.
                # The actual service name is "QuickBooksDB33" (not QBDBMgrN).
                self._emit("Stopping QuickBooksDB33 service to release file locks...", log_fn)
                try:
                    _sp.run(["sc", "stop", "QuickBooksDB33"],
                            timeout=15, capture_output=True)
                except Exception:
                    pass
                try:
                    _sp.run(
                        ["powershell", "-NoProfile", "-Command",
                         "Stop-Service QuickBooksDB* -Force -ErrorAction SilentlyContinue"],
                        timeout=15, capture_output=True,
                    )
                except Exception:
                    pass
                time.sleep(8)  # extra time for service to fully release locks

                self._emit(f"Renaming {working_qbw.name} -> {final_target_qbw.name}", log_fn)
                # Retry the actual move with backoff (file lock can linger on Windows)
                last_err: Optional[Exception] = None
                for attempt in range(30):
                    try:
                        if final_target_qbw.exists():
                            final_target_qbw.unlink()
                        shutil.move(str(working_qbw), str(final_target_qbw))
                        last_err = None
                        break
                    except (PermissionError, OSError) as e:
                        last_err = e
                        if attempt == 0:
                            self._emit(f"  File still locked, retrying... ({e})", log_fn)
                        time.sleep(3)
                if last_err is not None:
                    raise last_err
                # Move companion files from working\source\ -> Final Output\<Company>\
                for ext_suffix in (".qbw.ND", ".qbw.DSN", ".tlg"):
                    old_f = source_dir / f"{working_qbw.stem}{ext_suffix}"
                    new_f = output_dir / f"{final_target_qbw.stem}{ext_suffix}"
                    if old_f.exists():
                        try:
                            if new_f.exists():
                                new_f.unlink()
                            shutil.move(str(old_f), str(new_f))
                        except Exception:
                            pass
            set_progress(9)

            generated_files = {
                "target_qbw": str(target_qbw),
            }
            # Include whatever the export phase produced
            for key, val in exported.items():
                generated_files[key] = str(val)
            # Validation reports (may not exist on dry-run)
            if validation_files.get("excel"):
                generated_files["validation_excel"] = validation_files["excel"]
            if validation_files.get("html"):
                generated_files["validation_html"] = validation_files["html"]
            if import_results:
                generated_files["import_summary"] = str(import_results)

            # Generate Memorized Transactions report (Excel + PDF) so the
            # operator has a printable cheat-sheet for re-memorizing the
            # templates QBFC cannot recreate.
            try:
                snapshot_path = export_dir / "company_snapshot.json"
                if snapshot_path.exists():
                    import json as _json
                    with snapshot_path.open("r", encoding="utf-8") as fh:
                        snap = _json.load(fh)
                    memorized = snap.get("memorized_txns") or []
                    if memorized:
                        from memorized_report import generate_reports as _gen_memo
                        memo_out = _gen_memo(memorized, output_dir, original_qbw.stem)
                        if memo_out.get("xlsx"):
                            generated_files["memorized_xlsx"] = str(memo_out["xlsx"])
                            self._emit(f"  Memorized report (Excel): {memo_out['xlsx']}", log_fn)
                        if memo_out.get("pdf"):
                            generated_files["memorized_pdf"] = str(memo_out["pdf"])
                            self._emit(f"  Memorized report (PDF):   {memo_out['pdf']}", log_fn)
            except Exception as _exc:  # noqa: BLE001
                self._emit(f"  Memorized report generation failed: {_exc}", log_fn)

            return CompanyJobResult(
                qbw_path=str(job.qbw_path),
                success=True,
                message="Completed successfully",
                generated_files=generated_files,
            )
        except Exception as exc:  # noqa: BLE001
            screenshot = self._take_error_screenshot("run_job_failure")
            stack = traceback.format_exc()
            self._emit(f"Job failed: {exc}\n{stack}", log_fn)
            if screenshot is not None:
                self._emit(f"Failure screenshot: {screenshot}", log_fn)
            return CompanyJobResult(
                qbw_path=str(job.qbw_path),
                success=False,
                message=str(exc),
                generated_files={},
            )
        finally:
            if qb2023_app is not None:
                self._close_qb(qb2023_app, log_fn)
            if qb2021_app is not None:
                self._close_qb(qb2021_app, log_fn)
            self._qb2023_app = None
            self._qb2021_app = None
            # Stop dialog watchdog last so it can dismiss any final popups
            # produced while QB shuts down.
            try:
                if watchdog is not None:
                    watchdog.stop()
            except Exception:  # noqa: BLE001
                pass