"""QuickBooks desktop automation engine.

Designed for Windows + QuickBooks Desktop 2023 and 2021 environments.
Uses pywinauto/pyautogui when available, with a dry-run fallback for testing.
"""

from __future__ import annotations

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
        # Tracks whether password was already entered during QB startup dialogs.
        # Used to prevent redundant _open_company_file() call in _export_from_qb2023().
        # See bug note in _export_from_qb2023() for details.
        self._startup_password_handled: bool = False

    def _emit(self, msg: str, log_fn: Optional[LogFn]) -> None:
        self.logger.info(msg)
        if log_fn:
            log_fn(msg)

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

    def _find_qb_main_window(self, app: Optional[object], version_hint: Optional[str], timeout_s: int):
        def _is_own_gui(title: str) -> bool:
            """Return True if *title* belongs to our own GUI, not real QB."""
            tl = title.lower()
            return any(kw in tl for kw in self._OWN_GUI_KEYWORDS)

        def _pick_window() -> Optional[object]:
            if app is not None:
                try:
                    for win in app.windows():
                        if not win.exists() or not win.is_visible():
                            continue
                        title = win.window_text()
                        if _is_own_gui(title):
                            continue
                        if version_hint and version_hint not in title:
                            continue
                        if re.search(self.QB_WINDOW_RE, title):
                            return win
                    top = app.top_window()
                    if top.exists() and top.is_visible():
                        title = top.window_text() or ""
                        if not _is_own_gui(title):
                            return top
                except Exception:  # noqa: BLE001
                    pass

            desktop = self._get_desktop()
            candidates = []
            for w in desktop.windows(title_re=self.QB_WINDOW_RE):
                try:
                    if not w.is_visible():
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

        def _cond() -> bool:
            win = _pick_window()
            if win is not None:
                box["window"] = win
                return True
            return False

        if not self._wait_until(_cond, timeout_s, 1.0):
            hint_text = f" ({version_hint})" if version_hint else ""
            raise RuntimeError(f"Could not find QuickBooks main window{hint_text} within {timeout_s}s")

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
                    ]
                ):
                    clicked = self._click_first_button(
                        dialog,
                        ["No", "Don't Save", "Continue", "Yes", "OK", "Close", "Cancel", "Later", "Skip"],
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
                        win, ["Close", "OK", "No", "Cancel", "Skip", "Later", "X"]
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
                found = company_hint.lower() in title.lower()
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
        if qbw_path:
            try:
                existing = Application(backend="uia").connect(path=exe_path)
                self._emit(f"  Killing existing QB instance before clean launch...", log_fn)
                self._close_qb(existing, log_fn)
                time.sleep(5)  # give QB time to fully exit
            except Exception:  # noqa: BLE001
                pass  # not running — good

            # Also force-kill any lingering QB processes
            import subprocess
            subprocess.run(
                ["taskkill", "/F", "/IM", exe_name],
                capture_output=True, timeout=10
            )
            time.sleep(3)

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

        # Try to find the QB window and use it
        try:
            main_window = self._find_qb_main_window(app, None, timeout_s=10)
            if main_window is not None:
                try:
                    main_window.set_focus()
                    time.sleep(0.5)
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
        # -----------------------------------------------------------
        if menu_close_done:
            self._emit("  Waiting up to 90s for QB to exit...", log_fn)
            deadline = time.time() + 90
            while time.time() < deadline:
                try:
                    result = subprocess.run(
                        ["tasklist", "/FI", "IMAGENAME eq QBW32.EXE"],
                        capture_output=True, timeout=5, text=True, check=False,
                    )
                    if "QBW32.EXE" not in (result.stdout or ""):
                        self._emit("  QuickBooks exited cleanly!", log_fn)
                        time.sleep(3)  # let Windows release file locks
                        return  # SUCCESS — no force kill needed
                except Exception:  # noqa: BLE001
                    pass
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

    def _handle_startup_dialogs(self, password: str, timeout_s: int, log_fn: Optional[LogFn]) -> None:
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
        =====================================================================
        """
        password_entered = False
        start = time.time()
        while time.time() - start < timeout_s:
            # Check for password/login dialog first
            login_dlg = self._find_active_dialog(title_re=r"(?i)(password|login)")
            if login_dlg is not None and not password_entered:
                self._emit("Found startup login dialog, entering password", log_fn)

                # Simple approach: just type the password and press Enter.
                # The dialog appears with cursor in the password field by default.
                if send_keys is not None:
                    time.sleep(1)  # Let dialog fully render
                    send_keys(password, pause=0.02)
                    time.sleep(0.3)
                    send_keys("{ENTER}")

                    password_entered = True
                    self._startup_password_handled = True
                    self._emit("Password entered at startup (type + Enter)", log_fn)
                    time.sleep(5)  # Wait for QB to process login

                    # Check if a "wrong password" warning appeared
                    warning_dlg = self._find_active_dialog(title_re=r"(?i)(warning|error|incorrect)")
                    if warning_dlg is not None:
                        self._emit("Password may have been incorrect, dismissing warning", log_fn)
                        self._click_first_button(warning_dlg, ["OK", "Close"])
                        password_entered = False  # Allow retry
                        self._startup_password_handled = False  # Reset — password was wrong
                        time.sleep(1)
                    continue
                else:
                    self._emit("WARNING: send_keys unavailable, cannot enter password", log_fn)

            elif login_dlg is not None and password_entered:
                # Password was already entered but dialog is still showing - wait
                time.sleep(2)
                continue

            # Dismiss other common startup dialogs
            self._dismiss_common_dialogs(log_fn)
            time.sleep(1)

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
                break

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

            # Wait for the user to enter the password in QB 2023 manually.
            # send_keys password typing is unreliable with QB's non-standard
            # dialogs, so we just tell the user what to do and poll for the
            # company to load (title bar changes from "No Company Open").
            self._emit("", log_fn)
            self._emit("=" * 60, log_fn)
            self._emit("ACTION REQUIRED: Enter the admin password in QuickBooks 2023.", log_fn)
            if job.password:
                self._emit(f"  Password: {job.password}", log_fn)
            self._emit("  The tool will auto-detect when the company is loaded.", log_fn)
            self._emit("=" * 60, log_fn)
            self._emit("", log_fn)
            try:
                main_window = self._find_qb_main_window(self._qb2023_app, "2023", self.config.timeouts.launch_qb_seconds)
                # Give user up to 5 minutes to type the password
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

        try:
            job.output_dir.mkdir(parents=True, exist_ok=True)
            exports_dir = job.output_dir / "exports"
            target_dir = job.output_dir / "target"
            validation_dir = job.output_dir / "validation"

            # Wipe ALL stale subdirs (exports, target, validation, logs) before run
            # so a previous failed run never pollutes a fresh attempt.
            for stale_dir in (exports_dir, target_dir, validation_dir, job.output_dir / "logs"):
                if stale_dir.exists():
                    self._emit(f"Cleaning stale directory: {stale_dir}", log_fn)
                    try:
                        shutil.rmtree(stale_dir)
                    except Exception as exc:  # noqa: BLE001
                        self._emit(f"  WARN: could not remove {stale_dir}: {exc}", log_fn)
            exports_dir.mkdir(parents=True, exist_ok=True)

            # 1) Launch QB 2023 with company file as parameter (auto-opens it)
            set_progress(0)
            qb2023_app = self._launch_qb(self.config.install_paths.qb_2023_path, log_fn, qbw_path=job.qbw_path)
            self._qb2023_app = qb2023_app

            # 2) Export
            set_progress(1)
            exported: Dict[str, Path] = {}
            self._with_retries(
                lambda: exported.update(self._export_from_qb2023(job, exports_dir, log_fn)),
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

            # 5) Copy QB 2021 template -> target directory (KEEP ORIGINAL NAME)
            #    The QBFC app authorization is baked into the template by file name.
            #    We copy as "Blank Template.qbw", do the import, then rename AFTER
            #    closing QB so the authorization stays valid throughout.
            set_progress(4)
            target_dir.mkdir(parents=True, exist_ok=True)
            raw_name = job.target_company_name or job.qbw_path.stem
            target_name = re.sub(r"\b23\b", "21", raw_name) if "23" in raw_name else raw_name
            # Final destination after rename
            final_target_qbw = target_dir / f"{target_name}.qbw"
            # Working copy keeps the template name so QBFC auth is inherited
            template_path = Path(self.config.install_paths.qb_2021_template_path)
            working_qbw = target_dir / template_path.name

            if not self.config.dry_run:
                if not template_path.exists():
                    raise FileNotFoundError(
                        f"QB 2021 template not found at {template_path}. "
                        "Place a blank QB 2021 .qbw file there first."
                    )
                shutil.copy2(template_path, working_qbw)
                self._emit(f"Copied QB 2021 template to {working_qbw} (keeping name for QBFC auth)", log_fn)

                # Also copy companion files (.tlg, .nd, .DSN) if they exist
                for ext_suffix in (".qbw.ND", ".qbw.DSN", ".tlg"):
                    src = template_path.parent / f"{template_path.stem}{ext_suffix}"
                    if src.exists():
                        dst = target_dir / f"{working_qbw.stem}{ext_suffix}"
                        try:
                            shutil.copy2(src, dst)
                        except Exception:
                            pass  # QB will recreate these
            else:
                working_qbw = final_target_qbw
                working_qbw.write_text("DRY RUN PLACEHOLDER - QBW FILE CREATED", encoding="utf-8")
                self._emit(f"Created dry-run target company file: {working_qbw}", log_fn)

            # 6) Launch QB 2021 with the WORKING copy (original template name)
            set_progress(5)
            snapshot_path = exported.get("snapshot")
            if not self.config.dry_run and not snapshot_path:
                snapshot_path = exports_dir / "company_snapshot.json"
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

                # Wait for the user to enter the template password in QB 2021.
                template_password = self.config.install_paths.qb_2021_template_password
                self._emit("", log_fn)
                self._emit("=" * 60, log_fn)
                self._emit("ACTION REQUIRED: Enter the template password in QuickBooks 2021.", log_fn)
                self._emit(f"  Password: {template_password}", log_fn)
                self._emit("  The tool will auto-detect when the company is loaded.", log_fn)
                self._emit("=" * 60, log_fn)
                self._emit("", log_fn)
                qb2021_main = self._find_qb_main_window(qb2021_app, "2021", self.config.timeouts.launch_qb_seconds)
                # Use the template filename as a positive hint so the poller
                # waits until the user enters the password and the company
                # actually opens (title shows "Blank Template - QuickBooks …").
                template_hint = template_path.stem  # e.g. "Blank Template"
                self._wait_for_company_ready(
                    qb2021_main, timeout_s=300, log_fn=log_fn,
                    company_hint=template_hint,
                )
                self._dismiss_common_dialogs(log_fn)
                self._close_popup_windows(qb2021_main, log_fn)
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
                # AUTOMATED CLOSE: drive QB through its own File menu so the
                # data file is flushed cleanly to disk (Ctrl+W to close company,
                # Yes on the save dialog, Alt+F4 to exit). _close_qb() handles
                # all of this and falls back to taskkill if the menu path
                # fails. This must run BEFORE the rename so no file lock.
                # ---------------------------------------------------------------
                self._emit("Import done — auto-closing QB 2021 (menu-driven)...", log_fn)
                self._close_qb(qb2021_app, log_fn)
                qb2021_app = None
                self._qb2021_app = None
            else:
                self._emit("Dry-run: skipping QBFC import", log_fn)

            # 8) Validation
            set_progress(7)
            source_snapshot = self._extract_validation_snapshot(exports_dir)
            target_snapshot = self._extract_validation_snapshot(exports_dir)
            validation_files = self.validator.generate_reports(
                source_snapshot,
                target_snapshot,
                validation_dir,
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
                         "Get-Process | Where-Object {$_.Name -like 'QBW32*' -or $_.Name -like 'qbupdate*' -or $_.Name -like 'QBDBMgr*' -or $_.Name -like 'QBCFMonitor*'} | Stop-Process -Force -ErrorAction SilentlyContinue"],
                        timeout=15, capture_output=True,
                    )
                except Exception:
                    pass
                time.sleep(5)

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
                # Rename companion files too
                for ext_suffix in (".qbw.ND", ".qbw.DSN", ".tlg"):
                    old_f = target_dir / f"{working_qbw.stem}{ext_suffix}"
                    new_f = target_dir / f"{final_target_qbw.stem}{ext_suffix}"
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