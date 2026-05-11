"""Finish / download dialog shown after a company completes successfully.

Lets the operator copy the converted .qbw, the validation report, and the
memorized-transactions cheat sheet (Excel/PDF) into the user's Downloads
folder with one click. Also offers links to leave a review and a Close button.
"""

from __future__ import annotations

import os
import shutil
import sys
import webbrowser
from pathlib import Path
from tkinter import BOTH, LEFT, RIGHT, X, Y, messagebox
import tkinter as tk
from tkinter import ttk
from typing import Any, Dict, Optional


# TODO: Replace these placeholders with the real review URLs once Joseph
# sets up the listings. To update:
#   - Google:     Google Business Profile -> Reviews -> "Get more reviews"
#                 (copy the link that looks like https://g.page/r/.../review)
#   - Trustpilot: https://business.trustpilot.com -> claim domain -> share URL
#   - Facebook:   https://www.facebook.com/<page>/reviews/
REVIEW_URLS = {
    "Google":     "https://www.google.com/search?q=our+system+administrator+reviews",   # placeholder
    "Trustpilot": "https://www.trustpilot.com/review/oursystemadmin.com",                # placeholder
    "Facebook":   "https://www.facebook.com/oursystemadmin",                             # placeholder
}


def _downloads_dir() -> Path:
    """Best-effort Downloads folder for the current user across OSes."""
    home = Path.home()
    candidates = [home / "Downloads", home / "downloads"]
    for c in candidates:
        if c.exists():
            return c
    # Windows fallback via USERPROFILE
    up = os.environ.get("USERPROFILE")
    if up:
        d = Path(up) / "Downloads"
        if d.exists():
            return d
    # Last resort: create ~/Downloads
    d = home / "Downloads"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _reveal_in_explorer(path: Path) -> None:
    """Open the file's containing folder with the file selected."""
    try:
        if sys.platform == "win32":
            os.startfile(str(path.parent))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open -R "{path}"')
        else:
            os.system(f'xdg-open "{path.parent}"')
    except Exception:
        pass


class FinishDialog(tk.Toplevel):
    """Modal-ish download screen with copy buttons and review prompts."""

    def __init__(self, parent: tk.Misc, result: Any, original_name: str) -> None:
        super().__init__(parent)
        self.title(f"Conversion Complete — {original_name}")
        self.geometry("640x500")
        self.configure(bg="#0f172a")
        self.resizable(False, False)

        self._result = result
        self._files: Dict[str, str] = dict(getattr(result, "generated_files", {}) or {})
        self._downloads = _downloads_dir()

        self._build_ui(original_name)

        # Center on parent
        self.update_idletasks()
        try:
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            x = px + (pw - self.winfo_width()) // 2
            y = py + (ph - self.winfo_height()) // 2
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except Exception:
            pass

        self.transient(parent)  # type: ignore[arg-type]
        self.grab_set()
        self.focus_force()
        self.lift()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self, original_name: str) -> None:
        # Header
        header = tk.Frame(self, bg="#1e293b", height=80)
        header.pack(fill=X)
        header.pack_propagate(False)
        tk.Label(header, text="✓  Conversion Complete",
                 bg="#1e293b", fg="#4ade80",
                 font=("Segoe UI", 18, "bold")).pack(pady=(14, 0))
        tk.Label(header, text=original_name,
                 bg="#1e293b", fg="#cbd5e1",
                 font=("Segoe UI", 10)).pack()

        # Body
        body = tk.Frame(self, bg="#0f172a", padx=20, pady=14)
        body.pack(fill=BOTH, expand=True)

        tk.Label(body, text=f"Save your files to:  {self._downloads}",
                 bg="#0f172a", fg="#94a3b8",
                 font=("Segoe UI", 9, "italic")).pack(anchor="w", pady=(0, 8))

        # Per-file rows: label + Copy to Downloads + Show in folder
        self._row(body, "Converted QuickBooks file (.QBW)",
                  self._files.get("target_qbw"))
        self._row(body, "Memorized Transactions — Excel",
                  self._files.get("memorized_xlsx"))
        self._row(body, "Memorized Transactions — PDF",
                  self._files.get("memorized_pdf"))
        self._row(body, "Validation Report — Excel",
                  self._files.get("validation_excel"))
        self._row(body, "Validation Report — HTML",
                  self._files.get("validation_html"))

        # Copy ALL button
        tk.Button(body, text="⬇  Copy ALL to Downloads",
                  command=self._copy_all,
                  bg="#22d3ee", fg="#0f172a",
                  activebackground="#06b6d4", activeforeground="#0f172a",
                  font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=14, pady=8, cursor="hand2",
                  ).pack(fill=X, pady=(10, 4))

        # (Review section removed — Matt Mike internal build)

        # Footer / Close
        footer = tk.Frame(self, bg="#0f172a", padx=20, pady=10)
        footer.pack(fill=X, side=tk.BOTTOM)
        tk.Button(footer, text="Close",
                  command=self.destroy,
                  bg="#334155", fg="#e2e8f0",
                  activebackground="#475569", activeforeground="#e2e8f0",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=18, pady=6, cursor="hand2",
                  ).pack(side=RIGHT)

    def _row(self, parent: tk.Misc, label: str, path_str: Optional[str]) -> None:
        row = tk.Frame(parent, bg="#0f172a")
        row.pack(fill=X, pady=3)

        avail = bool(path_str) and Path(path_str).exists() if path_str else False
        fg = "#e2e8f0" if avail else "#64748b"
        suffix = "" if avail else "  (not generated)"

        tk.Label(row, text=label + suffix, bg="#0f172a", fg=fg,
                 font=("Segoe UI", 9), anchor="w").pack(side=LEFT, fill=X, expand=True)

        copy_btn = tk.Button(row, text="Copy to Downloads",
                             command=lambda p=path_str: self._copy_one(p),
                             bg="#0ea5e9" if avail else "#1e293b",
                             fg="#0f172a" if avail else "#64748b",
                             relief="flat", padx=10, pady=3,
                             cursor="hand2" if avail else "arrow",
                             state=tk.NORMAL if avail else tk.DISABLED,
                             font=("Segoe UI", 8, "bold"))
        copy_btn.pack(side=LEFT, padx=(6, 4))

        show_btn = tk.Button(row, text="Show in Folder",
                             command=lambda p=path_str: _reveal_in_explorer(Path(p)) if p else None,
                             bg="#1e293b", fg="#94a3b8",
                             relief="flat", padx=10, pady=3,
                             cursor="hand2" if avail else "arrow",
                             state=tk.NORMAL if avail else tk.DISABLED,
                             font=("Segoe UI", 8))
        show_btn.pack(side=LEFT)

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _copy_one(self, path_str: Optional[str]) -> Optional[Path]:
        if not path_str:
            return None
        src = Path(path_str)
        if not src.exists():
            messagebox.showwarning("Not found", f"{src.name} no longer exists.", parent=self)
            return None
        try:
            dst = self._downloads / src.name
            # Avoid clobbering an existing file: append " (n)" if needed
            n = 1
            while dst.exists():
                dst = self._downloads / f"{src.stem} ({n}){src.suffix}"
                n += 1
            shutil.copy2(str(src), str(dst))
            messagebox.showinfo("Copied", f"Saved to:\n{dst}", parent=self)
            return dst
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Copy failed", str(exc), parent=self)
            return None

    def _copy_all(self) -> None:
        copied: list[Path] = []
        keys = ("target_qbw", "memorized_xlsx", "memorized_pdf",
                "validation_excel", "validation_html")
        for k in keys:
            p = self._files.get(k)
            if not p or not Path(p).exists():
                continue
            dst = self._copy_silent(Path(p))
            if dst is not None:
                copied.append(dst)
        if copied:
            msg = f"Copied {len(copied)} file(s) to:\n{self._downloads}\n\n"
            msg += "\n".join(f"  • {p.name}" for p in copied)
            messagebox.showinfo("Done", msg, parent=self)
            # Open the Downloads folder
            try:
                if sys.platform == "win32":
                    os.startfile(str(self._downloads))  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    os.system(f'open "{self._downloads}"')
                else:
                    os.system(f'xdg-open "{self._downloads}"')
            except Exception:
                pass
        else:
            messagebox.showwarning("Nothing to copy",
                                   "None of the expected files were available.",
                                   parent=self)

    def _copy_silent(self, src: Path) -> Optional[Path]:
        try:
            dst = self._downloads / src.name
            n = 1
            while dst.exists():
                dst = self._downloads / f"{src.stem} ({n}){src.suffix}"
                n += 1
            shutil.copy2(str(src), str(dst))
            return dst
        except Exception:
            return None


__all__ = ["FinishDialog"]