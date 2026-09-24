"""Python API exposed to the web UI through pywebview.

All method names and signatures match exactly what the React app calls via
`callApi()` in `src/lib/wj-bridge.ts`. Place this file, `main.py`,
`excel_automations.py`, and `selenium_automations.py` in the same folder as your
existing Python scripts, then run `python main.py`.

Threading note: every long-running automation is started on a background
thread so the pywebview UI stays responsive. Progress and logs are pushed back
to the page with `window.evaluate_js`.
"""

import json
import os
import sys
import threading
import tkinter as tk
import traceback
from tkinter import filedialog
from typing import Any, Optional

import webview

import excel_automations
import selenium_automations


class UI:
    """Helper that forwards Python stdout/status/progress to the React console."""

    def __init__(self, finished_callback: Optional[str] = None):
        self.finished_callback = finished_callback

    def _get_window(self):
        """Get the pywebview window safely from any thread."""
        try:
            wins = webview.windows
            if wins:
                return wins[0]
        except Exception:
            pass
        try:
            return webview.active_window()
        except Exception:
            return None

    def _js(self, code: str) -> None:
        try:
            window = self._get_window()
            if window:
                window.evaluate_js(code)
        except Exception:
            pass

    def log(self, message: Any, tag: str = "info") -> None:
        text = str(message)
        tag_prefix = {"ok": "✓", "err": "✗", "fail": "✗", "warn": "⚠", "pass": "✓"}.get(tag, "·")
        try:
            if sys.__stdout__:
                sys.__stdout__.write(f"[{tag_prefix}] {text}\n")
                sys.__stdout__.flush()
        except Exception:
            pass
        self._js(
            f"if(window.wjLog){{try{{window.wjLog({json.dumps(text)}, {json.dumps(tag)});}}catch(e){{}}}}"
        )

    def status(self, message: Any, kind: str = "idle") -> None:
        text = str(message)
        try:
            if sys.__stdout__:
                sys.__stdout__.write(f"[STATUS] {text}\n")
                sys.__stdout__.flush()
        except Exception:
            pass
        self._js(
            f"if(window.wjStatus){{try{{window.wjStatus({json.dumps(text)}, {json.dumps(kind)});}}catch(e){{}}}}"
        )

    def progress(self, percent: float) -> None:
        self._js(
            f"if(window.wjProgress){{try{{window.wjProgress({float(percent)});}}catch(e){{}}}}"
        )

    def finished(self) -> None:
        if self.finished_callback:
            self._js(
                f"if(window.{self.finished_callback}){{try{{window.{self.finished_callback}();}}catch(e){{}}}}"
            )


class StdoutToUI:
    """Captures direct print() output and streams each line to the React log console."""

    def __init__(self, ui: UI):
        self.ui = ui
        self._buffer = ""
        self._lock = threading.Lock()

    def write(self, text: str) -> None:
        with self._lock:
            self._buffer += text
            while "\n" in self._buffer:
                line, self._buffer = self._buffer.split("\n", 1)
                clean_line = line.strip()
                if clean_line:
                    tag = "info"
                    lower = clean_line.lower()
                    words = lower.split()
                    if clean_line.startswith("FAIL") or clean_line.startswith("[✗]") or any(k in lower for k in ["error:", "error in", "exception:", "timed out"]):
                        tag = "fail"
                    elif clean_line.startswith("PASS") or clean_line.startswith("[✓]") or "✓" in clean_line:
                        tag = "pass"
                    elif any(w in ["pass", "passed", "success", "successful", "succeeded", "ok"] for w in words):
                        tag = "pass"
                    elif any(w in ["fail", "failed", "failure", "err", "error"] for w in words):
                        tag = "fail"
                    elif any(w in ["warn", "warning", "caution"] for w in words):
                        tag = "warn"

                    self.ui._js(
                        f"if(window.wjLog){{try{{window.wjLog({json.dumps(clean_line)}, {json.dumps(tag)});}}catch(e){{}}}}"
                    )

    def flush(self) -> None:
        with self._lock:
            clean = self._buffer.strip()
            if clean:
                self.ui._js(
                    f"if(window.wjLog){{try{{window.wjLog({json.dumps(clean)}, 'info');}}catch(e){{}}}}"
                )
                self._buffer = ""


def _run_with_captured_output(target, ui: UI) -> None:
    """Run target(), safely capturing print() output without recursion."""
    old_out, old_err = sys.stdout, sys.stderr
    try:
        redirector = StdoutToUI(ui)
        sys.stdout = redirector
        sys.stderr = redirector
        target()
    finally:
        try:
            if hasattr(sys.stdout, "flush"):
                sys.stdout.flush()
        except Exception:
            pass
        sys.stdout, sys.stderr = old_out, old_err


def _run_in_thread(name: str, finished_callback: str, target) -> None:
    def wrapper() -> None:
        ui = UI(finished_callback)
        ui.log(f"Starting: {name}")
        ui.status(f"Running {name}...", "idle")
        try:
            _run_with_captured_output(lambda: target(ui), ui)
        except Exception as exc:
            err_text = f"Error in {name}: {exc}"
            ui.log(err_text, "err")
            ui.status(f"{name} failed: {exc}", "fail")
            traceback.print_exc()
        finally:
            ui.finished()

    threading.Thread(target=wrapper, daemon=True).start()


class Api:
    """Methods on this class become `window.pywebview.api.<methodName>()`."""

    # ---------- File pickers ----------

    def pick_excel_file(self) -> str:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        root.destroy()
        return path or ""

    def pick_driver_file(self) -> str:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select the WebDriver executable",
            filetypes=[("Executables", "*.exe"), ("All files", "*.*")],
        )
        root.destroy()
        return path or ""

    # ---------- Excel automations ----------

    def run_salary_automation(self, salary_file: str, work_file: str, template_file: str) -> None:
        _run_in_thread(
            "Salary File Automation",
            "wjSalaryFinished",
            lambda ui: excel_automations.run_salary_automation(salary_file, work_file, template_file, ui),
        )

    def run_enrollment_audit(self, pre_file: str, post_file: str) -> None:
        _run_in_thread(
            "Enrollment Report Audit",
            "wjEnrollmentFinished",
            lambda ui: excel_automations.run_enrollment_audit(pre_file, post_file, ui),
        )

    def run_comparison_tool(self, pre_file: str, post_file: str) -> None:
        _run_in_thread(
            "Pre vs Post Comparison",
            "wjComparisonFinished",
            lambda ui: excel_automations.run_comparison_tool(pre_file, post_file, ui),
        )

    # ---------- Portal / Selenium automations ----------

    def run_new_client_password_setup(self, excel_path: str, portal_url: str, password: str) -> None:
        _run_in_thread(
            "New Client Password Setup",
            "wjPwSetupFinished",
            lambda ui: selenium_automations.run_new_client_password_setup(excel_path, portal_url, password, ui),
        )

    def run_new_client_auto_login(self, excel_path: str, portal_url: str) -> None:
        _run_in_thread(
            "New Client Auto Login",
            "wjTestFinished",
            lambda ui: selenium_automations.run_new_client_auto_login(excel_path, portal_url, ui),
        )

    def run_auto_login_test(self, excel_path: str, portal_url: str, expected_text: str) -> None:
        _run_in_thread(
            "Auto-Login Tester",
            "wjTestFinished",
            lambda ui: selenium_automations.run_auto_login_test(excel_path, portal_url, expected_text, ui),
        )

    def run_custom_report(self, params: dict) -> None:
        _run_in_thread(
            "Custom Report Generator",
            "wjReportFinished",
            lambda ui: selenium_automations.run_custom_report(params, ui),
        )
