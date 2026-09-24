"""
Benefits Enrollment Portal - GUI Test Runner (Fixed)
- Reads only Username + Password from Excel (handles any column structure)
- Finds popup via span#PopUp_LblMsg (correct element from DevTools)
- Shows editable assertion text box in UI
"""

import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


class TestRunnerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Benefits Enrollment Portal - Test Runner")
        self.root.geometry("860x800")
        self.root.resizable(True, True)
        self.root.configure(bg="#1e1e2e")
        self.excel_path = tk.StringVar()
        self.portal_url = tk.StringVar(value="https://www.benefitsjunction.com/")
        self.running    = False
        self._build_ui()

    def _build_ui(self):
        BG     = "#1e1e2e"
        CARD   = "#2a2a3e"
        ACCENT = "#7c3aed"
        TEXT   = "#e2e8f0"
        MUTED  = "#94a3b8"
        BORDER = "#3f3f5a"
        GREEN  = "#22c55e"

        # Header
        hdr = tk.Frame(self.root, bg=ACCENT, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Benefits Enrollment Portal",
                 font=("Segoe UI", 16, "bold"), bg=ACCENT, fg="white").pack()
        tk.Label(hdr, text="Automated Login & Popup Assertion Tester",
                 font=("Segoe UI", 9), bg=ACCENT, fg="#ddd6fe").pack()

        # Config card
        card = tk.Frame(self.root, bg=CARD, padx=20, pady=14,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", padx=20, pady=(14, 0))
        card.columnconfigure(0, weight=1)

        # Portal URL
        tk.Label(card, text="Portal URL", font=("Segoe UI", 9, "bold"),
                 bg=CARD, fg=MUTED).grid(row=0, column=0, sticky="w")
        tk.Entry(card, textvariable=self.portal_url, font=("Consolas", 10),
                 bg="#0f0f1a", fg=TEXT, insertbackground=TEXT, relief="flat",
                 highlightbackground=BORDER, highlightthickness=1
                 ).grid(row=1, column=0, sticky="ew", ipady=6, pady=(2,10), padx=1)

        # Excel file picker
        tk.Label(card,
                 text="Test Records Excel File  —  needs 'Username' & 'Password' columns",
                 font=("Segoe UI", 9, "bold"), bg=CARD, fg=MUTED
                 ).grid(row=2, column=0, sticky="w")
        fr = tk.Frame(card, bg=CARD)
        fr.grid(row=3, column=0, sticky="ew", pady=(2,10))
        fr.columnconfigure(0, weight=1)
        tk.Entry(fr, textvariable=self.excel_path, font=("Consolas", 9),
                 bg="#0f0f1a", fg=TEXT, insertbackground=TEXT, relief="flat",
                 highlightbackground=BORDER, highlightthickness=1,
                 state="readonly").grid(row=0, column=0, sticky="ew", ipady=6, padx=(1,6))
        tk.Button(fr, text="Browse", font=("Segoe UI", 9, "bold"),
                  bg=ACCENT, fg="white", relief="flat",
                  activebackground="#6d28d9", cursor="hand2",
                  padx=14, pady=6,
                  command=self._browse).grid(row=0, column=1)

        # ── Assertion box ─────────────────────────────────────────────────────
        tk.Label(card,
                 text="Expected Popup Text  (each line checked as substring — edit if needed)",
                 font=("Segoe UI", 9, "bold"), bg=CARD, fg=MUTED
                 ).grid(row=4, column=0, sticky="w", pady=(4,2))

        self.assert_box = tk.Text(card, height=5, font=("Consolas", 9),
                                  bg="#0f0f1a", fg="#a5f3fc",
                                  insertbackground=TEXT, relief="flat",
                                  highlightbackground="#7c3aed", highlightthickness=2,
                                  wrap="word")
        self.assert_box.grid(row=5, column=0, sticky="ew", padx=1, pady=(0,4))
        # Pre-fill with the exact strings visible in the popup screenshot
        self.assert_box.insert("end",
            "Open Enrollment is Currently in Session!\n"
            "The Open Enrollment period will end on May 31, 2026.\n"
            "All changes will be effective 07/01/2026."
        )
        tk.Label(card,
                 text="Tip: Each line = one independent check. Partial matches are fine.",
                 font=("Segoe UI", 8), bg=CARD, fg="#64748b"
                 ).grid(row=6, column=0, sticky="w", pady=(0,4))

        # Buttons
        br = tk.Frame(self.root, bg=BG)
        br.pack(fill="x", padx=20, pady=10)
        self.run_btn = tk.Button(br, text="Run Tests",
                                 font=("Segoe UI", 11, "bold"),
                                 bg=GREEN, fg="white", relief="flat",
                                 activebackground="#16a34a", cursor="hand2",
                                 padx=24, pady=10, command=self._start)
        self.run_btn.pack(side="left")
        tk.Button(br, text="Clear Log", font=("Segoe UI", 10),
                  bg=CARD, fg=MUTED, relief="flat",
                  activebackground=BORDER, cursor="hand2",
                  padx=16, pady=10, command=self._clear).pack(side="left", padx=(10,0))

        # Progress
        self.prog_var = tk.DoubleVar()
        sty = ttk.Style(); sty.theme_use("clam")
        sty.configure("P.Horizontal.TProgressbar",
                      troughcolor=CARD, background=ACCENT,
                      darkcolor=ACCENT, lightcolor=ACCENT, bordercolor=BG)
        ttk.Progressbar(self.root, variable=self.prog_var, maximum=100,
                        style="P.Horizontal.TProgressbar"
                        ).pack(fill="x", padx=20, pady=(0,4))
        self.status_lbl = tk.Label(self.root,
                                   text="Ready — select an Excel file to begin",
                                   font=("Segoe UI", 9), bg=BG, fg=MUTED)
        self.status_lbl.pack(anchor="w", padx=22)

        # Log
        lf = tk.Frame(self.root, bg=BG)
        lf.pack(fill="both", expand=True, padx=20, pady=(6,16))
        tk.Label(lf, text="TEST LOG", font=("Segoe UI", 8, "bold"),
                 bg=BG, fg=MUTED).pack(anchor="w")
        self.log = scrolledtext.ScrolledText(lf, font=("Consolas", 9),
                                             bg="#0f0f1a", fg=TEXT,
                                             relief="flat", state="disabled",
                                             highlightbackground=BORDER,
                                             highlightthickness=1)
        self.log.pack(fill="both", expand=True, pady=(4,0))
        self.log.tag_config("pass",   foreground="#22c55e")
        self.log.tag_config("fail",   foreground="#ef4444")
        self.log.tag_config("info",   foreground="#94a3b8")
        self.log.tag_config("header", foreground="#7c3aed",
                             font=("Consolas", 9, "bold"))

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _browse(self):
        p = filedialog.askopenfilename(
            title="Select Excel Test Records",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")])
        if p:
            self.excel_path.set(p)
            self._status(f"Loaded: {os.path.basename(p)}")

    def _log(self, msg, tag="info"):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        self.prog_var.set(0)
        self._status("Log cleared.")

    def _status(self, msg, kind="info"):
        c = {"info": "#94a3b8", "pass": "#22c55e", "fail": "#ef4444"}
        self.status_lbl.config(text=msg, fg=c.get(kind, "#94a3b8"))

    def _assert_lines(self):
        return [l.strip() for l in
                self.assert_box.get("1.0", "end").strip().splitlines() if l.strip()]

    def _start(self):
        if self.running: return
        if not self.excel_path.get():
            messagebox.showwarning("No File", "Please select the Excel file first."); return
        if not self._assert_lines():
            messagebox.showwarning("No Assertions", "Please enter expected popup text."); return
        self.running = True
        self.run_btn.config(state="disabled", text="Running...")
        threading.Thread(target=self._run, daemon=True).start()

    # ── Test runner ───────────────────────────────────────────────────────────
    def _run(self):
        passed, failed = [], []

        # Load Excel — only Username + Password
        try:
            df = pd.read_excel(self.excel_path.get())
            df.columns = [str(c).strip() for c in df.columns]
            cmap = {c.lower(): c for c in df.columns}
            uc = cmap.get("username")
            pc = cmap.get("password")
            if not uc or not pc:
                raise ValueError(
                    f"Need 'Username' and 'Password' columns.\nFound: {list(df.columns)}")
            records = (df[[uc, pc]].dropna()
                       .rename(columns={uc: "Username", pc: "Password"})
                       .to_dict("records"))
        except Exception as e:
            self.root.after(0, lambda: self._log(f"Failed to read Excel: {e}", "fail"))
            self.root.after(0, self._reset)
            return

        lines = self._assert_lines()
        total = len(records)

        self.root.after(0, lambda: self._log("=" * 65, "header"))
        self.root.after(0, lambda: self._log(
            f"  Starting {total} test(s)  |  {len(lines)} assertion(s)  |  {time.strftime('%H:%M:%S')}", "header"))
        self.root.after(0, lambda: self._log("=" * 65, "header"))

        opts = webdriver.ChromeOptions()
        opts.add_argument("--start-maximized")
        try:
            driver = webdriver.Chrome(options=opts)
        except Exception as e:
            self.root.after(0, lambda: self._log(f"Chrome failed: {e}", "fail"))
            self.root.after(0, self._reset)
            return

        wait = WebDriverWait(driver, 15)

        for i, row in enumerate(records, 1):
            user = str(row["Username"]).strip()
            pwd  = str(row["Password"]).strip()
            lbl  = f"[{i:02d}/{total}]  {user}"

            try:
                driver.get(self.portal_url.get().strip())

                # Username
                uf = wait.until(EC.presence_of_element_located((By.XPATH,
                    "//input[@type='text' and ("
                    "@placeholder='Username' or @name='username' or "
                    "@id='username' or @id='Username' or "
                    "contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                    "'abcdefghijklmnopqrstuvwxyz'),'user'))]")))
                uf.clear(); uf.send_keys(user)

                # Password
                pf = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@type='password']")))
                pf.clear(); pf.send_keys(pwd)

                # Log In
                wait.until(EC.element_to_be_clickable((By.XPATH,
                    "//input[@value='Log In'] | "
                    "//button[normalize-space()='Log In']"))).click()

                # Wait for popup by its real ID from DevTools
                popup = wait.until(EC.visibility_of_element_located(
                    (By.ID, "PopUp_LblMsg")))

                # Get text — try .text first, fall back to innerText via JS
                popup_text = popup.text.strip()
                if not popup_text:
                    popup_text = driver.execute_script(
                        "return arguments[0].innerText;", popup).strip()

                # Check every assertion line
                missing = [ln for ln in lines if ln not in popup_text]
                if missing:
                    raise AssertionError(
                        "Missing from popup:\n" +
                        "\n".join(f"        • {m}" for m in missing))

                self.root.after(0, lambda lb=lbl:
                                self._log(f"PASS  {lb}", "pass"))
                passed.append(user)

            except TimeoutException:
                r = "Timed out — login failed or popup not visible"
                self.root.after(0, lambda lb=lbl, r=r:
                                self._log(f"FAIL  {lb}\n        {r}", "fail"))
                failed.append({"Username": user, "Reason": r})

            except AssertionError as e:
                self.root.after(0, lambda lb=lbl, e=e:
                                self._log(f"FAIL  {lb}\n{e}", "fail"))
                failed.append({"Username": user, "Reason": str(e)})

            except Exception as e:
                r = f"{type(e).__name__}: {str(e)[:100]}"
                self.root.after(0, lambda lb=lbl, r=r:
                                self._log(f"FAIL  {lb}\n        {r}", "fail"))
                failed.append({"Username": user, "Reason": r})

            self.root.after(0, lambda p=(i/total*100): self.prog_var.set(p))
            self.root.after(0, lambda i=i, t=total, p=len(passed), f=len(failed):
                            self._status(f"Running {i}/{t}  —  {p} passed  {f} failed"))
            time.sleep(1)

        driver.quit()

        self.root.after(0, lambda: self._log("\n" + "=" * 65, "header"))
        self.root.after(0, lambda: self._log("  RESULTS SUMMARY", "header"))
        self.root.after(0, lambda: self._log("=" * 65, "header"))
        self.root.after(0, lambda: self._log(f"  Passed : {len(passed)} / {total}", "pass"))
        self.root.after(0, lambda: self._log(
            f"  Failed : {len(failed)} / {total}", "fail" if failed else "info"))

        if failed:
            self.root.after(0, lambda: self._log("\n  FAILED RECORDS:", "fail"))
            for f in failed:
                self.root.after(0, lambda f=f: self._log(
                    f"  Username : {f['Username']}\n"
                    f"  Reason   : {f['Reason']}\n"
                    f"  {'-'*60}", "fail"))
        else:
            self.root.after(0, lambda: self._log("  All records passed!", "pass"))

        self.root.after(0, lambda: self._status(
            f"Done — {len(passed)} passed, {len(failed)} failed",
            "fail" if failed else "pass"))
        self.root.after(0, self._reset)

    def _reset(self):
        self.running = False
        self.run_btn.config(state="normal", text="Run Tests")


if __name__ == "__main__":
    root = tk.Tk()
    TestRunnerApp(root)
    root.mainloop()
