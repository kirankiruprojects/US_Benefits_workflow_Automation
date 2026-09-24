"""
Workflow Capture Tool
----------------------
A small always-on-top toolbar for QA testers. Walk through your app's
workflow as usual. Whenever you hit a screen worth documenting, either
click Capture or press the capture hotkey (Ctrl+Shift+K) - the tool
grabs a screenshot of whichever monitor your mouse is currently on and
drops it into a Word document.

Buttons:
    Start   -> fill in the cover-page form (client, team member, team,
               date, purpose of review), choose where to save the .docx,
               begin. This builds the report's front page - the same
               one filled in for every new test session - with Total
               Issues Identified and Issues Log (Page No.) left for the
               tool to back-fill automatically at End.
    Capture -> screenshot the monitor the mouse cursor is currently on,
               insert into the doc. If you filled in Issue and/or Note,
               they're added ABOVE the image:
                   Issue -> red text, tagged with the page it lands on,
                            and also logged in an "Issues Summary" block
                            at the TOP of the document (with a live page
                            number, since it's a Word field)
                   Note  -> dark green text
    Undo    -> removes the most recently captured step (image, issue
               and/or note text, and its Issues Summary entry) in case
               a capture was taken by mistake.
    Pause Hotkey -> toggles the global capture hotkey on/off, for when
               you need to type "k" freely in the app you're testing
               without triggering captures.
    End     -> a short form asks for Total Issues Pending, Total Issues
               Resolved, and Date of Issues Resolved (Total Issues
               Identified and Issues Log (Page No.) are already known,
               so those fill in on their own); everything is written
               into the cover page, then the document is finalized

ALSO: once a session is running, pressing Ctrl+Shift+K anywhere on the
keyboard (not just while this window is focused) triggers a capture,
same as clicking the Capture button - so you can keep both hands on
the app you're testing and just tap the hotkey when you hit a screen
worth saving. Typing inside the Issue / Note boxes themselves is
excluded so it doesn't spam captures while you're writing a note, and
the "Pause Hotkey" button lets you disable it entirely while you work.

Requirements (install once):
    pip install pillow python-docx screeninfo keyboard pywin32

Notes:
    - Multi-monitor capture uses screeninfo to find which monitor the
      mouse cursor is on, then grabs just that monitor. Cursor position
      is read via the Windows GetCursorPos API, so this capture logic
      is Windows-only as written.
    - The page numbers (both next to each Issue, and in the Issues
      Summary at the top) are Word fields. During the session they show
      a placeholder until Word lays the document out. When you click
      End, this script opens the saved file in Word itself (via COM
      automation, using pywin32) to force a real update + repagination,
      then saves and closes it - so the page numbers you see when you
      next open the file are the actual final ones. This needs Word
      installed on this machine. If that automation step fails for any
      reason, you'll get a warning and can fix it by hand: open the
      .docx, press Ctrl+A then F9, and save again.
    - The capture hotkey is Ctrl+Shift+K, working globally (not just
      while this window is focused) so you can trigger it while testing
      another app, without colliding with normal typing (including
      typing the letter "k") in the app under test. Typing inside the
      Issue/Note boxes themselves is also excluded so it doesn't fire
      while you're writing a note. Use the "Pause Hotkey" button to
      disable/re-enable it entirely at any time.
    - Undo removes only the most recently captured step. There is no
      multi-level undo - if you need to remove more than one step,
      click Undo once per step, most recent first.

Run:
    python workflow_capture_tool.py
"""

import io
import ctypes
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

import keyboard
from PIL import ImageGrab
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from screeninfo import get_monitors

ISSUE_COLOR = RGBColor(0xC0, 0x00, 0x00)     # red
NOTE_COLOR = RGBColor(0x00, 0x64, 0x00)      # dark green
BRAND_COLOR = RGBColor(0x1F, 0x4E, 0x79)     # report title / accent blue
ISSUE_COLOR_HEX = "#C00000"
NOTE_COLOR_HEX = "#006400"

# ---- GUI palette (professional look, same widgets/logic underneath) ----
UI_BG = "#F4F6F8"
UI_PANEL_BG = "#FFFFFF"
UI_BORDER = "#D7DCE1"
UI_TEXT = "#1F2A37"
UI_MUTED = "#6B7280"
UI_ACCENT = "#1F4E79"
UI_START = "#1F4E79"
UI_CAPTURE = "#1E824C"
UI_END = "#B3261E"
UI_UNDO = "#B45309"
UI_PAUSE = "#374151"
FONT_FAMILY = "Segoe UI"

CAPTURE_HOTKEY = "k"

COVER_FIELDS = [
    "Client Name",
    "Team Member",
    "Date of Testing",
    "Total Issues Identified",
    "Issues Log (Page No.)",
    "Total Issues Pending",
    "Total Issues Resolved",
    "Date of Issues Resolved",
    "Purpose of Review",
]


# ---------- monitor / cursor helpers ----------

def get_cursor_pos():
    """Current mouse cursor position in virtual-screen coordinates (Windows)."""
    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def get_monitor_under_cursor():
    """The screeninfo Monitor object that currently contains the cursor."""
    monitors = get_monitors()
    x, y = get_cursor_pos()
    for m in monitors:
        if m.x <= x < m.x + m.width and m.y <= y < m.y + m.height:
            return m
    for m in monitors:
        if getattr(m, "is_primary", False):
            return m
    return monitors[0]


def grab_monitor_under_cursor():
    mon = get_monitor_under_cursor()
    bbox = (mon.x, mon.y, mon.x + mon.width, mon.y + mon.height)
    return ImageGrab.grab(bbox=bbox, all_screens=True)


# ---------- Word field / bookmark helpers ----------

def add_field(paragraph, instr_text, placeholder="1", color=None, bold=False):
    """Insert a live Word field (e.g. PAGE, PAGEREF x \\h) into paragraph."""
    run = paragraph.add_run()
    run.bold = bold
    if color:
        run.font.color.rgb = color

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instr_text

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    cached = OxmlElement("w:t")
    cached.text = placeholder

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(cached)
    run._r.append(fld_end)
    return run


def add_bookmark(paragraph, bookmark_name, bookmark_id):
    """Wrap paragraph in a named bookmark so it can be targeted by PAGEREF."""
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), bookmark_name)

    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))

    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def set_update_fields_on_open(document):
    """Flag the doc so Word recalculates fields (PAGE/PAGEREF) as soon as it opens."""
    settings = document.settings.element
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.append(update_fields)


def refresh_fields_via_word(path):
    """
    Open the saved .docx in actual Word (COM automation), force it to
    repaginate and recompute every field (PAGE/PAGEREF included), then
    save + close. This is what makes the page numbers in the Issues
    Summary correct without the user having to press F9 by hand.

    Requires pywin32 and a local install of Microsoft Word. Raises on
    failure so the caller can warn the user and suggest the manual fix.
    """
    import win32com.client as win32

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path)
        try:
            doc.Repaginate()
            doc.Fields.Update()
            doc.Repaginate()
            doc.Fields.Update()  # second pass: PAGEREF needs final pagination settled
            doc.Save()
        finally:
            doc.Close(False)
    finally:
        word.Quit()


def add_cover_page(document, client_name, team_member, team_name, date_of_testing, purpose):
    """
    Build the front page (title + fillable info table) matching the
    team's report template, and return a dict of {field label: cell}
    for the cells that get updated later (issue counts / page list).
    """
    title_p = document.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("Client Setup Review & Testing Report")
    title_run.bold = True
    title_run.font.size = Pt(22)
    title_run.font.color.rgb = BRAND_COLOR

    subtitle_p = document.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle_p.add_run(team_name)
    subtitle_run.underline = True
    subtitle_run.font.size = Pt(12)
    subtitle_run.font.color.rgb = BRAND_COLOR

    document.add_paragraph("")

    values = {
        "Client Name": client_name,
        "Team Member": team_member,
        "Date of Testing": date_of_testing,
        "Total Issues Identified": "",
        "Issues Log (Page No.)": "",
        "Total Issues Pending": "",
        "Total Issues Resolved": "",
        "Date of Issues Resolved": "",
        "Purpose of Review": purpose,
    }

    table = document.add_table(rows=0, cols=2)
    table.style = "Table Grid"

    cover_cells = {}
    for label in COVER_FIELDS:
        row = table.add_row().cells
        row[0].width = Inches(2.1)
        row[1].width = Inches(4.1)
        row[0].text = label
        row[0].paragraphs[0].runs[0].bold = True
        row[1].text = values[label]
        cover_cells[label] = row[1]

    document.add_page_break()
    return cover_cells


def fill_cover_summary(cover_cells, total_issues, bookmark_names,
                        pending="", resolved="", date_resolved=""):
    """After testing, back-fill the whole issues block on the cover page:
    the counts/pages the tool already knows automatically, plus whatever
    pending/resolved/date-resolved values were entered on End."""
    if not cover_cells:
        return

    count_cell = cover_cells.get("Total Issues Identified")
    if count_cell is not None:
        count_cell.text = str(total_issues)

    pages_cell = cover_cells.get("Issues Log (Page No.)")
    if pages_cell is not None:
        pages_cell.text = ""
        p = pages_cell.paragraphs[0]
        if not bookmark_names:
            p.add_run("None")
        else:
            for i, name in enumerate(bookmark_names):
                if i > 0:
                    p.add_run(", ")
                add_field(p, f"PAGEREF {name} \\h")

    pending_cell = cover_cells.get("Total Issues Pending")
    if pending_cell is not None:
        pending_cell.text = pending

    resolved_cell = cover_cells.get("Total Issues Resolved")
    if resolved_cell is not None:
        resolved_cell.text = resolved

    date_resolved_cell = cover_cells.get("Date of Issues Resolved")
    if date_resolved_cell is not None:
        date_resolved_cell.text = date_resolved


class EndSessionDialog(tk.Toplevel):
    """Modal form shown on End to collect the fields that can only be
    judged once testing is done, so the whole Issues block on the cover
    page gets filled in automatically instead of being left blank."""

    def __init__(self, parent, total_identified):
        super().__init__(parent)
        self.result = {"pending": "", "resolved": "", "date_resolved": ""}
        self.title("Finish Testing Session")
        self.configure(bg=UI_PANEL_BG)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        header = tk.Frame(self, bg=UI_ACCENT)
        header.pack(fill="x")
        tk.Label(
            header, text="Finish Up: Issues Summary", bg=UI_ACCENT, fg="white",
            font=(FONT_FAMILY, 12, "bold"), padx=14, pady=10,
        ).pack(anchor="w")

        body = tk.Frame(self, bg=UI_PANEL_BG)
        body.pack(fill="both", expand=True, padx=4, pady=4)
        pad = {"padx": 14, "pady": (10, 0)}

        tk.Label(
            body, text=f"Total Issues Identified (auto): {total_identified}",
            bg=UI_PANEL_BG, fg=UI_MUTED, font=(FONT_FAMILY, 9), anchor="w",
        ).pack(fill="x", **pad)

        self.vars = {}

        def add_entry(label_text, key, default=""):
            tk.Label(
                body, text=label_text, bg=UI_PANEL_BG, fg=UI_TEXT,
                font=(FONT_FAMILY, 9, "bold"), anchor="w",
            ).pack(fill="x", **pad)
            var = tk.StringVar(value=default)
            entry = tk.Entry(
                body, textvariable=var, font=(FONT_FAMILY, 10),
                relief="solid", bd=1, highlightthickness=0,
            )
            entry.pack(fill="x", padx=14, ipady=4)
            self.vars[key] = var
            return entry

        first_entry = add_entry("Total Issues Pending", "pending")
        add_entry("Total Issues Resolved", "resolved")
        add_entry("Date of Issues Resolved", "date_resolved", default=datetime.now().strftime("%Y-%m-%d"))

        btn_row = tk.Frame(body, bg=UI_PANEL_BG)
        btn_row.pack(fill="x", padx=14, pady=12)

        skip_btn = tk.Button(
            btn_row, text="Skip", command=self._on_skip,
            bg="#E5E7EB", fg=UI_TEXT, activebackground="#D1D5DB",
            relief="flat", font=(FONT_FAMILY, 9), padx=10, pady=5,
        )
        skip_btn.pack(side="right", padx=(6, 0))

        finish_btn = tk.Button(
            btn_row, text="Finish & Save", command=self._on_finish,
            bg=UI_END, fg="white", activebackground="#8C1E18",
            relief="flat", font=(FONT_FAMILY, 9, "bold"), padx=10, pady=5,
        )
        finish_btn.pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._on_skip)
        first_entry.focus_set()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() - 260}+{parent.winfo_rooty() + 20}")
        self.wait_window(self)

    def _on_finish(self):
        self.result = {
            "pending": self.vars["pending"].get().strip(),
            "resolved": self.vars["resolved"].get().strip(),
            "date_resolved": self.vars["date_resolved"].get().strip(),
        }
        self.destroy()

    def _on_skip(self):
        self.result = {"pending": "", "resolved": "", "date_resolved": ""}
        self.destroy()


class StartSessionDialog(tk.Toplevel):
    """Modal form collecting the cover-page details for a new test session."""

    def __init__(self, parent):
        super().__init__(parent)
        self.result = None
        self.title("New Testing Session")
        self.configure(bg=UI_PANEL_BG)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        pad = {"padx": 14, "pady": (10, 0)}
        header = tk.Frame(self, bg=UI_ACCENT)
        header.pack(fill="x")
        tk.Label(
            header, text="Client Setup Review & Testing Report",
            bg=UI_ACCENT, fg="white", font=(FONT_FAMILY, 12, "bold"),
            padx=14, pady=10,
        ).pack(anchor="w")

        body = tk.Frame(self, bg=UI_PANEL_BG)
        body.pack(fill="both", expand=True, padx=4, pady=4)

        self.vars = {}

        def add_entry(label_text, key, default=""):
            tk.Label(
                body, text=label_text, bg=UI_PANEL_BG, fg=UI_TEXT,
                font=(FONT_FAMILY, 9, "bold"), anchor="w",
            ).pack(fill="x", **pad)
            var = tk.StringVar(value=default)
            entry = tk.Entry(
                body, textvariable=var, font=(FONT_FAMILY, 10),
                relief="solid", bd=1, highlightthickness=0,
            )
            entry.pack(fill="x", padx=14, ipady=4)
            self.vars[key] = var
            return entry

        first_entry = add_entry("Workflow / Test Name", "workflow_name")
        add_entry("Client Name", "client_name")
        add_entry("Team Member", "team_member")
        add_entry("Team / Department", "team_name", default="Systems Configuration Team")
        add_entry("Date of Testing", "date_of_testing", default=datetime.now().strftime("%Y-%m-%d"))

        tk.Label(
            body, text="Purpose of Review", bg=UI_PANEL_BG, fg=UI_TEXT,
            font=(FONT_FAMILY, 9, "bold"), anchor="w",
        ).pack(fill="x", **pad)
        self.purpose_text = tk.Text(
            body, height=3, font=(FONT_FAMILY, 10), wrap="word",
            relief="solid", bd=1, highlightthickness=0,
        )
        self.purpose_text.pack(fill="x", padx=14, pady=(0, 4))

        btn_row = tk.Frame(body, bg=UI_PANEL_BG)
        btn_row.pack(fill="x", padx=14, pady=12)

        cancel_btn = tk.Button(
            btn_row, text="Cancel", command=self._on_cancel,
            bg="#E5E7EB", fg=UI_TEXT, activebackground="#D1D5DB",
            relief="flat", font=(FONT_FAMILY, 9), padx=10, pady=5,
        )
        cancel_btn.pack(side="right", padx=(6, 0))

        begin_btn = tk.Button(
            btn_row, text="Begin Session", command=self._on_begin,
            bg=UI_CAPTURE, fg="white", activebackground="#166638",
            relief="flat", font=(FONT_FAMILY, 9, "bold"), padx=10, pady=5,
        )
        begin_btn.pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        first_entry.focus_set()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() - 260}+{parent.winfo_rooty() + 20}")
        self.wait_window(self)

    def _on_begin(self):
        workflow_name = self.vars["workflow_name"].get().strip()
        if not workflow_name:
            messagebox.showwarning("Missing info", "Please enter a workflow / test name.", parent=self)
            return
        self.result = {
            "workflow_name": workflow_name,
            "client_name": self.vars["client_name"].get().strip(),
            "team_member": self.vars["team_member"].get().strip(),
            "team_name": self.vars["team_name"].get().strip() or "Systems Configuration Team",
            "date_of_testing": self.vars["date_of_testing"].get().strip(),
            "purpose": self.purpose_text.get("1.0", "end").strip(),
        }
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


class WorkflowCaptureTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Workflow Capture")
        self.root.attributes("-topmost", True)
        self.root.resizable(False, False)
        self.root.configure(bg=UI_BG)

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=UI_BG)
        style.configure("Panel.TFrame", background=UI_PANEL_BG)
        style.configure("TLabel", background=UI_BG, foreground=UI_TEXT, font=(FONT_FAMILY, 9))
        style.configure("Panel.TLabel", background=UI_PANEL_BG, foreground=UI_TEXT, font=(FONT_FAMILY, 9))
        style.configure("Muted.TLabel", background=UI_BG, foreground=UI_MUTED, font=(FONT_FAMILY, 8))
        style.configure("Header.TLabel", background=UI_ACCENT, foreground="white", font=(FONT_FAMILY, 11, "bold"))
        style.configure("Status.TLabel", background=UI_BG, foreground=UI_ACCENT, font=(FONT_FAMILY, 10, "bold"))

        # session state
        self.doc = None
        self.save_path = None
        self.workflow_name = None
        self.start_time = None
        self.step_count = 0
        self.bookmark_counter = 0
        self.bookmark_names = []
        self.last_summary_element = None
        self.cover_cells = {}

        # undo tracking - the elements added by the most recent capture_step()
        self.last_step_elements = []          # body-level <w:p> elements from the last step
        self.last_step_had_issue = False
        self.last_step_summary_element = None  # the mirrored Issues Summary <w:p>, if any

        # hotkey / focus tracking
        self.capture_in_progress = False
        self.typing_active = False
        self.hotkey_enabled = True

        self._build_ui()
        self._position_top_right()

        keyboard.add_hotkey(CAPTURE_HOTKEY, self._hotkey_capture, suppress=False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- UI ----------

    def _build_ui(self):
        outer = ttk.Frame(self.root, style="TFrame")
        outer.grid(row=0, column=0, sticky="nsew")

        # header banner
        header = ttk.Frame(outer, style="TFrame")
        header.pack(fill="x")
        header_inner = tk.Frame(header, bg=UI_ACCENT)
        header_inner.pack(fill="x")
        tk.Label(
            header_inner, text="Workflow Capture", bg=UI_ACCENT, fg="white",
            font=(FONT_FAMILY, 12, "bold"), padx=12, pady=8,
        ).pack(anchor="w")

        # card-style panel holding everything else
        panel = tk.Frame(outer, bg=UI_PANEL_BG, highlightbackground=UI_BORDER, highlightthickness=1)
        panel.pack(fill="both", expand=True, padx=10, pady=10)
        pad = {"padx": 12, "pady": (8, 0)}

        self.status_var = tk.StringVar(value="Not started")
        ttk.Label(panel, textvariable=self.status_var, style="Status.TLabel", background=UI_PANEL_BG).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 4)
        )

        btn_frame = tk.Frame(panel, bg=UI_PANEL_BG)
        btn_frame.grid(row=1, column=0, columnspan=3, padx=12, pady=6, sticky="w")

        def make_button(parent, text, color, hover, command, width=10):
            return tk.Button(
                parent, text=text, command=command,
                bg=color, fg="white", activebackground=hover, activeforeground="white",
                relief="flat", font=(FONT_FAMILY, 9, "bold"), width=width, pady=6, bd=0,
                disabledforeground="#E5E7EB",
            )

        self.start_btn = make_button(btn_frame, "Start", UI_START, "#163C5C", self.start_session)
        self.start_btn.grid(row=0, column=0, padx=(0, 6))

        self.capture_btn = make_button(btn_frame, "Capture", UI_CAPTURE, "#166638", self.capture_step)
        self.capture_btn.grid(row=0, column=1, padx=6)
        self.capture_btn.config(state="disabled", bg="#A7D7BB")

        self.end_btn = make_button(btn_frame, "End", UI_END, "#8C1E18", self.end_session)
        self.end_btn.grid(row=0, column=2, padx=(6, 0))
        self.end_btn.config(state="disabled", bg="#E8AFAA")

        # second row: Undo + Pause Hotkey
        btn_frame2 = tk.Frame(panel, bg=UI_PANEL_BG)
        btn_frame2.grid(row=2, column=0, columnspan=3, padx=12, pady=(0, 6), sticky="w")

        self.undo_btn = make_button(
            btn_frame2, "Undo Last", UI_UNDO, "#8A4109", self.undo_last_capture, width=10
        )
        self.undo_btn.grid(row=0, column=0, padx=(0, 6))
        self.undo_btn.config(state="disabled", bg="#E6C199")

        self.pause_btn = make_button(
            btn_frame2, "Pause Hotkey", UI_PAUSE, "#1F2937", self.toggle_hotkey, width=13
        )
        self.pause_btn.grid(row=0, column=1, padx=(6, 0))

        # keep a reference to each button's "enabled" color so state toggles keep the palette
        self._btn_colors = {
            self.capture_btn: (UI_CAPTURE, "#A7D7BB"),
            self.end_btn: (UI_END, "#E8AFAA"),
            self.start_btn: (UI_START, "#9FB6C9"),
            self.undo_btn: (UI_UNDO, "#E6C199"),
        }

        sep1 = tk.Frame(panel, bg=UI_BORDER, height=1)
        sep1.grid(row=3, column=0, columnspan=3, sticky="ew", padx=12, pady=10)

        ttk.Label(panel, text="Issue", style="Panel.TLabel", background=UI_PANEL_BG,
                  font=(FONT_FAMILY, 9, "bold"), foreground=ISSUE_COLOR_HEX).grid(
            row=4, column=0, columnspan=3, sticky="w", **pad
        )
        ttk.Label(panel, text="Logged above the screenshot in red, and listed with its page number at the top of the report.",
                  style="Muted.TLabel", background=UI_PANEL_BG).grid(
            row=5, column=0, columnspan=3, sticky="w", padx=12
        )
        self.issue_text = tk.Text(
            panel, height=2, width=42, wrap="word", fg=ISSUE_COLOR_HEX,
            font=(FONT_FAMILY, 10), relief="solid", bd=1, highlightthickness=0,
        )
        self.issue_text.grid(row=6, column=0, columnspan=3, padx=12, pady=(4, 0), sticky="ew")
        self.issue_text.bind("<FocusIn>", self._on_note_focus_in)
        self.issue_text.bind("<FocusOut>", self._on_note_focus_out)

        ttk.Label(panel, text="Note", style="Panel.TLabel", background=UI_PANEL_BG,
                  font=(FONT_FAMILY, 9, "bold"), foreground=NOTE_COLOR_HEX).grid(
            row=7, column=0, columnspan=3, sticky="w", **pad
        )
        self.note_text = tk.Text(
            panel, height=2, width=42, wrap="word", fg=NOTE_COLOR_HEX,
            font=(FONT_FAMILY, 10), relief="solid", bd=1, highlightthickness=0,
        )
        self.note_text.grid(row=8, column=0, columnspan=3, padx=12, pady=(4, 0), sticky="ew")
        self.note_text.bind("<FocusIn>", self._on_note_focus_in)
        self.note_text.bind("<FocusOut>", self._on_note_focus_out)

        self.count_var = tk.StringVar(value="")
        ttk.Label(panel, textvariable=self.count_var, style="Panel.TLabel",
                  background=UI_PANEL_BG, foreground=UI_MUTED).grid(
            row=9, column=0, columnspan=3, sticky="w", padx=12, pady=(10, 0)
        )

        hint = ttk.Label(
            panel,
            text='Tip: once running, press K anywhere to capture too.',
            style="Muted.TLabel", background=UI_PANEL_BG,
        )
        hint.grid(row=10, column=0, columnspan=3, sticky="w", padx=12, pady=(2, 12))

    def _set_btn_enabled(self, btn, enabled):
        enabled_color, disabled_color = self._btn_colors[btn]
        btn.config(state="normal" if enabled else "disabled",
                   bg=enabled_color if enabled else disabled_color)

    def _position_top_right(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        screen_w = self.root.winfo_screenwidth()
        self.root.geometry(f"+{screen_w - w - 30}+30")

    # ---------- focus / hotkey plumbing ----------

    def _on_note_focus_in(self, _event):
        self.typing_active = True

    def _on_note_focus_out(self, _event):
        self.typing_active = False

    def toggle_hotkey(self):
        self.hotkey_enabled = not self.hotkey_enabled
        self.pause_btn.config(text="Resume Hotkey" if not self.hotkey_enabled else "Pause Hotkey")

    def _hotkey_capture(self):
        # runs on keyboard library's own thread - only touch Tk from .after()
        if not self.hotkey_enabled or self.doc is None or self.capture_in_progress or self.typing_active:
            return
        self.root.after(0, self.capture_step)

    def _on_close(self):
        keyboard.unhook_all()
        self.root.destroy()

    # ---------- Actions ----------

    def start_session(self):
        dialog = StartSessionDialog(self.root)
        info = dialog.result
        if not info:
            return

        default_name = f"{info['workflow_name'].replace(' ', '_')}_test_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        path = filedialog.asksaveasfilename(
            title="Save test document as",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word document", "*.docx")],
        )
        if not path:
            return

        self.workflow_name = info["workflow_name"]
        self.save_path = path
        self.start_time = datetime.now()
        self.step_count = 0
        self.bookmark_counter = 0
        self.bookmark_names = []
        self.last_step_elements = []
        self.last_step_had_issue = False
        self.last_step_summary_element = None

        self.doc = Document()
        set_update_fields_on_open(self.doc)

        self.cover_cells = add_cover_page(
            self.doc,
            client_name=info["client_name"],
            team_member=info["team_member"],
            team_name=info["team_name"],
            date_of_testing=info["date_of_testing"],
            purpose=info["purpose"],
        )

        self.doc.add_heading(f"Workflow Test: {self.workflow_name}", level=1)
        meta = self.doc.add_paragraph()
        meta.add_run(f"Started: {self.start_time.strftime('%Y-%m-%d %H:%M:%S')}").italic = True

        summary_heading = self.doc.add_paragraph()
        heading_run = summary_heading.add_run("Issues Summary (page numbers)")
        heading_run.bold = True
        heading_run.underline = True
        self.last_summary_element = summary_heading._p

        self.doc.add_paragraph("")  # spacing before steps begin

        self.status_var.set(f"Recording: {self.workflow_name}")
        self.count_var.set("0 steps captured")

        self._set_btn_enabled(self.start_btn, False)
        self._set_btn_enabled(self.capture_btn, True)
        self._set_btn_enabled(self.end_btn, True)
        self._set_btn_enabled(self.undo_btn, False)

    def capture_step(self):
        if self.doc is None or self.capture_in_progress:
            return
        self.capture_in_progress = True
        try:
            issue = self.issue_text.get("1.0", "end").strip()
            note = self.note_text.get("1.0", "end").strip()

            # hide the toolbar so it doesn't appear in its own screenshot
            self.root.withdraw()
            self.root.update()
            time.sleep(0.35)

            try:
                img = grab_monitor_under_cursor()
            finally:
                self.root.deiconify()
                self.root.attributes("-topmost", True)

            self.step_count += 1

            # reset per-step undo tracking
            step_elements = []
            step_had_issue = False
            step_summary_element = None

            if issue:
                self.bookmark_counter += 1
                bookmark_name = f"issue_{self.bookmark_counter}"
                self.bookmark_names.append(bookmark_name)

                # inline, at the actual location, tagged with its own page
                issue_para = self.doc.add_paragraph()
                add_bookmark(issue_para, bookmark_name, self.bookmark_counter)
                step_elements.append(issue_para._p)

                lead = issue_para.add_run("Issue (page ")
                lead.bold = True
                lead.font.color.rgb = ISSUE_COLOR
                add_field(issue_para, "PAGE", color=ISSUE_COLOR, bold=True)
                trail = issue_para.add_run(f"): {issue}")
                trail.bold = True
                trail.font.color.rgb = ISSUE_COLOR

                # mirrored into the Issues Summary block at the top
                summary_para = self.doc.add_paragraph()
                self.last_summary_element.addnext(summary_para._p)
                self.last_summary_element = summary_para._p

                bullet = summary_para.add_run("\u2022 Page ")
                bullet.font.color.rgb = ISSUE_COLOR
                add_field(summary_para, f"PAGEREF {bookmark_name} \\h", color=ISSUE_COLOR)
                tail = summary_para.add_run(f": {issue}")
                tail.font.color.rgb = ISSUE_COLOR

                step_had_issue = True
                step_summary_element = summary_para._p

            if note:
                note_para = self.doc.add_paragraph()
                n_lead = note_para.add_run("Note: ")
                n_lead.bold = True
                n_lead.font.color.rgb = NOTE_COLOR
                n_body = note_para.add_run(note)
                n_body.font.color.rgb = NOTE_COLOR
                step_elements.append(note_para._p)

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            self.doc.add_picture(buf, width=Inches(6.2))
            # add_picture appends a new paragraph holding the image as the
            # last element in the document body - grab it for undo tracking
            picture_para_el = self.doc.paragraphs[-1]._p
            step_elements.append(picture_para_el)

            spacer_para = self.doc.add_paragraph("")  # spacing between steps
            step_elements.append(spacer_para._p)

            # commit this step's undo info now that everything succeeded
            self.last_step_elements = step_elements
            self.last_step_had_issue = step_had_issue
            self.last_step_summary_element = step_summary_element

            self.issue_text.delete("1.0", "end")
            self.note_text.delete("1.0", "end")
            self.count_var.set(f"{self.step_count} step(s) captured")
            self._set_btn_enabled(self.undo_btn, True)

            # autosave after every capture so nothing is lost if the app is closed
            try:
                self.doc.save(self.save_path)
            except PermissionError:
                messagebox.showwarning(
                    "Can't save",
                    "The document is open in Word and locked for writing.\n"
                    "Close it in Word, then capture again.",
                )
        finally:
            self.capture_in_progress = False

    def undo_last_capture(self):
        if self.doc is None or self.step_count == 0 or not self.last_step_elements:
            return
        if not messagebox.askyesno("Undo", "Remove the most recently captured step?"):
            return

        # remove every element this step added (issue/note paragraphs,
        # the picture paragraph, and the trailing spacer)
        for el in self.last_step_elements:
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

        # remove its mirrored Issues Summary entry, if it had one
        if self.last_step_had_issue:
            self.bookmark_counter -= 1
            if self.bookmark_names:
                self.bookmark_names.pop()
            if self.last_step_summary_element is not None:
                summary_parent = self.last_step_summary_element.getparent()
                if summary_parent is not None:
                    summary_parent.remove(self.last_step_summary_element)
                # move the "last summary element" pointer back so the next
                # captured issue gets inserted after the right paragraph
                if self.bookmark_names:
                    self.last_summary_element = self.last_summary_element.getprevious()
                    if self.last_summary_element is None:
                        # fall back: re-find the heading paragraph
                        for p in self.doc.paragraphs:
                            if p.text.startswith("Issues Summary"):
                                self.last_summary_element = p._p
                                break

        self.step_count -= 1
        self.count_var.set(f"{self.step_count} step(s) captured")

        # this step is gone - there is nothing further back tracked to undo
        self.last_step_elements = []
        self.last_step_had_issue = False
        self.last_step_summary_element = None
        self._set_btn_enabled(self.undo_btn, False)

        try:
            self.doc.save(self.save_path)
        except PermissionError:
            messagebox.showwarning(
                "Can't save",
                "The document is open in Word and locked for writing.\n"
                "Close it in Word, then try Undo again.",
            )

    def end_session(self):
        if self.doc is None:
            return

        end_time = datetime.now()
        duration = end_time - self.start_time

        self.doc.add_heading("Test Completed", level=2)
        summary = self.doc.add_paragraph()
        summary.add_run(f"Ended: {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n").italic = True
        summary.add_run(f"Duration: {str(duration).split('.')[0]}\n").italic = True
        summary.add_run(f"Total steps captured: {self.step_count}").italic = True

        # back-fill the whole Issues block on the cover page: counts the
        # tool already knows automatically, plus pending/resolved/date
        # resolved collected now, since only a human can judge those
        end_info = EndSessionDialog(self.root, total_identified=self.bookmark_counter).result
        fill_cover_summary(
            self.cover_cells, self.bookmark_counter, self.bookmark_names,
            pending=end_info["pending"],
            resolved=end_info["resolved"],
            date_resolved=end_info["date_resolved"],
        )

        try:
            self.doc.save(self.save_path)
        except PermissionError:
            messagebox.showwarning(
                "Can't save",
                "The document is open in Word and locked for writing.\n"
                "Close it in Word and try End again.",
            )
            return

        # bake in real page numbers by having Word itself update the fields
        try:
            refresh_fields_via_word(self.save_path)
            messagebox.showinfo("Saved", f"Test document saved to:\n{self.save_path}")
        except Exception:
            messagebox.showwarning(
                "Saved (page numbers may need a manual refresh)",
                f"Test document saved to:\n{self.save_path}\n\n"
                "Couldn't auto-update the page numbers via Word (it may not be "
                "installed, or pywin32 isn't available). Open the file, press "
                "Ctrl+A then F9, and save again to fix them.",
            )

        # reset for a new session
        self.doc = None
        self.save_path = None
        self.cover_cells = {}
        self.last_step_elements = []
        self.last_step_had_issue = False
        self.last_step_summary_element = None
        self.status_var.set("Not started")
        self.count_var.set("")
        self.issue_text.delete("1.0", "end")
        self.note_text.delete("1.0", "end")
        self._set_btn_enabled(self.start_btn, True)
        self._set_btn_enabled(self.capture_btn, False)
        self._set_btn_enabled(self.end_btn, False)
        self._set_btn_enabled(self.undo_btn, False)


if __name__ == "__main__":
    root = tk.Tk()
    app = WorkflowCaptureTool(root)
    root.mainloop()