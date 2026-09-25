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
    pip install pillow python-docx screeninfo keyboard pywin32 language-tool-python

Notes:
    - Spelling/grammar check: the "Check Spelling/Grammar" button under the
      Issue and Note boxes uses language_tool_python, which needs a Java
      runtime installed on this machine. The first check in a session takes
      longer while it starts a local LanguageTool server; after that it's
      fast. If Java or the package isn't available, the button just shows a
      short message instead of failing - everything else keeps working.
    - Every page gets a brand-blue border (set once via set_page_border()
      right after the document is created), since a single section covers
      the whole report.
    - Dates are shown/entered as MM/DD/YYYY throughout (cover page fields
      and the Started/Ended timestamps, which also use a 12-hour AM/PM
      clock). The autosave filename still uses YYYYMMDD_HHMM internally
      since "/" isn't allowed in filenames.
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
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from screeninfo import get_monitors

try:
    import language_tool_python
    _GRAMMAR_AVAILABLE = True
except ImportError:
    _GRAMMAR_AVAILABLE = False

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


# ---------- internal hyperlink helpers (click an issue -> jump to it) ----------

def _hyperlink_rpr(color=None, bold=False, underline=True):
    rpr = OxmlElement("w:rPr")
    if bold:
        rpr.append(OxmlElement("w:b"))
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rpr.append(u)
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color.lstrip("#"))
        rpr.append(c)
    return rpr


def _text_run_el(text, color=None, bold=False, underline=True):
    """A plain <w:r> text run, built as a raw element (for use inside a hyperlink)."""
    run = OxmlElement("w:r")
    run.append(_hyperlink_rpr(color, bold, underline))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    return run


def _field_run_el(instr_text, placeholder="1", color=None, bold=False, underline=True):
    """A live-field <w:r> run (PAGEREF etc.), built as a raw element (for use inside a hyperlink)."""
    run = OxmlElement("w:r")
    run.append(_hyperlink_rpr(color, bold, underline))

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

    run.append(fld_begin)
    run.append(instr)
    run.append(fld_sep)
    run.append(cached)
    run.append(fld_end)
    return run


def add_issue_summary_hyperlink(paragraph, bookmark_name, issue_text, color):
    """Build the whole 'Page N: issue text' summary entry as ONE clickable
    hyperlink (the page number stays a live PAGEREF field) that jumps
    straight to that issue's bookmark in the body of the document."""
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), bookmark_name)
    hyperlink.set(qn("w:history"), "1")
    hyperlink.append(_text_run_el("Page ", color=color))
    hyperlink.append(_field_run_el(f"PAGEREF {bookmark_name} \\h", color=color))
    hyperlink.append(_text_run_el(f": {issue_text}", color=color))
    paragraph._p.append(hyperlink)
    return hyperlink


def add_next_issue_arrow(paragraph_p, target_bookmark, color, text=" \u25B8 Next Issue"):
    """Append a small clickable 'Next Issue' link to the END of an issue
    paragraph already in the document, jumping forward to the bookmark of
    the issue captured right after it."""
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), target_bookmark)
    hyperlink.set(qn("w:history"), "1")
    hyperlink.append(_text_run_el(text, color=color, bold=True))
    paragraph_p.append(hyperlink)
    return hyperlink


def set_update_fields_on_open(document):
    """Flag the doc so Word recalculates fields (PAGE/PAGEREF) as soon as it opens."""
    settings = document.settings.element
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.append(update_fields)


def set_page_border(document, color="1F4E79", sz=18, space=24, val="single"):
    """Add a border around every page of the document. Applies to every
    section, so it covers the whole doc even if more sections get added."""
    for section in document.sections:
        sect_pr = section._sectPr
        pg_borders = sect_pr.find(qn("w:pgBorders"))
        if pg_borders is None:
            pg_borders = OxmlElement("w:pgBorders")
            sect_pr.append(pg_borders)
        pg_borders.set(qn("w:offsetFrom"), "page")
        for edge in ("top", "left", "bottom", "right"):
            element = pg_borders.find(qn(f"w:{edge}"))
            if element is None:
                element = OxmlElement(f"w:{edge}")
                pg_borders.append(element)
            element.set(qn("w:val"), val)
            element.set(qn("w:sz"), str(sz))
            element.set(qn("w:space"), str(space))
            element.set(qn("w:color"), color)


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


def set_cell_border(cell, **edges):
    """Set specific border edges on a cell, e.g.
    set_cell_border(cell, top={"sz": 18, "val": "single", "color": "70AD47"})"""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge_name, edge_opts in edges.items():
        tag = f"w:{edge_name}"
        element = tc_borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tc_borders.append(element)
        for key, val in edge_opts.items():
            element.set(qn(f"w:{key}"), str(val))


# ---- cover table look (clean white card, blue values, green accent rules) ----
COVER_LABEL_FG = RGBColor(0x00, 0x00, 0x00)   # black, bold labels
COVER_VALUE_FG = RGBColor(0x1F, 0x6F, 0xC1)   # medium blue values
COVER_ACCENT_HEX = "70AD47"                    # green top/bottom accent rule
COVER_ROW_HEIGHT = Inches(0.6)                 # tall rows so the table reads as half the page


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
    table.autofit = False

    cover_cells = {}
    last_row_index = len(COVER_FIELDS) - 1
    for i, label in enumerate(COVER_FIELDS):
        table_row = table.add_row()
        table_row.height = COVER_ROW_HEIGHT
        table_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row = table_row.cells
        row[0].width = Inches(2.3)
        row[1].width = Inches(4.4)

        row[0].text = label
        label_run = row[0].paragraphs[0].runs[0]
        label_run.bold = True
        label_run.font.size = Pt(11)
        label_run.font.color.rgb = COVER_LABEL_FG

        row[1].text = values[label]
        if row[1].paragraphs[0].runs:
            value_run = row[1].paragraphs[0].runs[0]
            value_run.font.size = Pt(11)
            value_run.font.color.rgb = COVER_VALUE_FG

        for cell in row:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i == 0:
                set_cell_border(cell, top={"sz": 18, "val": "single", "color": COVER_ACCENT_HEX})
            if i == last_row_index:
                set_cell_border(cell, bottom={"sz": 18, "val": "single", "color": COVER_ACCENT_HEX})

        cover_cells[label] = row[1]

    document.add_page_break()
    return cover_cells


def _style_value_cell(cell):
    """Re-apply the cover table's value styling (blue, 11pt) to every run
    in a cell - needed after cell.text = ... replaces the runs."""
    for run in cell.paragraphs[0].runs:
        run.font.size = Pt(11)
        run.font.color.rgb = COVER_VALUE_FG


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
        _style_value_cell(count_cell)

    pages_cell = cover_cells.get("Issues Log (Page No.)")
    if pages_cell is not None:
        pages_cell.text = ""
        p = pages_cell.paragraphs[0]
        if not bookmark_names:
            p.add_run("None").font.color.rgb = COVER_VALUE_FG
        else:
            for i, name in enumerate(bookmark_names):
                if i > 0:
                    p.add_run(", ").font.color.rgb = COVER_VALUE_FG
                add_field(p, f"PAGEREF {name} \\h", color=COVER_VALUE_FG)

    pending_cell = cover_cells.get("Total Issues Pending")
    if pending_cell is not None:
        pending_cell.text = pending
        _style_value_cell(pending_cell)

    resolved_cell = cover_cells.get("Total Issues Resolved")
    if resolved_cell is not None:
        resolved_cell.text = resolved
        _style_value_cell(resolved_cell)

    date_resolved_cell = cover_cells.get("Date of Issues Resolved")
    if date_resolved_cell is not None:
        date_resolved_cell.text = date_resolved
        _style_value_cell(date_resolved_cell)


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
        add_entry("Date of Issues Resolved", "date_resolved", default=datetime.now().strftime("%m/%d/%Y"))

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
        add_entry("Date of Testing", "date_of_testing", default=datetime.now().strftime("%m/%d/%Y"))

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

        # "Next Issue" chain - the most recently captured issue's bookmark
        # and paragraph, so the NEXT issue captured can append a forward
        # arrow onto it
        self.last_issue_bookmark_name = None
        self.last_issue_paragraph_p = None

        # undo tracking - the elements added by the most recent capture_step()
        self.last_step_elements = []          # body-level <w:p> elements from the last step
        self.last_step_had_issue = False
        self.last_step_summary_element = None  # the mirrored Issues Summary <w:p>, if any
        self.last_step_prev_issue_bookmark = None   # "Next Issue" chain state to restore on undo
        self.last_step_prev_issue_paragraph_p = None
        self.last_step_arrow_element = None         # the arrow hyperlink this step added, if any

        # hotkey / focus tracking
        self.capture_in_progress = False
        self.typing_active = False
        self.hotkey_enabled = True

        # spelling/grammar checker - created lazily on first use since it's
        # slow to start; False means "tried and failed, don't retry"
        self._grammar_tool = None

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

        def make_check_button(parent, command):
            return tk.Button(
                parent, text="\u2713 Check Spelling/Grammar", command=command,
                bg="#EEF2F7", fg=UI_ACCENT, activebackground="#DCE6F1",
                activeforeground=UI_ACCENT, relief="flat",
                font=(FONT_FAMILY, 8, "bold"), padx=8, pady=3, bd=0,
            )

        self.issue_check_btn = make_check_button(
            panel, lambda: self._check_and_fix_text(self.issue_text, "Issue")
        )
        self.issue_check_btn.grid(row=7, column=0, columnspan=3, sticky="w", padx=12, pady=(4, 0))

        ttk.Label(panel, text="Note", style="Panel.TLabel", background=UI_PANEL_BG,
                  font=(FONT_FAMILY, 9, "bold"), foreground=NOTE_COLOR_HEX).grid(
            row=8, column=0, columnspan=3, sticky="w", **pad
        )
        self.note_text = tk.Text(
            panel, height=2, width=42, wrap="word", fg=NOTE_COLOR_HEX,
            font=(FONT_FAMILY, 10), relief="solid", bd=1, highlightthickness=0,
        )
        self.note_text.grid(row=9, column=0, columnspan=3, padx=12, pady=(4, 0), sticky="ew")
        self.note_text.bind("<FocusIn>", self._on_note_focus_in)
        self.note_text.bind("<FocusOut>", self._on_note_focus_out)

        self.note_check_btn = make_check_button(
            panel, lambda: self._check_and_fix_text(self.note_text, "Note")
        )
        self.note_check_btn.grid(row=10, column=0, columnspan=3, sticky="w", padx=12, pady=(4, 0))

        self.count_var = tk.StringVar(value="")
        ttk.Label(panel, textvariable=self.count_var, style="Panel.TLabel",
                  background=UI_PANEL_BG, foreground=UI_MUTED).grid(
            row=11, column=0, columnspan=3, sticky="w", padx=12, pady=(10, 0)
        )

        hint = ttk.Label(
            panel,
            text='Tip: once running, press K anywhere to capture too.',
            style="Muted.TLabel", background=UI_PANEL_BG,
        )
        hint.grid(row=12, column=0, columnspan=3, sticky="w", padx=12, pady=(2, 12))

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

    # ---------- spelling / grammar check ----------

    def _get_grammar_tool(self):
        """Lazily start the LanguageTool checker (spelling + sentence/grammar).
        Returns None if the package or a Java runtime isn't available."""
        if not _GRAMMAR_AVAILABLE:
            return None
        if self._grammar_tool is False:
            return None
        if self._grammar_tool is None:
            try:
                self._grammar_tool = language_tool_python.LanguageTool("en-US")
            except Exception:
                self._grammar_tool = False
                return None
        return self._grammar_tool

    def _check_and_fix_text(self, text_widget, label):
        """Check the given Issue/Note box for spelling and sentence errors
        and, if the user agrees, replace its text with the corrected version."""
        text = text_widget.get("1.0", "end").strip()
        if not text:
            return

        tool = self._get_grammar_tool()
        if tool is None:
            messagebox.showinfo(
                "Spelling & Grammar",
                "Spelling/grammar checking isn't available on this machine.\n\n"
                "Install it with:\n    pip install language-tool-python\n\n"
                "It also needs a Java runtime installed.",
            )
            return

        try:
            matches = tool.check(text)
            corrected = language_tool_python.utils.correct(text, matches)
        except Exception as exc:
            messagebox.showwarning("Spelling & Grammar", f"Couldn't check the {label} text right now:\n{exc}")
            return

        if not matches or corrected.strip() == text:
            messagebox.showinfo("Spelling & Grammar", f"No issues found in {label}.")
            return

        if messagebox.askyesno(
            "Spelling & Grammar",
            f"Found {len(matches)} possible issue(s) in {label}.\n\n"
            f"Suggested:\n\"{corrected}\"\n\nApply this correction?",
        ):
            text_widget.delete("1.0", "end")
            text_widget.insert("1.0", corrected)

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
        self.last_issue_bookmark_name = None
        self.last_issue_paragraph_p = None
        self.last_step_elements = []
        self.last_step_had_issue = False
        self.last_step_summary_element = None
        self.last_step_prev_issue_bookmark = None
        self.last_step_prev_issue_paragraph_p = None
        self.last_step_arrow_element = None

        self.doc = Document()
        set_update_fields_on_open(self.doc)
        set_page_border(self.doc)

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
        meta.add_run(f"Started: {self.start_time.strftime('%m/%d/%Y %I:%M:%S %p')}").italic = True

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
            step_prev_issue_bookmark = None
            step_prev_issue_paragraph_p = None
            step_arrow_element = None

            if issue:
                self.bookmark_counter += 1
                bookmark_name = f"issue_{self.bookmark_counter}"
                self.bookmark_names.append(bookmark_name)

                # remember the chain state as it was BEFORE this issue, so
                # undo can put it back exactly
                step_prev_issue_bookmark = self.last_issue_bookmark_name
                step_prev_issue_paragraph_p = self.last_issue_paragraph_p

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

                # chain the PREVIOUS issue forward to this one with a
                # clickable "Next Issue" arrow at the end of its line
                if step_prev_issue_paragraph_p is not None:
                    step_arrow_element = add_next_issue_arrow(
                        step_prev_issue_paragraph_p, bookmark_name, color=ISSUE_COLOR_HEX
                    )

                # mirrored into the Issues Summary block at the top - the
                # whole entry is a clickable link straight to the issue
                summary_para = self.doc.add_paragraph()
                self.last_summary_element.addnext(summary_para._p)
                self.last_summary_element = summary_para._p

                bullet = summary_para.add_run("\u2022 ")
                bullet.font.color.rgb = ISSUE_COLOR
                add_issue_summary_hyperlink(summary_para, bookmark_name, issue, color=ISSUE_COLOR_HEX)

                step_had_issue = True
                step_summary_element = summary_para._p

                # this issue is now the head of the "Next Issue" chain
                self.last_issue_bookmark_name = bookmark_name
                self.last_issue_paragraph_p = issue_para._p

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
            self.last_step_prev_issue_bookmark = step_prev_issue_bookmark
            self.last_step_prev_issue_paragraph_p = step_prev_issue_paragraph_p
            self.last_step_arrow_element = step_arrow_element

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

            # undo the "Next Issue" arrow this issue chained onto the
            # previous issue, and roll the chain pointer back to it
            if self.last_step_arrow_element is not None:
                arrow_parent = self.last_step_arrow_element.getparent()
                if arrow_parent is not None:
                    arrow_parent.remove(self.last_step_arrow_element)
            self.last_issue_bookmark_name = self.last_step_prev_issue_bookmark
            self.last_issue_paragraph_p = self.last_step_prev_issue_paragraph_p

        self.step_count -= 1
        self.count_var.set(f"{self.step_count} step(s) captured")

        # this step is gone - there is nothing further back tracked to undo
        self.last_step_elements = []
        self.last_step_had_issue = False
        self.last_step_summary_element = None
        self.last_step_prev_issue_bookmark = None
        self.last_step_prev_issue_paragraph_p = None
        self.last_step_arrow_element = None
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
        summary.add_run(f"Ended: {end_time.strftime('%m/%d/%Y %I:%M:%S %p')}\n").italic = True
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
        self.last_issue_bookmark_name = None
        self.last_issue_paragraph_p = None
        self.last_step_elements = []
        self.last_step_had_issue = False
        self.last_step_summary_element = None
        self.last_step_prev_issue_bookmark = None
        self.last_step_prev_issue_paragraph_p = None
        self.last_step_arrow_element = None
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