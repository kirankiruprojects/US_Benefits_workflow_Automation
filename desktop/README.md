# Workforce Junction — Desktop launcher

This folder contains the desktop wrapper that turns the web UI into a local Windows application with a Python backend.

## Files

| File | Purpose |
|------|---------|
| `main.py` | Opens the web UI in a `pywebview` window. This is what creates the `window.pywebview.api` bridge. |
| `api.py` | Python methods exposed to the web UI (file pickers, Excel automations, portal automations). |
| `selenium_automations.py` | Place to put your Selenium portal automation code (password setup, auto-login, custom report, etc.). |

## Where to put your existing scripts

Keep your original Python files in the **same folder** as `main.py`:

```
workforce-junction-desktop/
├── main.py
├── api.py
├── selenium_automations.py
├── Enrollment_Report_Audit.py
├── Salary_File_Automation_Final.py
├── comparison_tool.py
└── (your other .py Selenium scripts)
```

Do **not** rename the Excel automation scripts above — `api.py` imports them by those exact names.

## Install

```sh
pip install pywebview pandas openpyxl selenium
```

## Run in development mode

Start the Lovable dev server first (`npm run dev` or `bun dev`), then:

```sh
python main.py
```

The window loads `http://localhost:8080` and the Browse / Run buttons become active.

## Run a packaged build

Build the web app first so a `.output/` or `dist/` folder exists, then:

```sh
python main.py
```

## Build Standalone Zero-Install Executable (.exe)

To build a standalone portable application that runs on **any laptop without Node.js or Python**:

Double-click `build-app.bat` or run:

```sh
python build-standalone.py
```

This creates:
- `release/Workforce Junction/` (Standalone folder with `Workforce Junction.exe`)
- `release/Workforce-Junction-Portable.zip` (Portable ZIP file ready to share)

## Adding live logs to your existing scripts

`api.py` captures anything your script prints with `print()` and forwards it to the React log console. If you want richer status updates, pass the `ui` object from `selenium_automations.py` into your code and call:

```python
ui.log("Step completed", "ok")
ui.status("Working...", "idle")
ui.progress(50)
```

