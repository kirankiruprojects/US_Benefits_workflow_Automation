"""Automated build script for creating the standalone, zero-prerequisite
Workforce Junction Windows desktop application.

When built, the application can be copied to ANY Windows 10/11 laptop and
run with a single double-click — NO Node.js or Python installation needed!
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import zipfile


def print_step(title: str) -> None:
    print("\n" + "=" * 65)
    print(f"  {title}")
    print("=" * 65 + "\n")


def run_cmd(cmd: list[str] | str, cwd: str, env: dict | None = None, shell: bool = False) -> None:
    merged_env = os.environ.copy()
    if env:
        merged_env.update(env)
    
    print(f">> Running: {cmd if isinstance(cmd, str) else ' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=cwd, env=merged_env, shell=shell)
    if result.returncode != 0:
        raise RuntimeError(f"Command failed with exit code {result.returncode}")


def find_system_node() -> str:
    """Find system Node.js executable."""
    candidates = [
        r"C:\Program Files\nodejs\node.exe",
        r"C:\Program Files (x86)\nodejs\node.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\nodejs\node.exe"),
        os.path.expandvars(r"%APPDATA%\npm\node.exe"),
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    which_node = shutil.which("node")
    if which_node:
        return which_node
    raise FileNotFoundError("Could not locate node.exe on this system.")


def kill_running_app() -> None:
    """Kill Workforce Junction and any node.exe it spawned so file locks are released."""
    import time

    killed_any = False
    for name in ["Workforce Junction.exe", "node.exe", "msedgedriver.exe", "chromedriver.exe"]:
        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", name],
                capture_output=True,
                text=True,
            )
            if "SUCCESS" in (result.stdout or ""):
                print(f"  Terminated {name}")
                killed_any = True
        except Exception:
            pass

    if killed_any:
        print("Waiting for processes to fully release file handles...")
        time.sleep(4)  # Give Windows time to release all file locks
    else:
        print("  No running instances found.")


def rmtree_force(path: str, retries: int = 5, delay: float = 1.5) -> None:
    """Remove a directory tree, retrying on Windows file-lock errors."""
    import time
    import stat

    def handle_error(func, fpath, excinfo):
        # Try to make read-only files writable then delete
        try:
            os.chmod(fpath, stat.S_IWRITE)
            func(fpath)
        except Exception:
            pass

    for attempt in range(retries):
        try:
            shutil.rmtree(path, onerror=handle_error)
            return
        except Exception as exc:
            if attempt < retries - 1:
                print(f"  Retry {attempt + 1}/{retries}: {exc}")
                time.sleep(delay)
            else:
                print(f"  Warning: Could not fully remove {path}: {exc}")


def build_application() -> None:
    root_dir = os.path.dirname(os.path.abspath(__file__))
    desktop_dir = os.path.join(root_dir, "desktop")
    release_dir = os.path.join(root_dir, "release")
    staging_dir = os.path.join(root_dir, "release_staging")  # Build here, then swap
    app_output_dir = os.path.join(release_dir, "Workforce Junction")

    print_step("Step 0: Stopping Any Running Application Instances")
    kill_running_app()

    # Clean up any leftover staging dir from a previous failed build
    if os.path.exists(staging_dir):
        print("Cleaning leftover staging directory...")
        rmtree_force(staging_dir)

    print_step("Step 1: Building Production Web UI Bundle (.output)")
    build_env = {"NITRO_PRESET": "node-server", "NODE_ENV": "production"}
    run_cmd("npm run build", cwd=root_dir, env=build_env, shell=True)

    dot_output_dir = os.path.join(root_dir, ".output")
    if not os.path.exists(os.path.join(dot_output_dir, "server", "index.mjs")):
        raise FileNotFoundError(f"Expected server build at {dot_output_dir}/server/index.mjs but not found.")

    print_step("Step 2: Locating Embedded Node.js Runtime")
    node_exe = find_system_node()
    print(f"Found Node.js at: {node_exe}")

    print_step("Step 3: Compiling Python Backend Executable with PyInstaller")
    # Build into staging_dir so we never touch release/ during compilation.
    # This avoids WinError 32 from node.exe locking files inside release/.
    os.makedirs(staging_dir, exist_ok=True)

    pyinstaller_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconsole",
        "--onedir",
        "--name=Workforce Junction",
        f"--distpath={os.path.join(staging_dir, 'dist')}",
        f"--workpath={os.path.join(staging_dir, 'build')}",
        f"--specpath={staging_dir}",
        f"--paths={desktop_dir}",
        "--hidden-import=api",
        "--hidden-import=excel_automations",
        "--hidden-import=selenium_automations",
        "--hidden-import=Screenshot",
        "--hidden-import=docx",
        "--hidden-import=screeninfo",
        "--hidden-import=keyboard",
        "--hidden-import=PIL",
        "--hidden-import=PIL.ImageGrab",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=win32gui",
        "--hidden-import=openpyxl",
        "--hidden-import=pandas",
        "--hidden-import=selenium",
        "--hidden-import=selenium.webdriver",
        "--hidden-import=selenium.webdriver.edge.service",
        "--hidden-import=selenium.webdriver.chrome.service",
        "--hidden-import=clr_loader",
        "--hidden-import=pythonnet",
        "--hidden-import=webview",
        "--hidden-import=webview.platforms.winforms",
        "--hidden-import=webview.platforms.edgechromium",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.filedialog",
        "--collect-all=webview",
        "--collect-all=selenium",
        "--collect-all=clr_loader",
        "--collect-all=pythonnet",
        "--collect-all=docx",
        "--collect-all=screeninfo",
        "--collect-all=keyboard",
        "--collect-all=PIL",
        "--collect-all=language_tool_python",
        "--hidden-import=language_tool_python",
        "--noconfirm",
        os.path.join(desktop_dir, "main.py"),
    ]

    run_cmd(pyinstaller_cmd, cwd=root_dir)

    print_step("Step 4: Assembling Standalone Distribution Package")
    built_bundle = os.path.join(staging_dir, "dist", "Workforce Junction")
    if not os.path.exists(built_bundle):
        raise FileNotFoundError(f"PyInstaller build directory not found at: {built_bundle}")

    # Assemble the complete app inside staging_dir/app/
    staging_app = os.path.join(staging_dir, "Workforce Junction")
    if os.path.exists(staging_app):
        rmtree_force(staging_app)
    shutil.move(built_bundle, staging_app)

    # Copy embedded node.exe into staging app folder
    shutil.copy2(node_exe, os.path.join(staging_app, "node.exe"))
    print(f"Copied node.exe -> {os.path.join(staging_app, 'node.exe')}")

    # Copy .output folder into staging app folder
    shutil.copytree(dot_output_dir, os.path.join(staging_app, ".output"))
    print(f"Copied .output -> {os.path.join(staging_app, '.output')}")

    # Create HOW_TO_RUN.txt
    readme_path = os.path.join(staging_app, "HOW_TO_RUN.txt")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(
            "===========================================================\n"
            "         Workforce Junction - Standalone Desktop App       \n"
            "===========================================================\n\n"
            "HOW TO RUN:\n"
            "1. Double-click 'Workforce Junction.exe' to launch the application.\n"
            "2. If Microsoft Defender SmartScreen appears:\n"
            "   Click 'More info' and then click 'Run anyway'.\n\n"
            "CRITICAL - IF SHARING AS A ZIP FILE:\n"
            "- Do NOT run 'Workforce Junction.exe' directly from inside the ZIP file.\n"
            "- Right-click 'Workforce-Junction-Portable.zip' and select 'Extract All...'.\n"
            "- Open the extracted folder and then run 'Workforce Junction.exe'.\n\n"
            "--- TROUBLESHOOTING ---\n"
            "ERROR: 'Failed to resolve Python.Runtime.Loader.Initialize'\n"
            "This error happens because Windows BLOCKS downloaded files for security.\n\n"
            "QUICK FIX (choose one):\n"
            "Option A - One-Click Fix (Recommended):\n"
            "  Double-click 'Fix-Blocked-Files.bat' in this folder.\n"
            "  It will unblock all files automatically, then re-run the app.\n\n"
            "Option B - Manual Fix via File Properties:\n"
            "  1. Delete the extracted folder.\n"
            "  2. Right-click the ZIP file -> Properties.\n"
            "  3. Tick the 'Unblock' checkbox at the bottom -> OK.\n"
            "  4. Re-extract the ZIP and run the app.\n\n"
            "Option C - PowerShell (Run as Administrator):\n"
            "  Get-ChildItem -Path \"<folder path>\" -Recurse | Unblock-File\n\n"
            "REQUIREMENTS ON TARGET LAPTOP:\n"
            "- None! No Node.js or Python installation is needed.\n"
            "- Microsoft Edge or Google Chrome (pre-installed on Windows 10/11).\n\n"
            "FEATURES INCLUDED:\n"
            "1. Automation Dashboard\n"
            "2. Rates Calculator\n"
            "3. Salary File Automation\n"
            "4. Enrollment Report Audit\n"
            "5. Pre vs Post Comparison Tool\n"
            "6. New Client Password Setup\n"
            "7. New Client Test Records\n"
            "8. Auto-Login Tester\n"
            "9. Custom Report Generator\n"
            "10. Clarification Form\n"
            "11. Workflow Capture / Screenshot Tool\n"
            "===========================================================\n"
        )

    # Create Fix-Blocked-Files.bat — colleagues double-click this to unblock all files
    bat_path = os.path.join(staging_app, "Fix-Blocked-Files.bat")
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(
            "@echo off\n"
            "echo ============================================================\n"
            "echo  Workforce Junction - Unblocking Files (Security Fix)\n"
            "echo ============================================================\n"
            "echo.\n"
            "echo Windows sometimes blocks downloaded files for security.\n"
            "echo This script removes those blocks so the app runs correctly.\n"
            "echo.\n"
            "powershell -NoProfile -ExecutionPolicy Bypass -Command "
            "\"Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File; "
            "Write-Host 'All files unblocked successfully!' -ForegroundColor Green\"\n"
            "echo.\n"
            "echo Done! You can now double-click 'Workforce Junction.exe' to run the app.\n"
            "pause\n"
        )
    print(f"Created Fix-Blocked-Files.bat -> {bat_path}")

    # === Move assembled app into final release dir ===
    print("Moving assembled app to release directory...")
    if os.path.exists(release_dir):
        rmtree_force(release_dir)
    os.makedirs(release_dir, exist_ok=True)
    # Move just the app folder — avoids os.rename cross-dir issues on Windows
    app_output_dir = os.path.join(release_dir, "Workforce Junction")
    shutil.move(staging_app, app_output_dir)
    # Clean up the rest of staging (build/dist intermediate files)
    rmtree_force(staging_dir)

    print_step("Step 5: Creating Compressed ZIP Archive for Easy Sharing")
    zip_path = os.path.join(release_dir, "Workforce-Junction-Portable.zip")
    if os.path.exists(zip_path):
        os.remove(zip_path)

    print(f"Zipping {app_output_dir} -> {zip_path}...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(app_output_dir):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, release_dir)
                zipf.write(full_path, rel_path)

    print_step("BUILD COMPLETE SUCCESSFUL!")
    print(f"1. Standalone Application Folder: {app_output_dir}")
    print(f"2. Executable to Run:             {os.path.join(app_output_dir, 'Workforce Junction.exe')}")
    print(f"3. Portable ZIP for Distribution: {zip_path}")
    print("\nYou can now copy this folder or ZIP to any other laptop and double-click to run!")


if __name__ == "__main__":
    try:
        build_application()
    except Exception as e:
        print(f"\n[BUILD ERROR] {e}")
        sys.exit(1)
