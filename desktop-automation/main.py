"""Desktop launcher for the unified Workforce Junction application.

This launcher opens the web UI inside a native pywebview window so that the
Python automation bridge (`window.pywebview.api`) is connected.

Works in two modes:
1. Standalone / Production mode: Automatically starts the embedded local web server
   (Node runtime + pre-built .output SSR server) on a local port, opens the window,
   and cleanly terminates the server upon window exit.
2. Development mode: Connects to a running Vite/Lovable dev server at http://localhost:8080.
"""

from __future__ import annotations

import argparse
import atexit
import os
import shutil
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import webview

from api import Api


def get_base_dirs() -> list[str]:
    """Return search directories for application assets and executables."""
    dirs = []
    if getattr(sys, "frozen", False):
        # PyInstaller bundled executable directory
        dirs.append(os.path.dirname(sys.executable))
        if hasattr(sys, "_MEIPASS"):
            dirs.append(sys._MEIPASS)
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dirs.append(script_dir)
    dirs.append(os.path.abspath(os.path.join(script_dir, "..")))
    return list(dict.fromkeys(dirs))


def find_file(relative_path: str) -> str | None:
    """Find a relative file across base search directories."""
    for base in get_base_dirs():
        candidate = os.path.join(base, relative_path)
        if os.path.exists(candidate):
            return os.path.abspath(candidate)
    return None


def find_node_executable() -> str | None:
    """Locate the Node.js executable (bundled or system)."""
    # 1. Check embedded node.exe next to executable or script
    for base in get_base_dirs():
        for candidate_name in ["node.exe", "node", os.path.join("bin", "node.exe"), os.path.join("runtime", "node.exe")]:
            candidate = os.path.join(base, candidate_name)
            if os.path.exists(candidate) and os.path.isfile(candidate):
                return os.path.abspath(candidate)

    # 2. Check standard installation paths on Windows
    win_paths = [
        r"C:\Program Files\nodejs\node.exe",
        r"C:\Program Files (x86)\nodejs\node.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\nodejs\node.exe"),
        os.path.expandvars(r"%APPDATA%\npm\node.exe"),
    ]
    for path in win_paths:
        if os.path.exists(path):
            return os.path.abspath(path)

    # 3. Check PATH
    which_node = shutil.which("node")
    if which_node:
        return os.path.abspath(which_node)

    return None


def is_port_in_use(port: int) -> bool:
    """Check if a local TCP port is already in use."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) == 0


def find_free_port(preferred: int = 3000) -> int:
    """Find an open TCP port starting from preferred."""
    if not is_port_in_use(preferred):
        return preferred
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def wait_for_server(url: str, timeout_seconds: float = 12.0) -> bool:
    """Poll URL until it responds or timeout occurs."""
    start_time = time.time()
    while time.time() - start_time < timeout_seconds:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "WorkforceJunctionLauncher"})
            with urllib.request.urlopen(req, timeout=1.0) as response:
                if response.status in (200, 301, 302, 304, 404):
                    return True
        except (urllib.error.HTTPError, urllib.error.URLError, ConnectionRefusedError, TimeoutError):
            pass
        except Exception:
            pass
        time.sleep(0.2)
    return False


# Global handle for background server process cleanup
_SERVER_PROCESS: subprocess.Popen | None = None


def cleanup_server_process() -> None:
    """Terminate any spawned background server process cleanly."""
    global _SERVER_PROCESS
    if _SERVER_PROCESS is not None:
        try:
            if _SERVER_PROCESS.poll() is None:
                _SERVER_PROCESS.terminate()
                try:
                    _SERVER_PROCESS.wait(timeout=2.0)
                except subprocess.TimeoutExpired:
                    _SERVER_PROCESS.kill()
        except Exception:
            pass
        finally:
            _SERVER_PROCESS = None


atexit.register(cleanup_server_process)


def start_embedded_server(port: int) -> str | None:
    """Start the embedded Nitro SSR server using the bundled/available Node runtime."""
    global _SERVER_PROCESS

    server_entry = find_file(os.path.join(".output", "server", "index.mjs"))
    if not server_entry:
        return None

    node_bin = find_node_executable()
    if not node_bin:
        print("[Launcher] Warning: Embedded server found but node executable is missing.")
        return None

    output_dir = os.path.dirname(os.path.dirname(server_entry))

    env = os.environ.copy()
    env["PORT"] = str(port)
    env["HOST"] = "127.0.0.1"
    env["NITRO_PORT"] = str(port)
    env["NITRO_HOST"] = "127.0.0.1"
    env["NODE_ENV"] = "production"

    creation_flags = 0
    if sys.platform == "win32":
        # CREATE_NO_WINDOW = 0x08000000
        creation_flags = 0x08000000

    try:
        _SERVER_PROCESS = subprocess.Popen(
            [node_bin, server_entry],
            cwd=output_dir,
            env=env,
            creationflags=creation_flags,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return f"http://127.0.0.1:{port}"
    except Exception as exc:
        print(f"[Launcher] Error starting server process: {exc}")
        return None


def resolve_start_url(preferred_url: str | None = None, prefer_dev: bool = False) -> str:
    """Determine the optimal URL to load inside the pywebview window."""
    if preferred_url:
        return preferred_url

    # Check if a dev server is active on port 8080 (e.g. `npm run dev`)
    if prefer_dev or is_port_in_use(8080):
        if wait_for_server("http://127.0.0.1:8080", timeout_seconds=1.0):
            print("[Launcher] Connected to active dev server at http://127.0.0.1:8080")
            return "http://127.0.0.1:8080"

    # Otherwise, start or connect to the standalone server
    target_port = 3000
    if is_port_in_use(target_port):
        if wait_for_server(f"http://127.0.0.1:{target_port}", timeout_seconds=1.0):
            print(f"[Launcher] Connected to existing server at http://127.0.0.1:{target_port}")
            return f"http://127.0.0.1:{target_port}"
        target_port = find_free_port(3000)

    server_url = start_embedded_server(target_port)
    if server_url:
        print(f"[Launcher] Starting embedded server on port {target_port}...")
        if wait_for_server(server_url, timeout_seconds=8.0):
            print(f"[Launcher] Embedded server ready at {server_url}")
            return server_url
        print("[Launcher] Server did not respond within timeout, attempting connection anyway...")
        return server_url

    # Fallback to dev server default
    return "http://127.0.0.1:8080"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Workforce Junction Unified Desktop Application"
    )
    parser.add_argument(
        "--url",
        type=str,
        default=None,
        help="Custom URL to open inside the application window.",
    )
    parser.add_argument(
        "--dev",
        action="store_true",
        help="Prioritize connecting to the development server (localhost:8080).",
    )
    args = parser.parse_args()

    start_url = resolve_start_url(preferred_url=args.url, prefer_dev=args.dev)
    print(f"[Launcher] Opening Workforce Junction UI at: {start_url}")

    api = Api()
    window = webview.create_window(
        "Workforce Junction",
        start_url,
        js_api=api,
        width=1450,
        height=950,
        min_size=(1100, 700),
    )

    try:
        webview.start()
    finally:
        cleanup_server_process()
        print("[Launcher] Workforce Junction closed.")


if __name__ == "__main__":
    main()
