from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
APP_PATH = BASE_DIR / "app.py"
OUTPUT_DIR = BASE_DIR / "outputs"
LOG_PATH = OUTPUT_DIR / "launcher.log"
SERVER_LOG_PATH = OUTPUT_DIR / "server-start.log"
PORT = 8000
PUBLIC_HOSTNAME = "oratoriocarloacutis.don"
HEALTHCHECK_URL = f"http://127.0.0.1:{PORT}/login"
URL = f"http://{PUBLIC_HOSTNAME}:{PORT}/login"
HOST = "127.0.0.1"
LOCAL_PYTHON = BASE_DIR / "runtime" / "python" / "python.exe"
BUNDLED_PYTHON = (
    Path.home()
    / ".cache"
    / "codex-runtimes"
    / "codex-primary-runtime"
    / "dependencies"
    / "python"
    / "python.exe"
)


def log(message: str) -> None:
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        with LOG_PATH.open("a", encoding="utf-8") as handle:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            handle.write(f"[{timestamp}] {message}\n")
    except OSError:
        pass


def http_ready(timeout: float = 12.0) -> bool:
    try:
        with urllib.request.urlopen(HEALTHCHECK_URL, timeout=timeout) as response:
            return 200 <= int(response.status) < 500
    except Exception:
        return False


def port_open(timeout: float = 0.4) -> bool:
    probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    probe.settimeout(timeout)
    try:
        return probe.connect_ex((HOST, PORT)) == 0
    finally:
        probe.close()


def stop_stale_listener() -> None:
    command = (
        'for /f "tokens=5" %p in (\'netstat -ano ^| findstr ":8000" ^| find "LISTENING"\') '
        "do taskkill /PID %p /F >nul 2>&1"
    )
    subprocess.run(
        ["cmd", "/c", command],
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    deadline = time.time() + 5.0
    while time.time() < deadline:
        if not port_open():
            return
        time.sleep(0.25)


def resolve_python() -> str:
    if LOCAL_PYTHON.exists():
        return str(LOCAL_PYTHON)
    if BUNDLED_PYTHON.exists():
        return str(BUNDLED_PYTHON)
    current = Path(sys.executable)
    sibling = current.with_name("python.exe")
    if sibling.exists():
        return str(sibling)
    return str(current)


def start_server() -> None:
    python_exe = resolve_python()
    creationflags = (
        getattr(subprocess, "DETACHED_PROCESS", 0)
        | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with SERVER_LOG_PATH.open("ab") as server_log:
        server_log.write(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] Avvio server\n".encode("utf-8"))
        subprocess.Popen(
            [python_exe, str(APP_PATH)],
            cwd=str(BASE_DIR),
            stdout=server_log,
            stderr=subprocess.STDOUT,
            creationflags=creationflags,
        )


def wait_until_ready(seconds: float = 35.0) -> bool:
    deadline = time.time() + seconds
    while time.time() < deadline:
        remaining = max(1.0, deadline - time.time())
        if http_ready(timeout=min(15.0, remaining)):
            return True
        time.sleep(0.5)
    return False


def open_browser() -> None:
    os.startfile(URL)


def main() -> None:
    try:
        if http_ready():
            log("Server gia attivo. Apro il browser.")
            open_browser()
            return

        if port_open():
            log("Porta 8000 occupata senza risposta HTTP. Riavvio del server.")
            stop_stale_listener()
            time.sleep(1.0)

        start_server()
        ready = wait_until_ready()
        if not ready:
            log("Primo avvio non riuscito. Nuovo tentativo.")
            stop_stale_listener()
            time.sleep(1.0)
            start_server()
            ready = wait_until_ready()
        if not ready:
            log("Il server non ha risposto entro il tempo previsto.")
        else:
            log("Server pronto. Apro il browser.")
        open_browser()
    except Exception as exc:
        log(f"Errore launcher: {exc!r}")
        raise


if __name__ == "__main__":
    main()
