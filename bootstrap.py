from __future__ import annotations

import json
import os
import shutil
import socket
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent
LOGS = ROOT / "logs"
RUNTIME = ROOT / "runtime"
LOGS.mkdir(exist_ok=True)
RUNTIME.mkdir(exist_ok=True)
STARTUP_LOG = LOGS / "startup.log"
SERVER_OUT = LOGS / "server.out"
SERVER_ERR = LOGS / "server.err"
PID_FILE = RUNTIME / "server.pid"
SERVER_INFO = RUNTIME / "server.json"
APP_ID = "journal-figure-extractor"
APP_VERSION = "1.2"
PORT_MIN = 8765
PORT_MAX = 8785


def log(msg: str = "") -> None:
    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line, flush=True)
    with STARTUP_LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(args, *, timeout=None, check=False):
    p = subprocess.run(
        [str(x) for x in args], cwd=ROOT, text=True, encoding="utf-8",
        errors="replace", stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        timeout=timeout,
    )
    if check and p.returncode != 0:
        raise RuntimeError(
            f"Command failed ({p.returncode}): {' '.join(map(str, args))}\n"
            f"STDOUT:\n{p.stdout}\nSTDERR:\n{p.stderr}"
        )
    return p


def health_info(port: int, timeout=0.8):
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/health", timeout=timeout) as r:
            if r.status != 200:
                return None
            data = json.loads(r.read().decode("utf-8", errors="replace"))
            return data if isinstance(data, dict) else None
    except Exception:
        return None


def port_is_free(port: int) -> bool:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(("127.0.0.1", port))
        return True
    except OSError:
        return False
    finally:
        s.close()


def choose_port() -> tuple[int, bool]:
    # Reuse only THIS exact app version. An old v0.4/v0.6 health endpoint must not
    # trick the launcher into opening the wrong backend.
    for port in range(PORT_MIN, PORT_MAX + 1):
        info = health_info(port)
        if info and info.get("app_id") == APP_ID and info.get("version") == APP_VERSION:
            log(f"Current v{APP_VERSION} server already running on port {port}.")
            return port, True
        if info:
            log(
                f"Port {port} is occupied by another/older HTTP service "
                f"(app_id={info.get('app_id')!r}, version={info.get('version')!r}); skipping it."
            )
            continue
        if port_is_free(port):
            if port != PORT_MIN:
                log(f"Port {PORT_MIN} is occupied; v{APP_VERSION} will use port {port} instead.")
            return port, False
        log(f"Port {port} is occupied by a non-HTTP or unknown service; skipping it.")
    raise RuntimeError(f"No free local port found in {PORT_MIN}-{PORT_MAX}.")


def ensure_venv() -> Path:
    if sys.version_info < (3, 10):
        raise RuntimeError(f"Need Python 3.10+, found {sys.version.split()[0]}")
    venv = ROOT / ".venv"
    vpy = venv / "Scripts" / "python.exe"
    if vpy.exists():
        probe = run([vpy, "-c", "import sys; print(sys.version)"])
        if probe.returncode == 0:
            log(f"Virtual environment OK: {vpy}")
            return vpy
        log("Existing .venv is broken; recreating it.")
        shutil.rmtree(venv, ignore_errors=True)
    log("Creating .venv ...")
    p = run([sys.executable, "-m", "venv", str(venv)])
    if p.returncode != 0 or not vpy.exists():
        raise RuntimeError(f"Could not create .venv.\n{p.stdout}\n{p.stderr}")
    return vpy


def ensure_packages(vpy: Path) -> None:
    probe = run([vpy, "-c", "import flask, pymupdf, win32com.client, PIL; print('PACKAGES_OK')"])
    if probe.returncode == 0:
        log("Packages OK.")
        return
    log("Installing required packages (first run may take a few minutes) ...")
    p = run([vpy, "-m", "pip", "install", "--disable-pip-version-check", "-r", ROOT / "requirements.txt"])
    if p.stdout.strip():
        for line in p.stdout.splitlines()[-40:]:
            log("pip: " + line)
    if p.stderr.strip():
        for line in p.stderr.splitlines()[-40:]:
            log("pip: " + line)
    if p.returncode != 0:
        raise RuntimeError("pip install failed. See logs/startup.log")
    probe = run([vpy, "-c", "import flask, pymupdf, win32com.client, PIL; print('PACKAGES_OK')"])
    if probe.returncode != 0:
        raise RuntimeError(f"Packages installed but imports still fail:\n{probe.stderr}")
    log("Packages installed successfully.")


def check_word(vpy: Path) -> None:
    log("Checking Microsoft Word COM ...")
    p = run([vpy, ROOT / "check_word.py"], timeout=30)
    if p.stdout.strip():
        for line in p.stdout.splitlines():
            log(line)
    if p.returncode != 0:
        if p.stderr.strip():
            for line in p.stderr.splitlines()[-30:]:
                log("Word: " + line)
        raise RuntimeError("Microsoft Word COM check failed. See logs/startup.log")


def start_server(vpy: Path) -> int:
    port, reuse = choose_port()
    if reuse:
        return port
    for f in (SERVER_OUT, SERVER_ERR):
        try:
            f.unlink()
        except FileNotFoundError:
            pass
    out = SERVER_OUT.open("w", encoding="utf-8")
    err = SERVER_ERR.open("w", encoding="utf-8")
    flags = 0
    if os.name == "nt":
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["JFE_PORT"] = str(port)
    proc = subprocess.Popen(
        [str(vpy), str(ROOT / "app.py")], cwd=str(ROOT), stdout=out,
        stderr=err, env=env, creationflags=flags,
    )
    PID_FILE.write_text(str(proc.pid), encoding="ascii")
    SERVER_INFO.write_text(json.dumps({"pid": proc.pid, "port": port, "version": APP_VERSION}, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"Server PID: {proc.pid}; port: {port}; version: {APP_VERSION}")
    for _ in range(60):
        if proc.poll() is not None:
            break
        info = health_info(port, timeout=0.8)
        if info and info.get("app_id") == APP_ID and info.get("version") == APP_VERSION:
            log(f"Server health check passed for v{APP_VERSION} on port {port}.")
            return port
        time.sleep(0.5)
    try:
        out.close(); err.close()
    except Exception:
        pass
    tail = ""
    if SERVER_ERR.exists():
        tail = "\n".join(SERVER_ERR.read_text(encoding="utf-8", errors="replace").splitlines()[-80:])
    raise RuntimeError("Backend failed to start.\n" + tail)


def main() -> int:
    STARTUP_LOG.write_text("", encoding="utf-8")
    log(f"Journal Figure Extractor v{APP_VERSION} startup")
    log(f"Root: {ROOT}")
    log(f"Bootstrap Python: {sys.executable}")
    log(f"Python version: {sys.version.split()[0]}")
    vpy = ensure_venv()
    ensure_packages(vpy)
    check_word(vpy)
    port = start_server(vpy)
    url = f"http://127.0.0.1:{port}/?v={APP_VERSION}"
    log("SUCCESS. Opening browser: " + url)
    webbrowser.open(url)
    return 0


if __name__ == "__main__":
    try:
        code = main()
    except Exception as exc:
        log("")
        log("STARTUP FAILED")
        log(f"{type(exc).__name__}: {exc}")
        print("\n启动失败。请把 logs\\startup.log 发给我。", flush=True)
        try:
            input("按 Enter 键退出...")
        except Exception:
            pass
        code = 1
    raise SystemExit(code)
