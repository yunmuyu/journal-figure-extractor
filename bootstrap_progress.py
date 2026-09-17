from __future__ import annotations

import json
import os
import subprocess
import time
from pathlib import Path

import bootstrap as b

b.APP_VERSION = "1.4"


def start_progress_server(vpy: Path) -> int:
    port, reuse = b.choose_port()
    if reuse:
        return port

    for f in (b.SERVER_OUT, b.SERVER_ERR):
        try:
            f.unlink()
        except FileNotFoundError:
            pass

    out = b.SERVER_OUT.open("w", encoding="utf-8")
    err = b.SERVER_ERR.open("w", encoding="utf-8")
    flags = 0
    if os.name == "nt":
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["JFE_PORT"] = str(port)

    proc = subprocess.Popen(
        [str(vpy), str(b.ROOT / "progress_server.py")],
        cwd=str(b.ROOT), stdout=out, stderr=err, env=env, creationflags=flags,
    )
    b.PID_FILE.write_text(str(proc.pid), encoding="ascii")
    b.SERVER_INFO.write_text(
        json.dumps({"pid": proc.pid, "port": port, "version": b.APP_VERSION}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    b.log(f"Server PID: {proc.pid}; port: {port}; version: {b.APP_VERSION}")

    for _ in range(60):
        if proc.poll() is not None:
            break
        info = b.health_info(port, timeout=0.8)
        if info and info.get("app_id") == b.APP_ID and info.get("version") == b.APP_VERSION:
            b.log(f"Server health check passed for v{b.APP_VERSION} on port {port}.")
            return port
        time.sleep(0.5)

    try:
        out.close(); err.close()
    except Exception:
        pass

    tail = ""
    if b.SERVER_ERR.exists():
        tail = "\n".join(b.SERVER_ERR.read_text(encoding="utf-8", errors="replace").splitlines()[-80:])
    raise RuntimeError("Backend failed to start.\n" + tail)


b.start_server = start_progress_server


if __name__ == "__main__":
    try:
        code = b.main()
    except Exception as exc:
        b.log("")
        b.log("STARTUP FAILED")
        b.log(f"{type(exc).__name__}: {exc}")
        print("\n启动失败。请把 logs\\startup.log 发给我。", flush=True)
        try:
            input("按 Enter 键退出...")
        except Exception:
            pass
        code = 1
    raise SystemExit(code)
