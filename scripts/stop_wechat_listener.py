#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Stop the background WeChat listener."""

from __future__ import annotations

import signal
import subprocess
from pathlib import Path

from wechat_listener_request import task_name_for_target

ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = ROOT / "runtime"
STOP_FILE = RUNTIME_DIR / "wechat_listener.stop"
PID_FILE = RUNTIME_DIR / "wechat_listener.pid"
REQUESTS_DIR = RUNTIME_DIR / "requests"


def main() -> int:
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    STOP_FILE.write_text("stop\n", encoding="utf-8")
    stopped_pid = None

    pid = None
    if PID_FILE.exists():
        raw = PID_FILE.read_text(encoding="utf-8").strip()
        if raw.isdigit():
            pid = int(raw)

    if pid:
        try:
            if hasattr(signal, "CTRL_BREAK_EVENT"):
                subprocess.run(
                    ["taskkill", "/PID", str(pid), "/T", "/F"],
                    check=False,
                    capture_output=True,
                    text=True,
                )
            else:
                raise OSError("unsupported platform")
        except Exception:
            pass
        try:
            PID_FILE.unlink()
        except FileNotFoundError:
            pass
        stopped_pid = pid

    if REQUESTS_DIR.exists():
        for request_file in REQUESTS_DIR.glob("*.json"):
            task_name = task_name_for_target(request_file.stem)
            subprocess.run(["schtasks", "/End", "/TN", task_name], check=False, capture_output=True, text=True)
            subprocess.run(["schtasks", "/Change", "/TN", task_name, "/DISABLE"], check=False, capture_output=True, text=True)
    if stopped_pid:
        print(f"wechat_listener stop signal sent (pid {stopped_pid})")
    else:
        print("wechat_listener stop signal written")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
