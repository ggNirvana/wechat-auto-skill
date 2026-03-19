#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Start the WeChat listener for one explicit target."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
import time

from wechat_listener_request import build_runtime_config


ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = ROOT / "runtime"
LOG_DIR = ROOT / "runtime" / "logs"
STOP_FILE = RUNTIME_DIR / "wechat_listener.stop"
PID_FILE = RUNTIME_DIR / "wechat_listener.pid"
LAUNCHER_FILE = RUNTIME_DIR / "launch_wechat_listener.cmd"
def main() -> int:
    parser = argparse.ArgumentParser(
        description="Start the WeChat listener for one contact/group using an independent window."
    )
    parser.add_argument("--target", required=True, help="WeChat contact or group name")
    parser.add_argument(
        "--is-group",
        action="store_true",
        help="Treat the target as a group chat and only respond when mentioned",
    )
    parser.add_argument("--duration", default=None, help="Override duration, e.g. 24h")
    parser.add_argument(
        "--auto-reply",
        action="store_true",
        help="Enable local fallback auto reply for this target",
    )
    parser.add_argument(
        "--reply-to-user",
        action="append",
        default=[],
        help="Optional trigger username for local fallback auto reply; repeatable",
    )
    parser.add_argument(
        "--reply-message",
        default=None,
        help="Optional fallback reply message stored as default_reply",
    )
    parser.add_argument(
        "--config",
        default=str(ROOT / "config.yaml"),
        help="Base config file path",
    )
    parser.add_argument(
        "--foreground",
        action="store_true",
        help="Run in foreground instead of background",
    )
    parser.add_argument(
        "--mention-token",
        default="@小小范",
        help="Mention token required for group replies",
    )
    parser.add_argument(
        "--session-key",
        default=None,
        help="Optional explicit hook session key; defaults to hook:wechat:<target>",
    )
    args = parser.parse_args()

    base_config_path = Path(args.config)
    if not base_config_path.exists():
        print(f"配置文件不存在: {base_config_path}", file=sys.stderr)
        return 1

    temp_config_path = build_runtime_config(
        base_config_path=base_config_path,
        target=args.target,
        is_group=args.is_group,
        duration=args.duration,
        auto_reply=args.auto_reply,
        reply_to_user=args.reply_to_user,
        reply_message=args.reply_message,
        mention_token=args.mention_token,
        session_key=args.session_key,
    )

    listener_script = ROOT / "scripts" / "wechat_listener.py"
    cmd = [sys.executable, "-u", str(listener_script), "--config", temp_config_path, "--no-close"]
    if args.duration:
        cmd.extend(["--duration", args.duration])
    else:
        cmd.extend(["--duration", "forever"])

    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    if STOP_FILE.exists():
        STOP_FILE.unlink()

    if args.foreground:
        completed = subprocess.run(cmd, cwd=str(ROOT))
        return completed.returncode

    if sys.platform == "win32":
        quoted = subprocess.list2cmdline(cmd)
        title = f"WeChat Listener - {args.target}"
        launcher_lines = [
            "@echo off",
            f"title {title}",
            f"cd /d \"{ROOT}\"",
            f"{quoted}",
            "echo.",
            "echo Listener exited. Press any key to close this window.",
            "pause >nul",
        ]
        LAUNCHER_FILE.write_text("\r\n".join(launcher_lines) + "\r\n", encoding="utf-8")
        launch_cmd = [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            f'Start-Process -FilePath "cmd.exe" -ArgumentList "/k","{LAUNCHER_FILE}" -WorkingDirectory "{ROOT}" -WindowStyle Normal',
        ]
        launched = subprocess.run(launch_cmd, cwd=str(ROOT), capture_output=True, text=True, timeout=15)
        if launched.returncode != 0:
            if launched.stderr.strip():
                print(launched.stderr.strip(), file=sys.stderr)
            return launched.returncode
        time.sleep(1.5)
        probe = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                f"Get-CimInstance Win32_Process | Where-Object {{ $_.Name -eq 'cmd.exe' -and $_.CommandLine -like '*{LAUNCHER_FILE.name}*' }} | Sort-Object CreationDate -Descending | Select-Object -First 1 -ExpandProperty ProcessId",
            ],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            timeout=15,
        )
        pid_text = probe.stdout.strip()
        if pid_text.isdigit():
            PID_FILE.write_text(pid_text, encoding="utf-8")
            print(f"wechat_listener started in visible cmd window (pid {pid_text})")
        else:
            PID_FILE.write_text("", encoding="utf-8")
            print("wechat_listener launched in visible cmd window")
        return 0
    else:
        proc = subprocess.Popen(
            cmd,
            cwd=str(ROOT),
        )
    PID_FILE.write_text(str(proc.pid), encoding="utf-8")
    print(f"wechat_listener started in background (pid {proc.pid})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
