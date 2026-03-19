#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Entry point executed by the Windows scheduled task."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

from wechat_listener_request import ROOT, build_runtime_config, load_request


def main() -> int:
    parser = argparse.ArgumentParser(description="Run one WeChat listener request.")
    parser.add_argument("--request", required=True, help="Path to the request json file")
    args = parser.parse_args()

    request = load_request(args.request)
    config_path = build_runtime_config(
        base_config_path=Path(request.get("config") or (ROOT / "config.yaml")),
        target=str(request["target"]),
        is_group=bool(request.get("is_group", False)),
        duration=request.get("duration"),
        auto_reply=bool(request.get("auto_reply", False)),
        reply_to_user=list(request.get("reply_to_user") or []),
        reply_message=request.get("reply_message"),
        mention_token=str(request.get("mention_token") or "@小小范"),
        session_key=request.get("session_key"),
    )
    listener_script = ROOT / "scripts" / "wechat_listener.py"
    cmd = [sys.executable, "-u", str(listener_script), "--config", config_path, "--no-close"]
    cmd.extend(["--duration", str(request.get("duration") or "forever")])
    log_file = Path(os.path.expanduser(r"~\.openclaw\logs\wechat-listener-task.log"))
    log_file.parent.mkdir(parents=True, exist_ok=True)

    with log_file.open("a", encoding="utf-8") as lf:
        lf.write(f"[listener_task_entry] start target={request.get('target')}\n")
        lf.flush()
        proc = subprocess.Popen(
            cmd,
            cwd=str(ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            sys.stdout.write(line)
            sys.stdout.flush()
            lf.write(line)
            lf.flush()
        return proc.wait()


if __name__ == "__main__":
    raise SystemExit(main())
