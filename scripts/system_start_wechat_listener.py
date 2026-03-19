#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Prepare a listener request and trigger the Windows scheduled task."""

from __future__ import annotations

import argparse
import subprocess
import sys

from wechat_listener_request import (
    ROOT,
    ensure_target_session_key,
    slugify_target,
    task_name_for_target,
    write_request,
)


def ensure_task(task_name: str, request_path: str) -> None:
    listener_cmd = ROOT / "scripts" / "listener_task_entry.cmd"
    cmd = f'cmd.exe /c ""{listener_cmd}" --request "{request_path}""'
    create = subprocess.run(
        [
            "schtasks",
            "/Create",
            "/TN",
            task_name,
            "/SC",
            "ONCE",
            "/ST",
            "00:00",
            "/RL",
            "LIMITED",
            "/IT",
            "/TR",
            cmd,
            "/F",
        ],
        capture_output=True,
        text=True,
    )
    if create.returncode != 0:
        message = (create.stderr or create.stdout or "").strip()
        raise RuntimeError(message or f"failed to create task {task_name}")

    change = subprocess.run(
        ["schtasks", "/Change", "/TN", task_name, "/ENABLE"],
        capture_output=True,
        text=True,
    )
    if change.returncode != 0:
        message = (change.stderr or change.stdout or "").strip()
        raise RuntimeError(message or f"failed to enable task {task_name}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Trigger the Windows task that starts a visible WeChat listener.")
    parser.add_argument("--target", required=True, help="WeChat contact or group name")
    parser.add_argument("--is-group", action="store_true", help="Treat target as a group chat")
    parser.add_argument("--duration", default="forever", help="Listener duration, defaults to forever")
    parser.add_argument("--auto-reply", action="store_true", help="Enable local fallback auto reply")
    parser.add_argument("--reply-to-user", action="append", default=[], help="Optional local fallback trigger user")
    parser.add_argument("--reply-message", default=None, help="Optional local fallback reply text")
    parser.add_argument("--mention-token", default="@小小范", help="Mention token required for group replies")
    parser.add_argument("--session-key", default=None, help="Optional explicit session key")
    parser.add_argument("--config", default=str(ROOT / "config.yaml"), help="Base config file path")
    args = parser.parse_args()

    payload = {
        "target": args.target,
        "is_group": args.is_group,
        "duration": args.duration,
        "auto_reply": args.auto_reply,
        "reply_to_user": args.reply_to_user,
        "reply_message": args.reply_message,
        "mention_token": args.mention_token,
        "session_key": ensure_target_session_key(args.session_key or f"hook:wechat:{slugify_target(args.target)}", args.target),
        "config": args.config,
    }
    request_path = write_request(payload)
    task_name = task_name_for_target(args.target)
    print(f"listener request written: {request_path}")
    try:
        ensure_task(task_name, str(request_path))
    except Exception as exc:
        sys.stderr.write(f"{exc}\n")
        return 1

    result = subprocess.run(
        ["schtasks", "/Run", "/TN", task_name],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        sys.stderr.write((result.stderr or result.stdout or "").strip() + "\n")
        return result.returncode
    print(f"scheduled task triggered: {task_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
