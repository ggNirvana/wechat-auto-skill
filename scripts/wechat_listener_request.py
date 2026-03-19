#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Shared helpers for preparing WeChat listener runtime requests."""

from __future__ import annotations

import json
import re
import tempfile
from pathlib import Path
from typing import Any

import yaml


ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = ROOT / "runtime"
REQUESTS_DIR = RUNTIME_DIR / "requests"


def slugify_target(name: str) -> str:
    cleaned = re.sub(r"\s+", "-", name.strip())
    cleaned = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff_-]+", "-", cleaned)
    cleaned = re.sub(r"-{2,}", "-", cleaned).strip("-")
    return cleaned or "target"


def ensure_target_session_key(session_key: str | None, target: str) -> str:
    target_slug = slugify_target(target)
    normalized = (session_key or "").strip()
    if not normalized or normalized == "hook:wechat":
        return f"hook:wechat:{target_slug}"
    if normalized.startswith("hook:wechat:"):
        return normalized
    return f"hook:wechat:{target_slug}"


def build_runtime_config(
    *,
    base_config_path: Path,
    target: str,
    is_group: bool = False,
    duration: str | None = None,
    auto_reply: bool = False,
    reply_to_user: list[str] | None = None,
    reply_message: str | None = None,
    mention_token: str = "@小小范",
    session_key: str | None = None,
) -> str:
    if not base_config_path.exists():
        raise FileNotFoundError(f"配置文件不存在: {base_config_path}")

    with base_config_path.open("r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}

    target_entry: dict[str, object] = {
        "name": target,
        "auto_reply": auto_reply,
        "is_group": is_group,
        "mention_token": mention_token,
        "session_key": ensure_target_session_key(session_key, target),
    }
    if reply_to_user:
        target_entry["reply_to_user"] = reply_to_user
    if reply_message:
        config["default_reply"] = reply_message
        target_entry["auto_reply"] = True

    config["groups"] = [target_entry]
    if duration:
        config["duration"] = duration

    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        suffix=".yaml",
        delete=False,
    ) as tmp:
        yaml.safe_dump(config, tmp, allow_unicode=True, sort_keys=False)
        return tmp.name


def request_file_for_target(target: str) -> Path:
    REQUESTS_DIR.mkdir(parents=True, exist_ok=True)
    return REQUESTS_DIR / f"{slugify_target(target)}.json"


def task_name_for_target(target: str) -> str:
    return f"OpenClaw WeChat Listener - {slugify_target(target)}"


def write_request(payload: dict[str, Any]) -> Path:
    target = str(payload.get("target") or "").strip()
    if not target:
        raise ValueError("payload.target is required")
    request_file = request_file_for_target(target)
    request_file.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return request_file


def load_request(path: str | Path | None = None) -> dict[str, Any]:
    request_file = Path(path) if path else None
    if request_file is None:
        raise FileNotFoundError("监听请求文件路径未提供")
    if not request_file.exists():
        raise FileNotFoundError(f"监听请求文件不存在: {request_file}")
    return json.loads(request_file.read_text(encoding="utf-8"))
