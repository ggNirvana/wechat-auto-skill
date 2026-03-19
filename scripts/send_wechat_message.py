#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
向指定微信群发送一条消息（可选 @ 成员）。供 OpenClaw 调用，实现「用微信发消息」能力。

依赖：pyweixin（pip install -e /path/to/pywechat）
"""
import argparse
import json
import os
import sys
import time


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUNTIME_DIR = os.path.join(ROOT, "runtime")
RECENT_SENT_FILE = os.path.join(RUNTIME_DIR, "recent_sent_messages.json")


def _normalize_message_text(text: str) -> str:
    return " ".join((text or "").strip().split())


def _remember_sent_message(group: str, text: str) -> None:
    normalized = _normalize_message_text(text)
    if not normalized:
        return
    now = time.time()
    payload: dict[str, list[dict[str, object]]] = {}
    if os.path.exists(RECENT_SENT_FILE):
        try:
            with open(RECENT_SENT_FILE, "r", encoding="utf-8") as f:
                payload = json.load(f) or {}
        except Exception:
            payload = {}
    bucket = payload.get(group, [])
    if not isinstance(bucket, list):
        bucket = []
    bucket.append({"message": normalized, "ts": now})
    payload[group] = [
        item for item in bucket
        if isinstance(item, dict)
        and isinstance(item.get("ts"), (int, float))
        and now - float(item["ts"]) <= 180
        and _normalize_message_text(str(item.get("message", "")))
    ]
    os.makedirs(RUNTIME_DIR, exist_ok=True)
    with open(RECENT_SENT_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def main():
    parser = argparse.ArgumentParser(description="向微信群发送消息（可选 @）")
    parser.add_argument("--group", required=True, help="群名称（与微信中显示一致）")
    parser.add_argument("--message", required=True, help="要发送的文本")
    parser.add_argument("--at", default="", help="要 @ 的群成员昵称，多个用逗号分隔")
    parser.add_argument("--close-weixin", action="store_true", help="发送后是否关闭微信窗口（默认不关闭）")
    args = parser.parse_args()

    try:
        from pyweixin import Navigator, Messages, GlobalConfig
    except ImportError as e:
        print("请先安装 pyweixin: pip install -e /path/to/pywechat", file=sys.stderr)
        sys.exit(1)

    GlobalConfig.close_weixin = getattr(args, "close_weixin", False)
    group_name = args.group.strip()
    message = args.message.strip()
    at_members = [x.strip() for x in args.at.split(",") if x.strip()]

    if not message:
        print("--message 不能为空", file=sys.stderr)
        sys.exit(1)

    try:
        Messages.send_messages_to_friend(
            friend=group_name,
            messages=[message],
            at_members=at_members,
            close_weixin=GlobalConfig.close_weixin,
        )
        _remember_sent_message(group_name, message)
        print(f"已发送到群【{group_name}】" + (f" 并 @ {at_members}" if at_members else ""))
    except Exception as e:
        print(f"发送失败: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
