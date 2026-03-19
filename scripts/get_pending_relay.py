#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
查询并可选清除「待回传」：当收到群 B 的消息时，OpenClaw 先调用本脚本查是否有待回传；
若有则用 to_group、at 把该消息发回原群并 @，再带 --clear 调用一次清除。
输出格式：一行 JSON 或 to_group=... at=... 便于解析；无记录时退出码 1 且无 to_group。

参数：
  --supplier-group  目标群名称（之前转发消息到的群）
  --clear           查询后清除该条待回传记录
  --format          输出格式：json 或 kv（默认 json）
"""
import argparse
import json
import os
import sys

FILENAME = "wechat_pending_relay.json"

def _state_path():
    base = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(base)
    return os.path.join(skill_root, FILENAME)

def main():
    parser = argparse.ArgumentParser(description="查询/清除待回传")
    parser.add_argument("--supplier-group", required=True, help="目标群名称")
    parser.add_argument("--clear", action="store_true", help="查询后清除该条待回传")
    parser.add_argument("--format", choices=["json", "kv"], default="json", help="输出格式")
    args = parser.parse_args()

    path = _state_path()
    data = {}
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    key = args.supplier_group.strip()
    rel = data.get("pending_relays") or data
    pending = rel.get(key)
    if pending is None:
        sys.exit(1)
    if isinstance(pending, dict) and pending.get("queue"):
        head = pending["queue"][0]
        out = {"to_group": head.get("to_group"), "at": head.get("at"), "thread_id": head.get("thread_id"), "queue_len": len(pending["queue"])}
    else:
        out = pending
    if args.clear:
        rel.pop(key, None)
        data["pending_relays"] = rel
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    if args.format == "json":
        print(json.dumps(out, ensure_ascii=False))
    else:
        print(f"to_group={out.get('to_group')}\nat={out.get('at')}")


if __name__ == "__main__":
    main()
