#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
记录「待回传」：当 OpenClaw 把 A 群的消息转发到 B 群时，调用本脚本记录 (B, A, 原发送者)，
以便收到 B 群回复时能把内容发回 A 群并 @ 原发送者。

参数：
  --supplier-group  目标群名称（消息转发到的群）
  --to-group        原群名称（消息来源群）
  --at              回传时要 @ 的成员昵称（原发送者）
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
    parser = argparse.ArgumentParser(description="记录待回传：目标群 -> (原群, @谁)")
    parser.add_argument("--supplier-group", required=True, help="目标群名称")
    parser.add_argument("--to-group", required=True, help="要回传到的群（原消息来源群）")
    parser.add_argument("--at", required=True, help="回传时要 @ 的成员昵称（原发送者）")
    args = parser.parse_args()

    path = _state_path()
    data = {}
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    data[args.supplier_group.strip()] = {
        "to_group": args.to_group.strip(),
        "at": args.at.strip(),
    }
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"已记录: 【{args.supplier_group}】回复将回传到【{args.to_group}】并 @ {args.at}")
    except Exception as e:
        print(f"写入失败: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
