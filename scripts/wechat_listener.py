#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
微信消息监听脚本 - 监听指定群消息并通过 Webhook 推送到 OpenClaw
支持接收 OpenClaw 返回的回复内容并自动发送

使用方法：
    python scripts/wechat_listener.py --config config.yaml

配置文件示例 (config.yaml)：
    duration: "24h"
    close_weixin: false
    openclaw_webhook_url: "http://127.0.0.1:18789/hooks/agent"
    openclaw_token: "your-token"
    groups:
      - name: "群名称1"
        auto_reply: false
      - name: "群名称2"
        auto_reply: true
        reply_to_user: "你的微信名"

OpenClaw Webhook 返回格式：
    {
        "reply": "回复内容",
        "at": ["成员1", "成员2"]  # 可选，要@的成员列表
    }
    或直接返回字符串作为回复内容
"""
from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
import threading
import time
import urllib.request
import urllib.error
from datetime import datetime
from typing import Any, Optional
import pyautogui

try:
    import yaml
except ImportError:
    print("请安装 PyYAML: pip install pyyaml")
    sys.exit(1)

try:
    from pyweixin import Navigator, GlobalConfig, Messages, Tools
    from pyweixin.WeChatAuto import AutoReply
    from pyweixin.Uielements import Edits, Main_window, SideBar
    from pyweixin.WinSettings import SystemSettings
    from pyweixin.utils import At
    from pyweixin.Errors import NoSuchFriendError
except ImportError as e:
    print("请先安装 pyweixin: pip install -e /path/to/pywechat")
    print(f"错误: {e}")
    sys.exit(1)

from wechat_listener_request import ensure_target_session_key

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUNTIME_DIR = os.path.join(ROOT, "runtime")
STOP_FILE = os.path.join(RUNTIME_DIR, "wechat_listener.stop")
RECENT_SENT_FILE = os.path.join(RUNTIME_DIR, "recent_sent_messages.json")
SESSION_STATE_FILE = os.path.join(RUNTIME_DIR, "session_rollover_state.json")
LISTENER_LOG_FILE = os.path.join(os.path.expanduser(r"~\.openclaw"), "logs", "wechat-listener.log")
OPENCLAW_HOME = os.path.expanduser(r"~\.openclaw")
SESSIONS_DIR = os.path.join(OPENCLAW_HOME, "agents", "main", "sessions")
SESSIONS_INDEX_FILE = os.path.join(OPENCLAW_HOME, "agents", "main", "sessions", "sessions.json")
HOOK_SESSION_KEY = "hook:wechat"
SELF_SENDERS = {"我", "自己", "me", "self"}
MAX_SESSION_JSONL_BYTES = 32 * 1024
MAX_SESSION_JSONL_LINES = 40
_recent_sent_messages: dict[str, list[tuple[str, float]]] = {}
_session_rollover_state: dict[str, dict[str, Any]] = {}


def _log(tag: str, message: str) -> None:
    """统一日志输出"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{tag}] {message}"
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        safe_line = line.encode("gbk", errors="replace").decode("gbk", errors="replace")
        print(safe_line, flush=True)
    try:
        os.makedirs(os.path.dirname(LISTENER_LOG_FILE), exist_ok=True)
        with open(LISTENER_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def _parse_duration(duration_str: str) -> float:
    """解析时长字符串，返回秒数"""
    s = (duration_str or "24h").strip().lower()
    if s in {"forever", "infinite", "always"}:
        return -1
    multipliers = {"s": 1, "min": 60, "h": 3600, "d": 86400}
    for suffix, mult in multipliers.items():
        if s.endswith(suffix):
            try:
                return float(s[:-len(suffix)]) * mult
            except ValueError:
                pass
    try:
        return float(s)
    except ValueError:
        return 86400


def _stop_requested() -> bool:
    return os.path.exists(STOP_FILE)


def _coerce_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (list, tuple, set)):
        return " ".join(_coerce_text(item) for item in value if item is not None)
    if isinstance(value, dict):
        try:
            return json.dumps(value, ensure_ascii=False, sort_keys=True)
        except Exception:
            return str(value)
    return str(value)


def _normalize_message_text(text: Any) -> str:
    return " ".join(_coerce_text(text).strip().split())


_SENDER_INLINE_PATTERN = re.compile(r"^\s*([^:\n：]{1,40})[：:]\s*(.+)$", re.S)
_SENDER_BRACKET_PATTERN = re.compile(r"^\s*\[([^\]\n]{1,40})\]\s*(.+)$", re.S)
_SENDER_TIME_PATTERN = re.compile(r"^\s*([^\n]{1,32}?)\s+(\d{1,2}:\d{2})\s+(.+)$", re.S)
_LEADING_TIME_PATTERN = re.compile(r"^\s*(\d{1,2}:\d{2})\s+(.+)$", re.S)
_SESSION_ITEM_SENDER_PATTERN = re.compile(r"^\s*([^:：]{1,40})[:：]\s*(.+)$")
_CLOCK_PATTERN = re.compile(r"^\d{1,2}:\d{2}$")
_UNKNOWN_SENDER_LOG_BUDGET: dict[str, int] = {}


def _is_sender_candidate(name: str) -> bool:
    normalized = _normalize_message_text(name)
    if not normalized or len(normalized) > 40:
        return False
    if _CLOCK_PATTERN.fullmatch(normalized):
        return False
    return True


def _extract_sender_and_message(candidate_text: Any) -> tuple[str, str]:
    raw = _coerce_text(candidate_text).strip()
    if not raw:
        return "", ""

    lines = [line.strip() for line in raw.replace("\r\n", "\n").split("\n") if line.strip()]
    if len(lines) >= 2 and _is_sender_candidate(lines[0]):
        return _normalize_message_text(lines[0]), _normalize_message_text(" ".join(lines[1:]))

    match = _SENDER_INLINE_PATTERN.match(raw)
    if match and _is_sender_candidate(match.group(1)):
        return _normalize_message_text(match.group(1)), _normalize_message_text(match.group(2))

    match = _SENDER_BRACKET_PATTERN.match(raw)
    if match and _is_sender_candidate(match.group(1)):
        return _normalize_message_text(match.group(1)), _normalize_message_text(match.group(2))

    match = _SENDER_TIME_PATTERN.match(raw)
    if match and _is_sender_candidate(match.group(1)):
        return _normalize_message_text(match.group(1)), _normalize_message_text(match.group(3))

    match = _LEADING_TIME_PATTERN.match(raw)
    if match:
        return "", _normalize_message_text(match.group(2))

    return "", _normalize_message_text(raw)


def _extract_group_sender_and_message(text: Any, contexts: list[str]) -> tuple[str, str]:
    sender, message = _extract_sender_and_message(text)
    if sender:
        return sender, message

    for item in reversed(contexts[-8:]):
        ctx_sender, ctx_message = _extract_sender_and_message(item)
        if ctx_sender:
            return ctx_sender, message or ctx_message

    return "", message


def _is_mentioned(message_text: str, mention_token: str) -> bool:
    normalized_message = _normalize_message_text(message_text)
    normalized_token = _normalize_message_text(mention_token)
    if not normalized_token:
        return True
    if normalized_token in normalized_message:
        return True

    # Be tolerant of full-width @ and optional spaces between @ and nickname.
    token_name = normalized_token.lstrip("@＠").strip()
    if not token_name:
        return False
    mention_pattern = re.compile(
        rf"[@＠]\s*{re.escape(token_name)}(?=$|[\s，。！？!?,:：、])"
    )
    return mention_pattern.search(normalized_message) is not None


def _log_unknown_sender_samples(group_name: str, text: Any, contexts: list[str]) -> None:
    count = _UNKNOWN_SENDER_LOG_BUDGET.get(group_name, 0)
    if count >= 3:
        return
    _UNKNOWN_SENDER_LOG_BUDGET[group_name] = count + 1
    snippet = _normalize_message_text(text)[:120]
    ctx_samples = [_normalize_message_text(c)[:100] for c in contexts[-5:]]
    _log(
        "DEBUG",
        f"群聊发送者未识别 | 群: {group_name} | text={snippet} | ctx={ctx_samples}",
    )


def _guess_sender_from_session_item(group_name: str, incoming_text: str) -> str:
    try:
        main = Navigator.open_weixin(is_maximize=GlobalConfig.is_maximize)
        main.child_window(**SideBar().Chats).click_input()
        session_list = main.child_window(**Main_window().SessionList)
        group_item = session_list.child_window(
            auto_id=f"session_item_{group_name}",
            control_type="ListItem",
        )
        if not group_item.exists(timeout=0.6):
            return ""
        raw = _coerce_text(group_item.window_text())
        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        if len(lines) < 2:
            return ""
        match = _SESSION_ITEM_SENDER_PATTERN.match(lines[1])
        if not match:
            return ""
        sender = _normalize_message_text(match.group(1))
        snippet = _normalize_message_text(match.group(2))
        target = _normalize_message_text(incoming_text)
        if not sender:
            return ""
        if snippet and (snippet in target or target in snippet):
            _log("DEBUG", f"通过会话列表识别发送者 | 群: {group_name} | sender={sender} | snippet={snippet[:80]}")
            return sender
        return ""
    except Exception:
        return ""


def _human_pause(min_seconds: float = 0.2, max_seconds: float = 0.6) -> None:
    time.sleep(random.uniform(min_seconds, max_seconds))


def _open_separate_dialog_window_robust(
    friend: str,
    *,
    window_minimize: bool = False,
    close_weixin: bool = False,
    retries: int = 3,
):
    """Open a contact/group chat window with a local fallback for flaky search_list lookups."""
    last_error: Exception | None = None
    sidebar = SideBar()
    main_ui = Main_window()

    for attempt in range(1, retries + 1):
        try:
            return Navigator.open_seperate_dialog_window(
                friend=friend,
                window_minimize=window_minimize,
                close_weixin=close_weixin,
            )
        except Exception as exc:
            last_error = exc
            _log("WARN", f"标准打开窗口失败，尝试兜底搜索 | target={friend} | attempt={attempt}/{retries} | error={exc}")
            try:
                main_window = Navigator.open_weixin(is_maximize=GlobalConfig.is_maximize)
                main_window.child_window(**sidebar.Chats).click_input()
                _human_pause(0.4, 0.8)
                session_list = main_window.child_window(**main_ui.SessionList)
                search = main_window.descendants(**main_ui.Search)[0]
                search.click_input()
                _human_pause(0.15, 0.35)
                search.set_text("")
                _human_pause(0.15, 0.35)
                search.set_text(friend)
                _human_pause(0.8, 1.4)

                search_results = main_window.child_window(**main_ui.SearchResult)
                if not search_results.exists(timeout=1.5):
                    raise NoSuchFriendError(f"搜索结果列表不存在: {friend}")

                result_items = [
                    item for item in search_results.children(control_type="ListItem")
                    if _normalize_message_text(item.window_text()) == _normalize_message_text(friend)
                ]
                if not result_items:
                    raise NoSuchFriendError(f"搜索结果中未找到目标: {friend}")

                result_items[0].click_input()
                _human_pause(0.5, 1.0)

                selected_items = [
                    item for item in session_list.children(control_type="ListItem")
                    if item.is_selected()
                ]
                if selected_items:
                    selected_items[0].double_click_input()
                else:
                    result_items[0].double_click_input()
                _human_pause(0.6, 1.2)

                dialog_window = Tools.move_window_to_center(Window={"class_name": "mmui::ChatSingleWindow", "title": f"{friend}"})
                if window_minimize:
                    dialog_window.minimize()
                if close_weixin:
                    main_window.close()
                return dialog_window
            except Exception as fallback_exc:
                last_error = fallback_exc
                _log("WARN", f"兜底打开窗口失败 | target={friend} | attempt={attempt}/{retries} | error={fallback_exc}")
                _human_pause(0.8, 1.6)
    raise last_error or RuntimeError(f"无法打开窗口: {friend}")


def _load_recent_sent_messages() -> None:
    global _recent_sent_messages
    if not os.path.exists(RECENT_SENT_FILE):
        return
    try:
        with open(RECENT_SENT_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        loaded: dict[str, list[tuple[str, float]]] = {}
        now = time.time()
        for group, items in (raw or {}).items():
            bucket: list[tuple[str, float]] = []
            for item in items or []:
                if not isinstance(item, dict):
                    continue
                msg = _normalize_message_text(item.get("message"))
                ts = item.get("ts")
                if msg and isinstance(ts, (int, float)) and now - float(ts) <= 180:
                    bucket.append((msg, float(ts)))
            if bucket:
                loaded[group] = bucket
        if loaded:
            _recent_sent_messages.update(loaded)
    except Exception:
        pass


def _save_recent_sent_messages() -> None:
    try:
        os.makedirs(RUNTIME_DIR, exist_ok=True)
        payload = {
            group: [{"message": msg, "ts": ts} for msg, ts in bucket]
            for group, bucket in _recent_sent_messages.items()
            if bucket
        }
        with open(RECENT_SENT_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _load_session_rollover_state() -> None:
    global _session_rollover_state
    if _session_rollover_state:
        return
    if not os.path.exists(SESSION_STATE_FILE):
        return
    try:
        with open(SESSION_STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            _session_rollover_state = data
    except Exception:
        _session_rollover_state = {}


def _save_session_rollover_state() -> None:
    try:
        os.makedirs(RUNTIME_DIR, exist_ok=True)
        with open(SESSION_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(_session_rollover_state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _get_session_stats(session_key: str) -> tuple[int, int]:
    session_file = _resolve_session_file(session_key)
    if not session_file or not os.path.exists(session_file):
        return 0, 0
    try:
        size_bytes = os.path.getsize(session_file)
    except OSError:
        size_bytes = 0
    rows = _read_jsonl_lines(session_file)
    return size_bytes, len(rows)


def _next_rollover_session_key(base_session_key: str, group_name: str) -> str:
    _load_session_rollover_state()
    normalized_base = ensure_target_session_key(base_session_key, group_name)
    state_key = f"{group_name}|{normalized_base}"
    state = _session_rollover_state.get(state_key) or {}
    generation = int(state.get("generation") or 1)
    current_key = str(state.get("current_key") or normalized_base)
    size_bytes, line_count = _get_session_stats(current_key)
    should_roll = (
        current_key == normalized_base and (size_bytes >= MAX_SESSION_JSONL_BYTES or line_count >= MAX_SESSION_JSONL_LINES)
    ) or (
        current_key != normalized_base and (size_bytes >= MAX_SESSION_JSONL_BYTES or line_count >= MAX_SESSION_JSONL_LINES)
    )
    if should_roll:
        generation += 1
        current_key = f"{normalized_base}:r{generation}"
        _log(
            "OPENCLAW",
            f"会话上下文达到阈值，切换新会话 | group={group_name} | session_key={current_key} | prev_bytes={size_bytes} | prev_lines={line_count}",
        )
    _session_rollover_state[state_key] = {
        "generation": generation,
        "current_key": current_key,
        "updated_at": datetime.now().isoformat(),
    }
    _save_session_rollover_state()
    return current_key


def _remember_sent_message(group: str, text: str) -> None:
    normalized = _normalize_message_text(text)
    if not normalized:
        return
    now = time.time()
    bucket = _recent_sent_messages.setdefault(group, [])
    bucket.append((normalized, now))
    _recent_sent_messages[group] = [
        (msg, ts) for msg, ts in bucket if now - ts <= 180
    ]
    _save_recent_sent_messages()


def _was_recently_sent(group: str, text: str) -> bool:
    _load_recent_sent_messages()
    normalized = _normalize_message_text(text)
    if not normalized:
        return False
    now = time.time()
    bucket = _recent_sent_messages.get(group, [])
    kept: list[tuple[str, float]] = []
    matched = False
    for msg, ts in bucket:
        if now - ts <= 180:
            kept.append((msg, ts))
            if msg == normalized:
                matched = True
    _recent_sent_messages[group] = kept
    _save_recent_sent_messages()
    return matched


def _send_to_webhook(
    message: Any,
    webhook_url: str,
    token: str,
    group: str,
    sender: Any,
    session_key: str,
    timeout: int = 30,
) -> Optional[dict]:
    """
    发送消息到 OpenClaw Webhook 并返回排队确认

    Returns:
        dict acknowledgement such as {"ok": true, "runId": "..."} or None on failure
    """
    if not webhook_url:
        return None

    # 把入站微信消息投递到 OpenClaw 独立会话，让模型只产出“要发给微信的文本”，
    # 由本地监听器负责真正发送，避免模型口头说“已发送”但并未执行。
    message_text = _normalize_message_text(message)
    sender_text = _normalize_message_text(sender)
    formatted = (
        "你正在为微信联系人或群生成自动回复。\n"
        "请只输出一段要发回微信的纯文本回复，不要解释，不要使用 markdown，不要加引号，"
        "不要说你已经发送，也不要描述你的搜索过程。\n"
        "如果对方要求搜索/查询，请先自行搜索后，再只输出最终要发给对方的回复内容。\n"
        f"目标名称：{group}\n"
        f"发送者：{sender_text}\n"
        f"消息内容：{message_text}"
    )

    body = {
        "message": formatted,
        "name": "WeChatListener",
        "wakeMode": "now",
        "deliver": True,
        "sessionKey": session_key or HOOK_SESSION_KEY,
        "metadata": {
            "group": group,
            "sender": sender_text,
            "timestamp": datetime.now().isoformat(),
        }
    }

    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    headers = {
        "Content-Type": "application/json",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"

    req = urllib.request.Request(
        webhook_url.rstrip("/"),
        data=data,
        headers=headers,
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status == 200:
                _log("WEBHOOK", f"已推送到 OpenClaw | 群: {group} | 发送者: {sender_text}")

                # hooks/agent 返回的是排队确认，不是最终回复正文
                try:
                    response_data = resp.read().decode("utf-8")
                    if response_data:
                        response_json = json.loads(response_data)
                        _log("WEBHOOK", f"收到确认: {response_data[:200]}")
                        return response_json
                except (json.JSONDecodeError, UnicodeDecodeError):
                    pass
                return None
            else:
                _log("ERROR", f"Webhook 返回状态码: {resp.status}")
                return None
    except urllib.error.URLError as e:
        _log("ERROR", f"Webhook 连接失败: {e}")
        return None
    except Exception as e:
        _log("ERROR", f"Webhook 发送异常: {e}")
        return None


def _resolve_session_file(session_key: str) -> Optional[str]:
    if not os.path.exists(SESSIONS_INDEX_FILE):
        return None
    try:
        with open(SESSIONS_INDEX_FILE, "r", encoding="utf-8") as f:
            sessions = json.load(f) or {}
    except Exception:
        return None

    candidates = [
        session_key,
        f"agent:main:{session_key}",
        f"agent:main:{session_key.lower()}",
    ]
    for key in candidates:
        entry = sessions.get(key)
        if isinstance(entry, dict):
            session_id = entry.get("sessionId")
            if isinstance(session_id, str):
                session_id_file = os.path.join(SESSIONS_DIR, f"{session_id}.jsonl")
                if os.path.exists(session_id_file):
                    return session_id_file
            session_file = entry.get("sessionFile")
            if isinstance(session_file, str) and os.path.exists(session_file):
                return session_file

    lowered = session_key.lower()
    for key, entry in sessions.items():
        if not isinstance(key, str) or not isinstance(entry, dict):
            continue
        if key.lower().endswith(lowered):
            session_id = entry.get("sessionId")
            if isinstance(session_id, str):
                session_id_file = os.path.join(SESSIONS_DIR, f"{session_id}.jsonl")
                if os.path.exists(session_id_file):
                    return session_id_file
            session_file = entry.get("sessionFile")
            if isinstance(session_file, str) and os.path.exists(session_file):
                return session_file
    return None


def _read_jsonl_lines(path: str) -> list[dict]:
    rows: list[dict] = []
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    except Exception:
        return []
    return rows


def _parse_row_timestamp(row: dict) -> Optional[float]:
    timestamp = row.get("timestamp")
    if not isinstance(timestamp, str):
        return None
    try:
        return datetime.fromisoformat(timestamp.replace("Z", "+00:00")).timestamp()
    except ValueError:
        return None


def _extract_assistant_text(row: dict) -> Optional[str]:
    message = row.get("message")
    if not isinstance(message, dict):
        return None
    if message.get("role") != "assistant":
        return None
    content = message.get("content")
    if not isinstance(content, list):
        return None
    parts: list[str] = []
    for item in content:
        if not isinstance(item, dict):
            continue
        if item.get("type") != "text":
            continue
        text = _normalize_message_text(item.get("text"))
        if text:
            parts.append(text)
    if not parts:
        return None
    return "\n".join(parts).strip()


def _wait_for_openclaw_reply(
    session_key: str,
    marker_count: int,
    timeout: int = 180,
    request_started_at: Optional[float] = None,
) -> Optional[str]:
    deadline = time.time() + timeout
    last_session_file = None
    _log("OPENCLAW", f"等待 OpenClaw 回复 | session_key={session_key} | timeout={timeout}s | marker_count={marker_count}")
    while time.time() < deadline:
        session_file = _resolve_session_file(session_key)
        if not session_file:
            time.sleep(1.0)
            continue
        if session_file != last_session_file:
            _log("OPENCLAW", f"已定位会话文件 | session_key={session_key} | file={session_file}")
            last_session_file = session_file
        rows = _read_jsonl_lines(session_file)
        if not rows:
            time.sleep(1.0)
            continue
        start_index = marker_count if len(rows) > marker_count else 0
        candidate_rows = rows[start_index:]
        if request_started_at is not None:
            candidate_rows = [
                row for row in candidate_rows
                if (_parse_row_timestamp(row) or 0) >= request_started_at
            ]
            if not candidate_rows and start_index != 0:
                candidate_rows = [
                    row for row in rows
                    if (_parse_row_timestamp(row) or 0) >= request_started_at
                ]
        for row in candidate_rows:
            text = _extract_assistant_text(row)
            if text:
                _log("OPENCLAW", f"捕获到回复文本 | session_key={session_key} | text={text[:160]}")
                return text
        time.sleep(1.0)
    _log("ERROR", f"等待 OpenClaw 回复超时 | session_key={session_key}")
    return None


def _parse_reply(response: Any) -> tuple[Optional[str], list]:
    """
    解析 OpenClaw 返回的回复内容

    Args:
        response: OpenClaw 返回的响应，可能是 dict 或字符串

    Returns:
        (reply_text, at_members) - 回复内容和要@的成员列表
    """
    if response is None:
        return None, []

    # 如果是字符串，直接作为回复
    if isinstance(response, str):
        text = response.strip()
        return text if text else None, []

    # 如果是字典
    if isinstance(response, dict):
        # 尝试多种可能的字段名
        reply = None
        for key in ["reply", "response", "message", "content", "text", "answer"]:
            if key in response and response[key]:
                reply = response[key]
                break

        if reply is None:
            # 尝试 data 字段
            if "data" in response:
                data = response["data"]
                if isinstance(data, str):
                    reply = data
                elif isinstance(data, dict):
                    for key in ["reply", "message", "content", "text"]:
                        if key in data and data[key]:
                            reply = data[key]
                            break

        if reply is None:
            return None, []

        # 获取要@的成员
        at_members = []
        for key in ["at", "at_members", "mention", "mentions"]:
            if key in response:
                at_data = response[key]
                if isinstance(at_data, str):
                    at_members = [a.strip() for a in at_data.split(",") if a.strip()]
                elif isinstance(at_data, list):
                    at_members = [str(a).strip() for a in at_data if a]
                break

        return reply, at_members

    return None, []


def _send_reply(group: str, reply: str, at_members: list) -> bool:
    """
    发送回复到微信群

    Returns:
        True if sent successfully, False otherwise
    """
    try:
        _log("REPLY", f"准备发送回复 | 目标={group} | at={at_members} | text={_normalize_message_text(reply)[:160]}")
        # Reuse the already-open independent chat window first. This matches
        # the original listener design and avoids flaky main-window search.
        window = _group_windows.get(group)
        if window is not None:
            edit_area = window.child_window(**Edits().CurrentChatEdit)
            if edit_area.exists(timeout=0.5):
                _human_pause(0.25, 0.7)
                edit_area.click_input()
                _human_pause(0.15, 0.4)
                edit_area.set_text("")
                if at_members:
                    _human_pause(0.2, 0.5)
                    At(window, at_members)
                _human_pause(0.25, 0.65)
                SystemSettings.copy_text_to_clipboard(reply)
                _human_pause(0.2, 0.5)
                pyautogui.hotkey("ctrl", "v", _pause=False)
                _human_pause(0.35, 0.9)
                pyautogui.hotkey("alt", "s", _pause=False)
                _human_pause(0.3, 0.8)
                _remember_sent_message(group, reply)
                at_str = f" and @{at_members}" if at_members else ""
                _log("REPLY", f"已发送回复到【{group}】{at_str}: {reply[:50]}...")
                return True

        _log("REPLY", f"独立窗口不可用，回退到搜索发送 | 目标={group}")
        Messages.send_messages_to_friend(
            friend=group,
            messages=[reply],
            at_members=at_members,
            search_pages=8,
            close_weixin=False,
        )
        _human_pause(0.4, 1.0)
        _remember_sent_message(group, reply)
        at_str = f" and @{at_members}" if at_members else ""
        _log("REPLY", f"已发送回复到【{group}】{at_str}: {reply[:50]}...")
        return True
    except Exception as e:
        _log("ERROR", f"发送回复失败: {e}")
        return False


# 全局变量：存储群窗口引用
_group_windows: dict[str, Any] = {}


def make_callback(
    group_name: str,
    webhook_url: str,
    token: str,
    session_key: str,
    is_group: bool = False,
    mention_token: str = "@小小范",
    auto_reply: bool = False,
    reply_to_user: list = None,
    default_reply: str = "收到，稍后处理。",
):
    """创建群消息回调函数"""
    watch_users = reply_to_user or []

    def callback(text: Any, sender: Any = "对方", receiver: Any = "我"):
        contexts: list[str] = []
        sender_hint: Any = sender
        receiver_hint: Any = receiver
        if isinstance(sender, (list, tuple, set)):
            contexts = [_coerce_text(item) for item in sender if _normalize_message_text(item)]
            sender_hint = ""
            receiver_hint = "我"
        elif isinstance(receiver, (list, tuple, set)):
            contexts = [_coerce_text(item) for item in receiver if _normalize_message_text(item)]
            receiver_hint = "我"

        normalized_text = _normalize_message_text(text)
        if is_group:
            parsed_sender, parsed_message = _extract_group_sender_and_message(text, contexts)
            normalized_text = parsed_message or normalized_text
            sender_text = _normalize_message_text(parsed_sender or sender_hint or "")
            if not sender_text:
                sender_text = _guess_sender_from_session_item(group_name, normalized_text)
            if not sender_text:
                sender_text = "群成员"
                _log_unknown_sender_samples(group_name, text, contexts)
        else:
            sender_text = _normalize_message_text(sender_hint or "对方")
        receiver_text = _normalize_message_text(receiver_hint or "我")
        mention_text = _normalize_message_text(mention_token)

        if sender_text.lower() in SELF_SENDERS or sender_text in SELF_SENDERS:
            _log("SKIP", f"忽略自己发出的消息 | 群: {group_name} | 内容: {normalized_text[:60]}")
            return None
        if _was_recently_sent(group_name, normalized_text):
            _log("SKIP", f"忽略刚刚由监听器发送的回环消息 | 群: {group_name} | 内容: {normalized_text[:60]}")
            return None
        if is_group and mention_text and not _is_mentioned(normalized_text, mention_text):
            _log("SKIP", f"群聊消息未命中提及词，忽略回复 | 群: {group_name} | 提及词: {mention_text} | 内容: {normalized_text[:80]}")
            return None
        _log("MESSAGE", f"捕获消息 | 群: {group_name} | 发送者: {sender_text or '对方'} | 接收者: {receiver_text or '我'} | 内容: {normalized_text[:80]}")

        effective_session_key = _next_rollover_session_key(session_key, group_name)
        session_file = _resolve_session_file(effective_session_key)
        marker_count = len(_read_jsonl_lines(session_file)) if session_file else 0
        request_started_at = time.time()

        # 发送到 Webhook 并获取回复
        response = _send_to_webhook(
            normalized_text,
            webhook_url,
            token,
            group_name,
            sender_text or "对方",
            effective_session_key,
        )

        if response and response.get("ok") and response.get("runId"):
            _log("OPENCLAW", f"消息已进入 OpenClaw 队列 | runId={response.get('runId')}")
            reply_text = _wait_for_openclaw_reply(
                effective_session_key,
                marker_count,
                request_started_at=request_started_at,
            )
            if reply_text:
                _log("OPENCLAW", f"收到 OpenClaw 回复文本 | session_key={effective_session_key}: {reply_text[:120]}")
                sent = _send_reply(group_name, reply_text, [])
                if not sent:
                    _log("ERROR", f"微信发送失败 | 群: {group_name} | session_key={effective_session_key}")
            else:
                _log("ERROR", f"等待 OpenClaw 回复超时 | 群: {group_name} | session_key={effective_session_key}")
            return None

        # 如果 OpenClaw 没有回复，检查是否需要本地自动回复
        if auto_reply and watch_users:
            t = (normalized_text or "").strip().lower()
            for user in watch_users:
                if user and (user in normalized_text or user.lower() in t):
                    if "在吗" in normalized_text or "在不在" in normalized_text:
                        return "在的，请说~"
                    return default_reply
        return None

    return callback


def main() -> None:
    parser = argparse.ArgumentParser(
        description="微信消息监听 - 监听指定群并推送到 OpenClaw，支持自动回复"
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="配置文件路径 (YAML)",
    )
    parser.add_argument(
        "--duration",
        default=None,
        help="覆盖配置中的监听时长，如 30min、24h",
    )
    parser.add_argument(
        "--no-close",
        action="store_true",
        help="结束后不关闭微信",
    )
    parser.add_argument(
        "--webhook-url",
        default=None,
        help="覆盖配置中的 OpenClaw Webhook URL",
    )
    parser.add_argument(
        "--token",
        default=None,
        help="覆盖配置中的 OpenClaw Token",
    )
    parser.add_argument(
        "--default-reply",
        default=None,
        help="覆盖配置中的默认回复",
    )
    args = parser.parse_args()

    # 加载配置
    config_path = args.config
    if not os.path.isabs(config_path):
        config_path = os.path.join(os.getcwd(), config_path)

    if not os.path.isfile(config_path):
        print(f"配置文件不存在: {config_path}")
        print("请复制 config.example.yaml 为 config.yaml 并修改配置")
        sys.exit(1)

    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    # 解析配置
    duration = args.duration or raw.get("duration", "24h")
    duration_seconds = _parse_duration(duration)
    close_weixin = not args.no_close and raw.get("close_weixin", True)
    GlobalConfig.close_weixin = close_weixin

    webhook_url = args.webhook_url or raw.get("openclaw_webhook_url", "")
    token = args.token or raw.get("openclaw_token", "")
    default_reply = args.default_reply or raw.get("default_reply", "收到，稍后处理。")

    groups_cfg: list[dict] = raw.get("groups") or []
    if not groups_cfg:
        print("配置中 groups 为空，请至少配置一个群")
        sys.exit(1)

    _log("INFO", "=" * 50)
    _log("INFO", "微信消息监听服务")
    _log("INFO", "=" * 50)
    _log("INFO", f"配置文件: {config_path}")
    _log("INFO", f"监听时长: {duration} ({duration_seconds:.0f}秒)")
    _log("INFO", f"关闭微信: {close_weixin}")
    _log("INFO", f"Webhook: {webhook_url or '未配置'}")
    _log("INFO", f"自动回复: {'启用' if webhook_url else '仅推送'}")
    _log("INFO", f"监听群数: {len(groups_cfg)}")
    for g in groups_cfg:
        _log("INFO", f"  - {g.get('name', '未命名')}")
    _log("INFO", "请确保微信已登录")
    _log("INFO", "=" * 50)
    os.makedirs(RUNTIME_DIR, exist_ok=True)
    if os.path.exists(STOP_FILE):
        os.remove(STOP_FILE)

    # 打开所有群的独立窗口
    windows: dict[str, Any] = {}
    callbacks: dict[str, Any] = {}

    for i, g in enumerate(groups_cfg):
        name = g.get("name", "")
        if not name:
            continue

        _log("INFO", f"打开群【{name}】的独立窗口...")
        close_wx = GlobalConfig.close_weixin and (i == len(groups_cfg) - 1)

        try:
            _human_pause(0.6, 1.4)
            w = _open_separate_dialog_window_robust(
                friend=name,
                window_minimize=False,
                close_weixin=close_wx,
            )
            windows[name] = w
            _group_windows[name] = w

            # 创建回调
            reply_to_user = g.get("reply_to_user", "")
            if isinstance(reply_to_user, str):
                reply_to_user = [reply_to_user] if reply_to_user else []
            session_key = ensure_target_session_key(_normalize_message_text(g.get("session_key")), name)
            is_group = bool(g.get("is_group", False))
            mention_token = _normalize_message_text(g.get("mention_token")) or "@小小范"

            callbacks[name] = make_callback(
                group_name=name,
                webhook_url=webhook_url,
                token=token,
                session_key=session_key,
                is_group=is_group,
                mention_token=mention_token,
                auto_reply=g.get("auto_reply", False),
                reply_to_user=reply_to_user,
                default_reply=default_reply,
            )

            _log("INFO", f"群【{name}】窗口已打开 | session_key={session_key} | is_group={is_group}")
        except Exception as e:
            _log("ERROR", f"打开群【{name}】失败: {e}")

        _human_pause(0.5, 1.2)

    if not windows:
        _log("ERROR", "没有成功打开任何窗口")
        sys.exit(1)

    # 并发监听
    threads = []
    keep_running = duration_seconds < 0
    listen_span = "1h" if keep_running else duration

    def run_group(group_name: str):
        cb = callbacks.get(group_name)
        while not _stop_requested():
            w = windows.get(group_name)
            if w is None:
                try:
                    _log("INFO", f"重新打开群【{group_name}】的独立窗口...")
                    _human_pause(0.8, 1.8)
                    w = _open_separate_dialog_window_robust(
                        friend=group_name,
                        window_minimize=False,
                        close_weixin=False,
                    )
                    windows[group_name] = w
                    _group_windows[group_name] = w
                except Exception as e:
                    _log("ERROR", f"群【{group_name}】重开窗口失败: {e}")
                    _human_pause(1.5, 3.0)
                    continue
            try:
                AutoReply.auto_reply_to_friend(w, listen_span, cb, close_dialog_window=False)
                if not keep_running:
                    break
                _log("WARN", f"群【{group_name}】本轮监听结束，准备继续监听")
            except Exception as e:
                _log("ERROR", f"群【{group_name}】监听异常: {e}")
                windows[group_name] = None
                _group_windows[group_name] = None
            if keep_running and not _stop_requested():
                _human_pause(1.0, 2.2)

    for name in windows:
        t = threading.Thread(target=run_group, args=(name,), daemon=False)
        threads.append(t)
        t.start()
        _log("INFO", f"开始监听群【{name}】")

    try:
        while any(t.is_alive() for t in threads):
            if _stop_requested():
                _log("INFO", "检测到停止信号，正在结束监听...")
                break
            for t in threads:
                t.join(timeout=1)
    except KeyboardInterrupt:
        _log("INFO", "用户中断，正在退出...")

    _log("INFO", "监听结束")


if __name__ == "__main__":
    main()
