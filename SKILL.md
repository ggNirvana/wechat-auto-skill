---
name: wechat-skill
description: 为 OpenClaw 提供微信联系人/群消息收发能力，支持发送消息到指定联系人或群、@ 成员、开启独立微信窗口监听、以及消息转发回传队列管理。用户说“监听某个联系人”“监听某个微信群”“打开独立微信窗口监听”“收到微信消息后自动回复”“给微信发消息”时使用。遇到“开始监听某人/某群”这类请求时，直接执行 `python scripts/system_start_wechat_listener.py --target "<目标名>"`，由该入口负责通过系统级启动器弹出独立可见的 cmd 黑框并启动监听。未明确说明是群聊时，默认按联系人处理，即 `is_group=False`；只有显式使用 `--is-group` 时才按群聊监听。群聊监听时，只有消息中出现 `@小小范` 才会转入 OpenClaw 并回复；不同目标绑定不同 hook session。
version: 1.0.0
author: wechat-skill
permissions: 本地微信窗口自动化（pywinauto）、剪贴板
---

> 说明：本技能基于原始仓库 [neo0591/weichat-skill](https://github.com/neo0591/weichat-skill) 进行二次开发与适配。

# 微信对接技能

## 1. Description

本技能为 OpenClaw 提供**微信消息发送和监听接入**能力。监听阶段会打开独立微信窗口，并把新消息通过 Webhook 推送到 OpenClaw。

核心能力：
- 向指定微信群发送文本消息
- 启动指定联系人/群的独立窗口监听
- 支持 @ 群成员
- 待回传队列管理（用于消息转发场景）

## 2. When to use

当 OpenClaw 需要在微信中发送消息时调用本技能：
- 收到微信消息后需要回复
- 需要将消息转发到其他群
- 需要管理消息转发回传关系
- 需要开始监听某个联系人或群，并让消息进入 OpenClaw 网关对话

## 3. How to use

当用户要求“开始监听某个联系人/群”时，优先执行：

```bash
python scripts/system_start_wechat_listener.py --target "目标名称"
```

未明确说明是群聊时，默认按联系人处理，也就是 `is_group=False`。

如果监听的是群，并且只在被提及时回复，执行：

```bash
python scripts/system_start_wechat_listener.py --target "群名称" --is-group --mention-token "@小小范"
```

如果用户明确要求自动回复，可执行：

```bash
python scripts/system_start_wechat_listener.py --target "目标名称" --auto-reply --reply-message "默认回复内容"
```

这个入口负责：
- 读取现有 `config.yaml`
- 生成只包含该目标的监听请求
- 触发 Windows 计划任务，在当前桌面会话里打开独立可见的 `cmd` 黑窗
- 由 `wechat_listener.py` 打开目标的独立微信窗口并开始监听
- 收到消息后通过 webhook 投递到 OpenClaw 网关对话
- 不同目标自动使用不同的 hook session
- 如果 OpenClaw 返回回复内容，则自动发回该独立窗口
- 持续监听，直到显式执行停止命令

当用户要求“停止监听微信”“停止监听某个联系人/群”时，执行：

```bash
python scripts/stop_wechat_listener.py
```

### 3.1 发送消息

```bash
python scripts/send_wechat_message.py --group "群名称" --message "消息内容" [--at "成员1,成员2"]
```

参数说明：
- `--group`：微信群名称（与微信显示一致）
- `--message`：要发送的文本
- `--at`：可选，要 @ 的群成员昵称，多个用逗号分隔

### 3.2 待回传管理

**记录待回传**（A群转发到B群时）：
```bash
python scripts/record_pending_relay.py --supplier-group "B群" --to-group "A群" --at "原发送者"
```

**查询待回传**（收到B群消息时）：
```bash
python scripts/get_pending_relay.py --supplier-group "B群"
# 输出: {"to_group": "A群", "at": "原发送者"}
# 退出码: 0=有待回传, 1=无待回传
```

**清除待回传**（回传完成后）：
```bash
python scripts/get_pending_relay.py --supplier-group "B群" --clear
```

### 3.3 依赖

- **pyweixin**：来自 pywechat 项目
- **PyYAML**：配置解析
- **PC 微信 4.1+**：已登录状态

### 3.4 启动独立窗口监听

```bash
python scripts/system_start_wechat_listener.py --target "文件传输助手"
python scripts/system_start_wechat_listener.py --target "项目群" --duration 24h
```

说明：
- 这个入口会基于现有 `config.yaml` 临时生成监听配置。
- 监听时会由 `wechat_listener.py` 打开该目标的独立微信窗口。
- 微信消息会通过 webhook 投递到 OpenClaw 网关对话。
### 3.5 监听端配置（配合使用）

监听由 pyweixin 的 `wechat_listener_webhook.py` 完成：

1. 在 pywechat 目录配置 `listener_config.yaml`：
```yaml
openclaw_webhook_url: "http://127.0.0.1:18789/hooks/agent"
openclaw_token: "your-token"
groups:
  - "群名称1"
```

2. OpenClaw 的 `openclaw.json` 配置：
```json
{
  "hooks": {
    "enabled": true,
    "token": "your-token",
    "path": "/hooks"
  }
}
```

3. 运行监听：
```bash
python scripts/wechat_listener_webhook.py --config listener_config.yaml --no-close
```

## 4. 典型工作流

### 消息转发回传

```
A群消息 ──webhook──> OpenClaw
                        │
                        ↓ 判断需转发到B群
                   record_pending_relay(B, A, 发送者)
                        │
                        ↓
                  send_wechat_message(B群, 消息)
                        │
B群回复 ──webhook──> OpenClaw
                        │
                        ↓
                  get_pending_relay(B群) → 得到(A群, 发送者)
                        │
                        ↓
                  send_wechat_message(A群, 回复, @发送者)
                        │
                        ↓
                  get_pending_relay(B群, --clear)
```

## 5. Implementation

- `scripts/send_wechat_message.py`：发送消息脚本
- `scripts/record_pending_relay.py`：记录待回传
- `scripts/get_pending_relay.py`：查询/清除待回传
- `wechat_pending_relay.json`：待回传状态文件（运行时生成）

## 6. Edge cases

- 群名、@ 的昵称需与微信显示一致
- 发送消息时微信需在前台或最小化，不能完全隐藏
- 待回传文件在技能根目录，多个实例需注意并发
