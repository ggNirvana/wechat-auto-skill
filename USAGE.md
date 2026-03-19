# 微信对接技能使用指南

本文档详细介绍 pyweixin 库的使用方法和 OpenClaw Hook 配置说明。

## 目录

1. [环境准备](#1-环境准备)
2. [pyweixin 安装与配置](#2-pyweixin-安装与配置)
3. [OpenClaw Hook 配置](#3-openclaw-hook-配置)
4. [监听服务部署](#4-监听服务部署)
5. [发送消息脚本详解](#5-发送消息脚本详解)
6. [常见问题](#6-常见问题)

---

## 1. 环境准备

### 系统要求

- Windows 10/11
- Python 3.8+
- PC 微信 4.1+

### 安装依赖

```bash
pip install pyyaml
pip install pyautogui
pip install pywinauto
pip install emoji
```

### 安装 pyweixin

从 pywechat 项目安装：

```bash
# 方式一：指定路径安装
pip install -e D:\WorkSpace\python\pywechat

# 方式二：在 pywechat 目录下安装
cd D:\WorkSpace\python\pywechat
pip install -e .
```

---

## 2. pyweixin 安装与配置

### 2.1 pyweixin 核心模块

pyweixin 提供以下核心功能：

```python
from pyweixin import Navigator, Messages, GlobalConfig
from pyweixin.WeChatAuto import AutoReply
from pyweixin.Uielements import Edits, Windows
from pyweixin.WinSettings import SystemSettings
```

| 模块 | 功能 |
|------|------|
| `Navigator` | 窗口导航，打开聊天窗口 |
| `Messages` | 消息发送 |
| `AutoReply` | 自动回复监听 |
| `GlobalConfig` | 全局配置 |
| `SystemSettings` | 系统设置（剪贴板等） |

### 2.2 发送消息

```python
from pyweixin import Messages

# 发送消息到群
Messages.send_messages_to_friend(
    friend="群名称",
    messages=["消息内容"],
    at_members=["成员1", "成员2"],  # 可选
    close_weixin=False
)
```

### 2.3 打开独立聊天窗口

```python
from pyweixin import Navigator

# 打开群聊独立窗口
window = Navigator.open_seperate_dialog_window(
    friend="群名称",
    window_minimize=False,
    close_weixin=False
)
```

### 2.4 监听消息

```python
from pyweixin import Navigator, GlobalConfig
from pyweixin.WeChatAuto import AutoReply

GlobalConfig.close_weixin = False

def on_message(text: str, sender: str, receiver: str):
    print(f"收到消息: {sender} -> {text}")
    # 返回 None 表示不自动回复
    # 返回字符串表示自动回复该内容
    return None

window = Navigator.open_seperate_dialog_window(friend="群名称")
AutoReply.auto_reply_to_friend(window, "24h", on_message, close_dialog_window=False)
```

---

## 3. OpenClaw Hook 配置

### 3.1 OpenClaw 配置文件

在 `~/.openclaw/openclaw.json` 中配置：

```json
{
  "gateway": {
    "port": 18789
  },
  "hooks": {
    "enabled": true,
    "token": "your-secure-token-here",
    "path": "/hooks"
  }
}
```

配置说明：

| 字段 | 说明 |
|------|------|
| `gateway.port` | OpenClaw 网关端口，默认 18789 |
| `hooks.enabled` | 是否启用 Hook 接收 |
| `hooks.token` | 验证令牌，需与监听端配置一致 |
| `hooks.path` | Hook 路径前缀，默认 `/hooks` |

### 3.2 监听端配置

在 pywechat 项目目录创建 `listener_config.yaml`：

```yaml
# OpenClaw Webhook 配置
openclaw_webhook_url: "http://127.0.0.1:18789/hooks/agent"
openclaw_token: "your-secure-token-here"

# 监听的群列表
groups:
  - "群名称1"
  - "群名称2"
  - "群名称3"

# 监听配置
duration: "8760h"      # 监听时长，8760h = 1年
close_weixin: false    # 结束后不关闭微信
```

### 3.3 Webhook 消息格式

监听脚本会将微信消息 POST 到 OpenClaw：

```json
{
  "message": "【群】群名称【发件人】发送者昵称【内容】消息正文",
  "name": "WeChatListener",
  "wakeMode": "now"
}
```

OpenClaw 可以解析该格式获取群名、发送者和消息内容。

---

## 4. 监听服务部署

### 4.1 启动监听

```bash
cd D:\WorkSpace\python\pywechat
python scripts/wechat_listener_webhook.py --config listener_config.yaml --no-close
```

参数说明：

| 参数 | 说明 |
|------|------|
| `--config` | 配置文件路径 |
| `--no-close` | 结束后不关闭微信 |
| `--duration` | 覆盖配置中的监听时长 |

### 4.2 后台运行（Windows）

使用 VBS 脚本后台运行：

```vbs
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d D:\WorkSpace\python\pywechat && python scripts/wechat_listener_webhook.py --config listener_config.yaml --no-close", 0, False
```

保存为 `start_listener.vbs`，双击运行即可后台启动。

### 4.3 开机自启

将 VBS 脚本放入 Windows 启动目录：

```
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
```

---

## 5. 发送消息脚本详解

### 5.1 send_wechat_message.py

```bash
python scripts/send_wechat_message.py \
    --group "群名称" \
    --message "消息内容" \
    --at "成员1,成员2" \
    --close-weixin
```

| 参数 | 必填 | 说明 |
|------|------|------|
| `--group` | 是 | 群名称，与微信显示一致 |
| `--message` | 是 | 消息内容 |
| `--at` | 否 | 要 @ 的成员，逗号分隔 |
| `--close-weixin` | 否 | 发送后关闭微信窗口 |

### 5.2 record_pending_relay.py

记录消息转发关系，用于后续回传：

```bash
python scripts/record_pending_relay.py \
    --supplier-group "供应商群" \
    --to-group "客户群" \
    --at "原发送者"
```

状态保存在 `wechat_pending_relay.json`：

```json
{
  "供应商群": {
    "to_group": "客户群",
    "at": "原发送者"
  }
}
```

### 5.3 get_pending_relay.py

查询和清除待回传：

```bash
# 查询
python scripts/get_pending_relay.py --supplier-group "供应商群"

# 清除
python scripts/get_pending_relay.py --supplier-group "供应商群" --clear

# 指定输出格式
python scripts/get_pending_relay.py --supplier-group "供应商群" --format kv
```

退出码：
- 0：有待回传记录
- 1：无待回传记录

---

## 6. 常见问题

### 6.1 微信窗口无法操作

**原因**：微信未正确初始化或讲述人未启动。

**解决**：
1. 确保微信已登录并正常运行
2. 启动 Windows 讲述人（监听功能需要）
3. 检查 pyweixin 版本是否匹配微信版本

### 6.2 @ 功能不生效

**原因**：昵称不匹配或成员不在群内。

**解决**：
1. 确保昵称与群内显示一致（包括表情符号）
2. @ 功能需要成员在群内
3. 检查是否有特殊字符

### 6.3 监听不到消息

**原因**：讲述人未启动或窗口未正确打开。

**解决**：
1. 启动 Windows 讲述人：设置 → 轻松使用 → 讲述人
2. 确保监听脚本能打开群聊独立窗口
3. 检查群名称配置是否正确

### 6.4 Webhook 连接失败

**原因**：OpenClaw 未启动或配置错误。

**解决**：
1. 确认 OpenClaw 已启动并监听配置端口
2. 检查 `openclaw.json` 中的 hooks 配置
3. 确认 token 在监听端和 OpenClaw 端一致
4. 测试 Webhook：`curl http://127.0.0.1:18789/hooks/agent`

### 6.5 消息发送失败

**原因**：微信窗口状态异常。

**解决**：
1. 确保微信未被最小化到托盘
2. 尝试手动激活微信窗口后再发送
3. 检查剪贴板功能是否正常

---

## 附录：pyweixin 参考

### 主要类和函数

```python
# 导航类
Navigator.open_seperate_dialog_window(friend, window_minimize, close_weixin)
    # 打开独立聊天窗口

# 消息类
Messages.send_messages_to_friend(friend, messages, at_members, close_weixin)
    # 发送消息到好友/群

# 自动回复类
AutoReply.auto_reply_to_friend(window, duration, callback, close_dialog_window)
    # 监听并自动回复
    # callback: Callable[[text, sender, receiver], str|None]

# 全局配置
GlobalConfig.close_weixin = False  # 是否在结束时关闭微信
```

### UI 元素

```python
from pyweixin.Uielements import Edits, Windows

Edits.CurrentChatEdit   # 当前聊天输入框
Windows.MentionPopOverWindow  # @ 弹窗
```

更多 API 请参考 pywechat 项目文档。
