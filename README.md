# wechat-auto-skill

面向 OpenClaw 的微信自动化技能，支持在当前桌面会话中拉起独立监听黑框、接收微信消息、回传到 OpenClaw，并将 OpenClaw 回复发回微信。

## 项目来源（基于原仓库开发）

本仓库为二次开发版本，明确基于原始仓库：

- 原始仓库：[neo0591/weichat-skill](https://github.com/neo0591/weichat-skill)

在原仓库能力上，本仓库主要补充了 OpenClaw 场景所需的系统级启动与会话路由能力。

## 这个 skill 的实现方式

OpenClaw 不直接长期持有 Python 监听进程，而是走“系统级启动器”链路：

1. OpenClaw 调用 `scripts/system_start_wechat_listener.py`
2. 写入单目标请求文件：`runtime/requests/<target>.json`
3. 创建并触发 Windows 计划任务：`OpenClaw WeChat Listener - <target-slug>`
4. 计划任务执行 `scripts/listener_task_entry.cmd`，弹出可见 `cmd` 黑框
5. `listener_task_entry.py` 读取请求，生成临时 YAML，仅保留当前目标
6. 启动 `scripts/wechat_listener.py -u`，实时输出日志到黑框并写入日志文件
7. `wechat_listener.py` 打开该目标独立微信窗口，监听消息并回调 OpenClaw webhook
8. 等待 OpenClaw 会话文件中的回复文本，拿到后自动发送回微信

## 调用细节（OpenClaw 应该怎么调）

建议 OpenClaw 统一只调用 `system_start_wechat_listener.py`，不要直接调用 `wechat_listener.py`。

### 1) 启动监听（联系人，默认）

```bash
python scripts/system_start_wechat_listener.py --target "Nirvana"
```

说明：
- 默认按联系人处理：`is_group=False`
- 每个目标自动绑定独立 session key：`hook:wechat:<target-slug>`

### 2) 启动监听（群聊）

```bash
python scripts/system_start_wechat_listener.py --target "新的周末就要来啦" --is-group --mention-token "@小小范"
```

说明：
- 群聊模式下，默认只处理被提及消息
- 提及判断兼容 `@小小范`、`@ 小小范`、`＠小小范`
- 未命中提及词会被 SKIP，不会自动回复

### 3) 本地兜底自动回复（可选）

```bash
python scripts/system_start_wechat_listener.py --target "Nirvana" --auto-reply --reply-message "收到，稍后处理。"
```

### 4) 停止监听

```bash
python scripts/stop_wechat_listener.py
```

行为：
- 写入停止信号文件
- 停止并禁用已创建的监听计划任务
- 尝试终止监听进程

## 参数说明（启动入口）

`system_start_wechat_listener.py` 关键参数：

- `--target`：必填，联系人或群名称（与微信显示一致）
- `--is-group`：可选，标记为群聊；不传即 `False`
- `--duration`：监听时长，默认 `forever`
- `--mention-token`：群聊提及词，默认 `@小小范`
- `--session-key`：可选，显式指定会话键；不传时自动生成
- `--auto-reply` / `--reply-to-user` / `--reply-message`：本地兜底回复参数
- `--config`：基础配置文件路径，默认 `config.yaml`

## 回复链路与防回环

- 监听消息后，发送到 OpenClaw `hooks/agent`
- 使用目标独立 `session_key`，不同联系人/群互不串会话
- 从 OpenClaw session jsonl 等待本轮回复文本
- 成功后回发微信
- 内置“最近已发送消息缓存”，避免把自己刚发出的消息再次当作输入造成死循环

## 日志与排障

- 黑框实时日志：由 `listener_task_entry.cmd` + `python -u` 输出
- 聚合日志：
  - `%USERPROFILE%\\.openclaw\\logs\\wechat-listener.log`
  - `%USERPROFILE%\\.openclaw\\logs\\wechat-listener-task.log`
- 请求文件：`runtime/requests/*.json`

## ⚠️ 免责声明

**请勿将 pywechat 用于任何非法商业活动，因此造成的一切后果由使用者自行承担！**

本项目仅供学习和研究使用，使用者应遵守相关法律法规和微信用户协议。
