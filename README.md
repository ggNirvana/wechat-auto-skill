# wechat-skill

**OpenClaw 微信对接技能**：为 OpenClaw 提供微信群消息收发能力。

## 项目来源说明

本仓库**基于原始项目**开发：

- 原始仓库：[neo0591/weichat-skill](https://github.com/neo0591/weichat-skill)

在原项目基础上，本仓库增加并整理了 OpenClaw 场景下的二次开发能力（系统级启动器、多目标监听任务、Webhook 会话路由、监听稳定性修复等）。

## 核心功能

1. **发送消息**：向指定微信群/好友发送消息，支持 @ 群成员
2. **待回传队列**：记录消息转发关系，实现 A 群 → B 群 → A 群的消息回传
3. **消息监听**：配合 pyweixin 监听微信群消息并推送到 OpenClaw

## 架构说明

```
┌─────────────────┐     Webhook      ┌─────────────────┐
│   pyweixin      │ ───────────────> │    OpenClaw     │
│  (微信监听)      │                 │   (智能体)       │
└─────────────────┘                  └─────────────────┘
        ↑                                    │
        │                                    ↓
        │                            ┌─────────────────┐
        └────────────────────────────│  wechat-skill   │
              发送消息                │  (本技能)        │
                                       └─────────────────┘
```

- **监听**：由 pyweixin 的监听模块完成，新消息通过 Webhook 推送到 OpenClaw
- **发送**：由本技能的 `send_wechat_message.py` 完成，OpenClaw 需要发消息时调用此脚本

---

## ⚠️ 免责声明

**请勿将 pywechat 用于任何非法商业活动，因此造成的一切后果由使用者自行承担！**

本项目仅供学习和研究使用，使用者应遵守相关法律法规和微信用户协议。

---

## 安装

### 环境要求

- Windows 10/11
- Python 3.9+
- PC 微信 4.0+（推荐 4.1+）

### 1. 安装 pyweixin

```bash
# 方式一：指定路径安装
pip install -e /path/to/pywechat

# 方式二：在 pywechat 目录下安装
cd /path/to/pywechat
pip install -e .
```

### 2. 安装其他依赖

```bash
pip install -r requirements.txt
```

### 3. 前置条件

- **PC 微信已登录**：确保微信在前台或最小化状态
- **讲述人已启动**（监听功能需要）：设置 → 轻松使用 → 讲述人

---

## 快速开始

### 一键安装

```bash
# 运行安装脚本
install.bat
```

安装脚本会自动：
1. 检查 Python 环境
2. 安装依赖包
3. 引导安装 pyweixin
4. 创建配置文件

### 启动服务

```bash
# 方式一：交互式菜单（推荐）
run.bat

# 方式二：命令行模式
run.bat env           # 准备环境（启动讲述人+微信）
run.bat listen        # 开始监听
run.bat listen --duration 1h  # 监听1小时
run.bat stop          # 停止监听
run.bat send          # 发送消息
run.bat status        # 查看状态
```

### 环境准备功能

`run.bat` 提供一键环境准备：

1. **检查 Python** - 验证 Python 环境
2. **检查 pyweixin** - 验证依赖安装
3. **启动微信** - 自动启动 PC 微信
4. **启动讲述人** - 启动 Windows 讲述人（监听功能需要）
5. **定时关闭讲述人** - 5分钟后自动关闭讲述人

### 配置文件

复制 `config.example.yaml` 为 `config.yaml` 并修改：

```yaml
# 监听时长
duration: "24h"

# OpenClaw Webhook 配置
openclaw_webhook_url: "http://127.0.0.1:18789/hooks/agent"
openclaw_token: "your-token-here"

# 监听的群列表
groups:
  - name: "群名称1"
    auto_reply: false              # 是否自动回复
  - name: "群名称2"
    auto_reply: true
    reply_to_user: "你的微信名"    # 触发回复的用户名
```

### 配置 OpenClaw Webhook

在 `~/.openclaw/openclaw.json` 中：

```json
{
  "gateway": {
    "port": 18789
  },
  "hooks": {
    "enabled": true,
    "token": "your-secure-token",
    "path": "/hooks"
  }
}
```

---

## 使用方法

### 命令行参数

```bash
# 环境准备
run.bat env                    # 检查环境、启动微信和讲述人

# 监听服务
run.bat listen                 # 开始监听（交互式输入时长）
run.bat listen --duration 1h   # 监听1小时
run.bat stop                   # 停止监听服务

# 发送消息
run.bat send                   # 交互式发送消息

# 待回传管理
run.bat relay                  # 进入待回传管理菜单

# 状态检查
run.bat status                 # 查看系统状态

# 安装依赖
run.bat install                # 安装Python依赖
```

### 监听服务

监听指定群的消息并推送到 OpenClaw：

```bash
# 使用配置文件监听
python scripts/wechat_listener.py --config config.yaml

# 指定时长监听
python scripts/wechat_listener.py --config config.yaml --duration 1h

# 不自动关闭微信
python scripts/wechat_listener.py --config config.yaml --no-close

# 覆盖 Webhook 配置
python scripts/wechat_listener.py --config config.yaml \
    --webhook-url "http://127.0.0.1:18789/hooks/agent" \
    --token "your-token"
```

监听服务会将消息推送到 OpenClaw Webhook，格式为：

```json
{
  "message": "【群】群名称【发件人】发送者【内容】消息正文",
  "name": "WeChatListener",
  "wakeMode": "now",
  "metadata": {
    "group": "群名称",
    "sender": "发送者",
    "timestamp": "2024-01-01T12:00:00"
  }
}
```

### 发送消息

```bash
# 发送到群
python scripts/send_wechat_message.py --group "群名称" --message "消息内容"

# 发送并 @ 成员
python scripts/send_wechat_message.py --group "群名称" --message "消息内容" --at "成员1,成员2"

# 发送到好友
python scripts/send_wechat_message.py --group "好友昵称" --message "消息内容"
```

### 待回传管理

用于实现消息转发回传场景：

```bash
# 记录：A群消息转发到B群时
python scripts/record_pending_relay.py --supplier-group "B群" --to-group "A群" --at "发送者"

# 查询：收到B群回复时
python scripts/get_pending_relay.py --supplier-group "B群"
# 输出: {"to_group": "A群", "at": "发送者"}

# 清除：回传完成后
python scripts/get_pending_relay.py --supplier-group "B群" --clear
```

---

## pyweixin 功能参考

本项目基于 [pywechat](https://github.com/Hello-Mr-Crab/pywechat) 项目，以下是常用功能：

### Navigator - 窗口导航

```python
from pyweixin import Navigator

# 打开聊天窗口
window = Navigator.open_seperate_dialog_window(
    friend="群名称",        # 群名或好友昵称
    window_minimize=False,  # 是否最小化
    close_weixin=False      # 结束后是否关闭微信
)

# 搜索联系人
Navigator.search_friend(name="好友昵称")

# 搜索公众号
Navigator.search_official_account(name="公众号名称")

# 搜索小程序
Navigator.search_miniprogram(name="小程序名称")

# 搜索视频号
Navigator.search_channels(search_content="视频号名称")
```

### Messages - 消息发送

```python
from pyweixin import Messages

# 发送消息
Messages.send_messages_to_friend(
    friend="群名称",
    messages=["消息1", "消息2"],
    at_members=["成员1", "成员2"],  # 可选，@ 成员
    close_weixin=False
)
```

### AutoReply - 自动回复监听

```python
from pyweixin import Navigator, GlobalConfig
from pyweixin.WeChatAuto import AutoReply

GlobalConfig.close_weixin = False

def on_message(text: str, sender: str, receiver: str):
    print(f"[{sender}] {text}")
    # 返回字符串表示自动回复，返回 None 表示不回复
    if "hello" in text.lower():
        return "你好！"
    return None

window = Navigator.open_seperate_dialog_window(friend="群名称")
AutoReply.auto_reply_to_friend(
    window,
    duration="24h",           # 监听时长
    callback=on_message,      # 回调函数
    close_dialog_window=False
)
```

### Monitor - 消息监听

```python
from pyweixin import Navigator, Monitor

window = Navigator.open_seperate_dialog_window(friend="群名称")
result = Monitor.listen_on_chat(
    dialog_window=window,
    duration="30s"
)
print(result)
# 返回: {'新消息总数': x, '文本数量': x, '文件数量': x, ...}
```

### 多群并发监听

```python
from concurrent.futures import ThreadPoolExecutor
from pyweixin import Navigator, GlobalConfig
from pyweixin.WeChatAuto import AutoReply

GlobalConfig.close_weixin = False

friends = ["群1", "群2"]
dialog_windows = []

# 打开所有窗口
for friend in friends:
    window = Navigator.open_seperate_dialog_window(
        friend=friend,
        window_minimize=True,
        close_weixin=True
    )
    dialog_windows.append(window)

# 定义回调
def reply_func1(text, sender, receiver):
    return "群1回复"

def reply_func2(text, sender, receiver):
    return "群2回复"

callbacks = [reply_func1, reply_func2]
durations = ["1min", "1min"]

# 并发监听
with ThreadPoolExecutor() as pool:
    results = pool.map(
        lambda args: AutoReply.auto_reply_to_friend(*args),
        list(zip(dialog_windows, durations, callbacks))
    )

for friend, result in zip(friends, results):
    print(friend, result)
```

### Moments - 朋友圈

```python
from pyweixin import Moments

# 获取朋友圈数据
posts = Moments.dump_recent_posts(recent='Today')
for post in posts:
    print(post)

# 发布朋友圈
Moments.post_moments(
    texts="发布内容",
    medias=["/path/to/image1.png", "/path/to/image2.png"]
)

# 导出好友朋友圈
Moments.dump_friend_posts(
    friend="好友昵称",
    number=10,
    save_detail=True,
    target_folder="/path/to/save"
)

# 点赞并评论好友朋友圈
def comment_func(content):
    if "关键词" in content:
        return "评论内容"
    return None

Moments.like_friend_posts(
    friend="好友昵称",
    number=20,
    callback=comment_func,
    use_green_send=True
)
```

### Collections - 收藏/公众号

```python
from pyweixin import Collections

# 获取公众号文章
Collections.collect_offAcc_articles(name="新华社", number=10)
urls = Collections.cardLink_to_url(number=10)
for url, text in urls.items():
    print(f"{text}\n{url}")
```

### GlobalConfig - 全局配置

```python
from pyweixin import GlobalConfig

# 设置全局参数
GlobalConfig.load_delay = 2.5      # 加载延迟
GlobalConfig.is_maximize = True    # 是否最大化
GlobalConfig.close_weixin = False  # 结束后不关闭微信
```

---

## 典型场景

### 场景：消息转发与回传

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

### 完整流程脚本

```bash
# 1. A群收到消息，OpenClaw判断需转发到B群
# 2. 记录待回传
python scripts/record_pending_relay.py \
    --supplier-group "B群" \
    --to-group "A群" \
    --at "原发送者"

# 3. 发送到B群
python scripts/send_wechat_message.py \
    --group "B群" \
    --message "转发内容"

# 4. B群回复后，查询待回传
python scripts/get_pending_relay.py --supplier-group "B群"
# 输出: {"to_group": "A群", "at": "原发送者"}

# 5. 回传到A群
python scripts/send_wechat_message.py \
    --group "A群" \
    --message "回复内容" \
    --at "原发送者"

# 6. 清除待回传
python scripts/get_pending_relay.py --supplier-group "B群" --clear
```

---

## 文件说明

```
wechat-skill/
├── README.md                    # 本文档
├── SKILL.md                     # OpenClaw 技能描述
├── USAGE.md                     # 详细使用指南
├── LICENSE                      # 许可证
├── config.example.yaml          # 配置示例
├── requirements.txt             # Python 依赖
├── install.bat                  # 安装脚本
├── run.bat                      # 交互式启动脚本
├── img/                         # 图片资源
│   └── *.jpg                    # 赞助图片
├── lib/                         # 依赖库目录
│   └── README.md                # pywechat 放置说明
└── scripts/
    ├── wechat_listener.py       # 微信监听服务
    ├── send_wechat_message.py   # 发送微信消息
    ├── record_pending_relay.py  # 记录待回传
    └── get_pending_relay.py     # 查询/清除待回传
```

---

## 常见问题

### 1. 微信窗口无法操作

**原因**：微信未正确初始化或讲述人未启动。

**解决**：
- 确保微信已登录并正常运行
- 启动 Windows 讲述人：设置 → 轻松使用 → 讲述人
- 检查 pyweixin 版本是否匹配微信版本

### 2. @ 功能不生效

**原因**：昵称不匹配或成员不在群内。

**解决**：
- 确保昵称与群内显示一致（包括表情符号）
- @ 功能需要成员在群内

### 3. 监听不到消息

**原因**：讲述人未启动或窗口未正确打开。

**解决**：
- 启动 Windows 讲述人
- 确保监听脚本能打开群聊独立窗口
- 检查群名称配置是否正确

### 4. Webhook 连接失败

**原因**：OpenClaw 未启动或配置错误。

**解决**：
- 确认 OpenClaw 已启动并监听配置端口
- 检查 `openclaw.json` 中的 hooks 配置
- 确认 token 在监听端和 OpenClaw 端一致
- 测试：`curl http://127.0.0.1:18789/hooks/agent`

---

## 更多文档

- [USAGE.md](USAGE.md) - 详细使用说明，包含 pyweixin 配置和 OpenClaw Hook 配置
- [pywechat 项目](https://github.com/Hello-Mr-Crab/pywechat) - pyweixin 源码及文档

## 许可证

GNU LESSER GENERAL PUBLIC LICENSE

---

## ❤️ 支持作者

若对您有帮助，可以赞助并支持下作者哦，谢谢！

![赞助](img/d34489d98d3633333ed872b2d2087289.jpg)
