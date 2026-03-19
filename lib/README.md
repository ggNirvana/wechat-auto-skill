# lib 目录说明

此目录用于存放 pywechat 项目源码，实现开箱即用。

## 使用方法

### 方式一：复制 pywechat 源码（推荐）

1. 将 pywechat 项目复制到此目录：

```
wechat-skill/
├── lib/
│   └── pywechat/           <-- 复制整个 pywechat 项目到这里
│       ├── pyweixin/
│       ├── setup.py
│       └── ...
├── scripts/
├── config.yaml
└── ...
```

2. 运行安装：

```bash
pip install -e ./lib/pywechat
```

或使用安装脚本：

```bash
install.bat
```

### 方式二：指定路径安装

如果你已有 pywechat 项目在其他位置：

```bash
pip install -e D:\WorkSpace\python\pywechat
```

## pywechat 项目地址

- GitHub: https://github.com/Hello-Mr-Crab/pywechat

## 验证安装

```bash
python -c "from pyweixin import Navigator; print('OK')"
```

输出 `OK` 表示安装成功。
