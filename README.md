# 会议室预订通知系统

这是一个自动监控会议室预订并通过Telegram发送通知的工具。系统会定期检查Outlook日历中的会议室预订，并将即将到来的会议信息发送到指定的Telegram群组。

## 特性

- 自动监控Outlook会议室日历
- 使用Telethon库发送Telegram通知
- 支持自定义检查时间间隔
- 防止重复通知
- 优雅的异常处理和日志记录
- 使用配置文件进行安全配置

## 安装需求

- Python 3.7+
- Windows系统 (由于使用win32com访问Outlook)
- Outlook客户端已安装并配置
- Telegram账号 (用于获取API凭据)

## 安装

1. 克隆或下载此仓库
2. 安装依赖包:

```bash
pip install -r requirements.txt
```

## 配置

你需要修改`config.json`文件来配置系统。首次运行会创建模板文件。

### 必需的配置项:

```
room_email 会议室邮箱地址
telegram_api_id 你的Telegram API ID
telegram_api_hash 你的Telegram API Hash
telegram_chat_id 目标群组ID (通常为负数)
```

### 获取Telegram API凭据:

1. 访问 https://my.telegram.org/auth
2. 登录您的Telegram账号
3. 点击 "API development tools"
4. 创建一个新的应用，记下API ID和API Hash

### 获取Telegram群组ID:

使用脚本`get_telegram_group_id.py`来列出您的所有对话及其ID


## 使用方法

```bash
python room_notifier.py
```

该服务将持续运行，根据配置的间隔时间监控会议室预订状况并发送通知。

## 工作原理

1. 连接到本地Outlook客户端
2. 访问指定的会议室邮箱日历
3. 检索指定时间范围内的约会
4. 对每个尚未处理的约会，格式化并通过Telegram发送通知
5. 记录处理过的事件以避免重复通知

## 许可

本项目采用MIT许可 - 详情请参见[LICENSE](LICENSE)文件。

## 免责声明

该工具仅供学习和参考目的。请确保您有适当的权限访问相关邮箱和发送消息到相关群组。不要用于未经许可的会议室监控。