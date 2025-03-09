import os
import json

# 默认配置
DEFAULT_CONFIG = {
    'room_email': 'room@example.com',                   # 会议室邮箱
    'telegram_api_id': 0,                               # Telegram API ID
    'telegram_api_hash': '',                            # Telegram API Hash
    'telegram_chat_id': -1000000000000,                 # Telegram群组ID
    'check_interval_minutes': 5,                        # 检查间隔（分钟）
    'db_file': 'room_bookings.db',                      # SQLite数据库文件
    'session_name': 'room_notifier',                    # Telethon会话名称
}

# 配置文件路径
CONFIG_FILE = 'config.json'

def load_config():
    """加载配置，先加载配置文件，最后是默认值"""
    config = DEFAULT_CONFIG.copy()
    
    # 尝试从配置文件加载
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                file_config = json.load(f)
                config.update(file_config)
    except Exception as e:
        print(f"警告: 读取配置文件失败: {str(e)}")
    
    return config

def save_config_template():
    """保存配置模板，用于用户第一次使用"""
    # 不覆盖现有配置文件
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)
            print(f"已创建配置模板: {CONFIG_FILE}")
    

def validate_config(config):
    """验证配置是否有效"""
    required_keys = ['room_email', 'telegram_api_id', 'telegram_api_hash', 'telegram_chat_id']
    missing_keys = [key for key in required_keys if not config.get(key)]
    
    if missing_keys:
        print("错误: 缺少必要的配置项:")
        for key in missing_keys:
            print(f"  - {key}")
        return False
    
    # 验证Telegram API ID
    if not isinstance(config['telegram_api_id'], int) or config['telegram_api_id'] <= 0:
        print("错误: telegram_api_id 必须是正整数")
        return False
    
    # 验证API Hash
    if len(config['telegram_api_hash']) < 8:
        print("警告: telegram_api_hash 可能无效")
    
    # 验证群组ID
    if not isinstance(config['telegram_chat_id'], int):
        print("错误: telegram_chat_id 必须是整数")
        return False
    
    return True

# 示例使用
if __name__ == "__main__":
    save_config_template()
    config = load_config()
    if validate_config(config):
        print("配置有效，可以使用")
    else:
        print("配置无效，请修正上述错误")