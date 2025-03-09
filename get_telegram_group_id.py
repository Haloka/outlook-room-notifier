from telethon import TelegramClient, events, sync
from config_manager import load_config, save_config_template, validate_config

save_config_template()  # 确保配置模板存在
CONFIG = load_config()

# These example values won't work. You must get your own api_id and
# api_hash from https://my.telegram.org, under API Development.
api_id = CONFIG['telegram_api_id']
api_hash = CONFIG['telegram_api_hash']

client = TelegramClient(CONFIG['session_name'], api_id, api_hash)
client.start()

print(client.get_me().stringify())

dialogs = client.get_dialogs()
print("=== 您的Telegram对话列表 ===")
print("{:<4} | {:<30} | {:<15} | {:<10}".format("序号", "名称", "类型", "ID"))
print("-" * 70)

# 枚举并打印所有对话
for i, dialog in enumerate(dialogs):
    entity_type = "用户" if dialog.is_user else ("群组" if dialog.is_group else "频道")
    print("{:<4} | {:<30} | {:<15} | {:<10}".format(
        i, 
        dialog.name[:30], 
        entity_type, 
        dialog.id
    ))
