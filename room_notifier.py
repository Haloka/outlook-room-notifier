import win32com.client
import datetime
import time
import schedule
import requests
import os
import hashlib
import sqlite3
from telethon import TelegramClient, events, sync
from config_manager import load_config, save_config_template, validate_config

# 加载配置
save_config_template()  # 确保配置模板存在
CONFIG = load_config()

if not validate_config(CONFIG):
    print("配置无效，请检查您的配置文件或环境变量")
    exit(1)

client = TelegramClient(CONFIG['session_name'], CONFIG['telegram_api_id'], CONFIG['telegram_api_hash'])

# 跟踪已通知的事件
processed_events = set()

def check_room_bookings():
    """通过本地Outlook客户端检查会议室预订并发送通知"""
    try:
        # 连接Outlook应用程序
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # 尝试获取会议室日历
        recipient = namespace.CreateRecipient(CONFIG['room_email'])
        
        # 如果无法解析会议室邮箱，则给出警告
        if not recipient.Resolved:
            print(f"警告: 无法解析会议室邮箱 {CONFIG['room_email']}，尝试使用GAL查找")
            # 尝试在GAL中查找
            recipient.Resolve()
            if not recipient.Resolved:
                print("仍然无法解析，将使用默认日历")
                calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
            else:
                print("已在GAL中找到会议室")
                calendar = namespace.GetSharedDefaultFolder(recipient, 9)  # 9 = olFolderCalendar
        else:
            # 获取共享日历
            try:
                calendar = namespace.GetSharedDefaultFolder(recipient, 9)  # 9 = olFolderCalendar
                print(f"成功获取会议室 {CONFIG['room_email']} 的共享日历")
            except Exception as e:
                print(f"无法获取共享日历: {str(e)}")
                print("将使用默认日历并过滤会议室相关事件")
                calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
        
        # 计算时间范围
        # now = datetime.datetime.now()
        # end_time = now + datetime.timedelta(hours=1)  # 查看未来1小时的预订
        now = datetime.datetime.now() + datetime.timedelta(hours=-1)
        end_time = now + datetime.timedelta(hours=24)  # 查看未来1小时的预订

        # 格式化为Outlook所需的时间字符串格式
        start_str = now.strftime("%m/%d/%Y %H:%M %p")
        end_str = end_time.strftime("%m/%d/%Y %H:%M %p")
        
        # 获取指定时间范围内的约会
        restriction = f"[Start] >= '{start_str}' AND [End] <= '{end_str}'"
        appointments = calendar.Items
        appointments.Sort("[Start]")
        appointments = appointments.Restrict(restriction)
        
        # 记录是否有新通知发送
        notifications_sent = 0

        # 遍历约会
        for appointment in appointments:
            # 生成唯一ID
            event_id = generate_event_id(appointment)

            # 如果使用个人日历需要过滤
            if calendar == namespace.GetDefaultFolder(9):
                # 检查地点是否包含会议室名称
                is_room_event = False
                room_name = CONFIG['room_email'].split('@')[0].lower()
                
                if appointment.Location and room_name in appointment.Location.lower():
                    is_room_event = True
                
                # 检查与会者是否包含会议室
                if not is_room_event:
                    for recipient in appointment.Recipients:
                        if CONFIG['room_email'].lower() in str(recipient).lower():
                            is_room_event = True
                            break
                
                # 如果与会议室无关，跳过
                if not is_room_event:
                    continue
            
            # 检查是否已处理过此事件
            if not is_event_processed(event_id):
                # 获取事件详情
                organizer = appointment.Organizer
                subject = appointment.Subject if appointment.Subject else "无主题"
                
                # 格式化时间
                start_time = appointment.Start
                end_time = appointment.End
                
                start_date = start_time.strftime('%Y-%m-%d')
                start_time_str = start_time.strftime('%H:%M')
                end_time_str = end_time.strftime('%H:%M')
                
                # 创建消息
                message = (
                    f"🔔 会议室预订通知（自动任务）\n\n"
                    f"📅 日期: {start_date}\n"
                    f"🕒 时间: {start_time_str} - {end_time_str}\n"
                    f"👤 预订人: {organizer}\n"
                    f"📝 主题: {subject}"
                )
                
                # 如果有位置信息，添加到消息中
                if appointment.Location:
                    message += f"\n📍 地点: {appointment.Location}"
                
                # 发送Telegram通知
                if send_telegram_message(message):
                    # 标记事件为已处理
                    if mark_event_processed(event_id, appointment):
                        log_message("INFO", f"已发送通知: {subject}")
                        notifications_sent += 1

        if notifications_sent > 0:
            log_message("INFO", f"本次检查共发送 {notifications_sent} 条新通知")
        else:
            log_message("INFO", "本次检查未发现新预订")

    except Exception as e:
        log_message("ERROR", f"检查会议室预订失败，配置是否正确？: {str(e)}")

def send_telegram_message(message):
    """使用Telethon发送Telegram消息"""
    try:
        # 确保客户端已连接
        if not client.is_connected():
            client.connect()
        
        # 发送消息到指定的群组
        # client.send_message('我一个人的群', 'message')
        client.send_message(CONFIG['telegram_chat_id'], message)
        return True
    except Exception as e:
        log_message("ERROR", f"发送Telegram消息失败: {str(e)}")
        # 尝试重新连接
        try:
            if client.is_connected():
                client.disconnect()
            client.connect()
            # 重试发送
            client.send_message(CONFIG['telegram_chat_id'], message, parse_mode='md')
            return True
        except Exception as retry_error:
            log_message("ERROR", f"重试发送Telegram消息失败: {str(retry_error)}")
            return False

def init_database():
    """初始化SQLite数据库"""
    try:
        conn = sqlite3.connect(CONFIG['db_file'])
        cursor = conn.cursor()
        
        # 创建已处理事件表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS processed_events (
            event_id TEXT PRIMARY KEY,
            subject TEXT,
            organizer TEXT,
            start_time TEXT,
            end_time TEXT,
            event_date TEXT,
            location TEXT,
            processed_time TIMESTAMP,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        # 创建日志表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            log_type TEXT,
            message TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        conn.commit()
        conn.close()
        log_message("INFO", "数据库初始化成功")
    except Exception as e:
        print(f"数据库初始化失败: {str(e)}")

def log_message(log_type, message):
    """记录日志到数据库"""
    try:
        conn = sqlite3.connect(CONFIG['db_file'])
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO logs (log_type, message, created_at) VALUES (?, ?, CURRENT_TIMESTAMP)",
            (log_type, message)
        )
        conn.commit()
        conn.close()
        print(f"[{log_type}] {message}")
    except Exception as e:
        print(f"日志记录失败: {str(e)}")
        print(f"[{log_type}] {message}")

def is_event_processed(event_id):
    """检查事件是否已处理"""
    try:
        conn = sqlite3.connect(CONFIG['db_file'])
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM processed_events WHERE event_id = ?", (event_id,))
        count = cursor.fetchone()[0]
        conn.close()
        return count > 0
    except Exception as e:
        log_message("ERROR", f"检查事件处理状态失败: {str(e)}")
        return False

def mark_event_processed(event_id, appointment):
    """将事件标记为已处理"""
    try:
        conn = sqlite3.connect(CONFIG['db_file'])
        cursor = conn.cursor()
        
        subject = appointment.Subject if appointment.Subject else "无主题"
        organizer = appointment.Organizer
        start_time = appointment.Start.strftime('%H:%M')
        end_time = appointment.End.strftime('%H:%M')
        event_date = appointment.Start.strftime('%Y-%m-%d')
        location = appointment.Location if appointment.Location else ""
        
        cursor.execute('''
        INSERT INTO processed_events 
        (event_id, subject, organizer, start_time, end_time, event_date, location, processed_time) 
        VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ''', (event_id, subject, organizer, start_time, end_time, event_date, location))
        
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        log_message("ERROR", f"标记事件失败: {str(e)}")
        return False

def clean_old_events():
    """清理过期的事件数据"""
    try:
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=CONFIG['data_retention_days'])
        cutoff_str = cutoff_date.strftime('%Y-%m-%d %H:%M:%S')
        
        conn = sqlite3.connect(CONFIG['db_file'])
        cursor = conn.cursor()
        
        # 获取要删除的记录数
        cursor.execute("SELECT COUNT(*) FROM processed_events WHERE created_at < ?", (cutoff_str,))
        count = cursor.fetchone()[0]
        
        # 删除过期数据
        cursor.execute("DELETE FROM processed_events WHERE created_at < ?", (cutoff_str,))
        
        # 清理旧日志
        cursor.execute("DELETE FROM logs WHERE created_at < ?", (cutoff_str,))
        
        conn.commit()
        conn.close()
        
        if count > 0:
            log_message("INFO", f"已清理 {count} 条过期事件记录")
    except Exception as e:
        log_message("ERROR", f"清理过期数据失败: {str(e)}")

def generate_event_id(appointment):
    """生成事件的唯一ID，使用哈希以确保稳定性"""
    # 使用主题、开始时间、结束时间和组织者生成唯一标识
    event_str = f"{appointment.Subject}_{appointment.Start}_{appointment.End}_{appointment.Organizer}"
    # 使用SHA-256哈希确保ID的唯一性和一致性
    return hashlib.sha256(event_str.encode('utf-8')).hexdigest()

def main():
    """主函数"""
     # 初始化数据库
    init_database()

    client.start()
    log_message("INFO", f"启动会议室预订监控服务, 检查间隔: {CONFIG['check_interval_minutes']}分钟")
    
    # try:
    #     send_telegram_message("🔄 会议室预订监控服务已启动2")
    #     log_message("INFO", "已成功发送启动通知")
    # except Exception as e:
    #     log_message("ERROR", f"发送启动通知失败: {str(e)}")

    # log_message("INFO", f"启动会议室预订监控服务, 检查间隔: {CONFIG['check_interval_minutes']}分钟")
    
    # 立即执行一次
    check_room_bookings()
    
    # 设置定时任务
    schedule.every(CONFIG['check_interval_minutes']).minutes.do(check_room_bookings)
    
 # 保持运行
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        log_message("INFO", "服务被用户中断")
    except Exception as e:
        log_message("ERROR", f"服务运行错误: {str(e)}")
    finally:
        # 断开Telegram连接
        if client.is_connected():
            client.disconnect()
        log_message("INFO", "服务已停止")

if __name__ == "__main__":
    main()
