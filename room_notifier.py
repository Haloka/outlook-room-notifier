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
        calendar = namespace.GetSharedDefaultFolder(recipient, 9)  # 9 = olFolderCalendar
        
        # 计算时间范围
        # now = datetime.datetime.now()
        # end_time = now + datetime.timedelta(hours=1)  # 查看未来1小时的预订
        now = datetime.datetime.now() + datetime.timedelta(hours=-1)
        end_time = now + datetime.timedelta(hours=72)  # 查看未来1小时的预订

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
            
            # 检查是否已处理过此事件
            if not is_event_processed(event_id):
                # 获取事件详情
                organizer = appointment.Organizer
                subject = get_subject(appointment)
                
                # 格式化时间
                start_time = appointment.Start
                end_time = appointment.End
                end_time_str = end_time.strftime('%H:%M')
                formatted_start = format_date_with_today_tomorrow(start_time)

                message = ""
                
                # 如果有位置信息，添加到消息中
                if appointment.Location:
                    message += f"🔔 会议室: {appointment.Location}\n\n"
                else:
                    message += "🔔 会议室: 没有填入位置\n\n"
                
                message += (
                    f"时间: {formatted_start} - {end_time_str}\n"
                    f"主题: {subject}\n"
                    f"预订: {organizer}\n"
                )

                body = appointment.Body
                if body:
                    message += f"\n\n{body}"

                # 发送Telegram通知
                if send_telegram_message(message):
                    # 标记事件为已处理
                    if mark_event_processed(event_id, appointment):
                        log_message("INFO", f"已发送通知: {appointment.Subject}_{appointment.Start}_{appointment.End}_{appointment.Organizer}")
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
        client.send_message(CONFIG['telegram_chat_id'], "```\n" + message + "\n```")
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
        
        subject = get_subject(appointment)
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
    """生成事件的唯一ID"""
    subject = get_subject(appointment)
    startTime = appointment.Start.strftime('%Y-%m-%d %H:%M:%S')
    endTime = appointment.End.strftime('%Y-%m-%d %H:%M:%S')
    organizer = appointment.Organizer if appointment.Organizer else "未知"
    location = appointment.Location if appointment.Location else "无地点"
    event_str = f"{startTime}_{endTime}_{subject}_{organizer}_{location}"
    
    return event_str

def get_subject(appointment):
    """获取约会的主题"""
    subject = appointment.ConversationTopic if appointment.ConversationTopic else appointment.Subject
    subject = subject if subject else "无主题"
    return subject

def format_date_with_today_tomorrow(appointment_time):
    today = datetime.datetime.now().date()
    tomorrow = today + datetime.timedelta(days=1)
    appointment_date = appointment_time.date()
    
    start_time = appointment_time.strftime('%H:%M')
    
    if appointment_date == today:
        return f"今天 {start_time}"
    elif appointment_date == tomorrow:
        return f"明天 {start_time}"
    else:
        return f"{appointment_time.strftime('%m月%d日')} {start_time}"
    

def main():
    """主函数"""
     # 初始化数据库
    init_database()

    client.start()
    log_message("INFO", f"启动会议室预订监控服务, 检查间隔: {CONFIG['check_interval_minutes']}分钟")
    
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
