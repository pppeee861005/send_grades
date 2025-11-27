"""
批量发送成绩邮件脚本
读取 contacts_file.csv 文件，给每个联系人发送个性化的成绩通知邮件
"""
import csv
import smtplib
import ssl
import os
import sys
from email.mime.text import MIMEText
from dotenv import load_dotenv

# 加载.env文件 - 尝试从多个位置加载
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

# 尝试从脚本目录加载
env_path_script = os.path.join(script_dir, '.env')
# 尝试从项目根目录加载
env_path_root = os.path.join(project_root, '.env')

# 优先加载项目根目录的.env，如果不存在则加载脚本目录的.env
if os.path.exists(env_path_root):
    load_dotenv(env_path_root)
    print(f'从项目根目录加载.env: {env_path_root}')
elif os.path.exists(env_path_script):
    load_dotenv(env_path_script)
    print(f'从脚本目录加载.env: {env_path_script}')
else:
    # 尝试默认位置
    load_dotenv()
    print('尝试从默认位置加载.env')

def create_email_template(name, grade):
    """
    创建HTML格式的邮件模板
    """
    html_template = f"""
    <!doctype html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: Arial, "Microsoft JhengHei", sans-serif;
                line-height: 1.6;
                color: #333;
            }}
            .container {{
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
            }}
            h1 {{
                color: #2c3e50;
                border-bottom: 3px solid #3498db;
                padding-bottom: 10px;
            }}
            .grade-box {{
                background-color: #ecf0f1;
                padding: 15px;
                border-radius: 5px;
                margin: 20px 0;
                text-align: center;
            }}
            .grade {{
                font-size: 48px;
                font-weight: bold;
                color: #27ae60;
            }}
            .message {{
                margin-top: 20px;
                font-size: 16px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>成績通知</h1>
            <p>親愛的 {name} 同學：</p>
            <div class="grade-box">
                <p>你的學期成績是</p>
                <div class="grade">{grade}</div>
            </div>
            <div class="message">
                <p>恭喜你完成本學期的課程！</p>
                <p>希望你繼續保持努力學習的態度，在未來的學習中取得更好的成績。</p>
            </div>
            <hr>
            <p style="color: #7f8c8d; font-size: 12px;">此為系統自動發送郵件，請勿直接回覆。</p>
        </div>
    </body>
    </html>
    """
    return html_template

def read_contacts(csv_file_path):
    """
    读取CSV文件中的联系人信息
    返回联系人列表：[(name, email, grade), ...]
    """
    contacts = []
    try:
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # 跳过标题行
            
            for row in reader:
                # 跳过空行
                if not row or len(row) < 3:
                    continue
                
                name = row[0].strip()
                email = row[1].strip()
                grade = row[2].strip()
                
                # 验证基本格式
                if name and email and grade:
                    contacts.append((name, email, grade))
                    
    except FileNotFoundError:
        print(f'错误: 找不到CSV文件: {csv_file_path}')
        sys.exit(1)
    except Exception as e:
        print(f'读取CSV文件时出错: {e}')
        sys.exit(1)
    
    return contacts

def send_emails(contacts, sender_email, password):
    """
    发送邮件给所有联系人
    """
    smtp_server = "smtp.gmail.com"
    port = 587
    
    # 创建SSL上下文
    context = ssl.create_default_context()
    
    success_count = 0
    fail_count = 0
    
    try:
        # 建立SMTP连接
        server = smtplib.SMTP(smtp_server, port)
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(sender_email, password)
        print(f'\n成功登录邮件服务器: {sender_email}\n')
        
        # 遍历每个联系人发送邮件
        for name, email, grade in contacts:
            try:
                # 创建邮件内容
                html_content = create_email_template(name, grade)
                
                # 创建MIMEText对象
                msg = MIMEText(html_content, 'html', 'utf-8')
                msg['Subject'] = '成績通知'
                msg['From'] = sender_email
                msg['To'] = email
                
                # 发送邮件
                server.sendmail(sender_email, [email], msg.as_string())
                print(f'✓ 成功发送邮件给 {name} ({email}) - 成绩: {grade}')
                success_count += 1
                
            except Exception as e:
                print(f'✗ 发送邮件给 {name} ({email}) 时出错: {e}')
                fail_count += 1
                # 继续发送其他邮件
                continue
        
        # 关闭连接
        server.quit()
        
    except smtplib.SMTPAuthenticationError:
        print('错误: 邮件服务器认证失败，请检查邮箱地址和密码')
        sys.exit(1)
    except smtplib.SMTPException as e:
        print(f'错误: SMTP服务器错误: {e}')
        sys.exit(1)
    except Exception as e:
        print(f'错误: 连接邮件服务器时出错: {e}')
        sys.exit(1)
    
    # 显示发送结果统计
    print(f'\n{"="*50}')
    print(f'发送完成！')
    print(f'成功: {success_count} 封')
    print(f'失败: {fail_count} 封')
    print(f'总计: {len(contacts)} 封')
    print(f'{"="*50}')

def main():
    """
    主函数
    """
    # 设置路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_file_path = os.path.join(script_dir, 'contacts_file.csv')
    
    # 读取联系人信息
    print('正在读取联系人信息...')
    contacts = read_contacts(csv_file_path)
    
    if not contacts:
        print('错误: CSV文件中没有有效的联系人信息')
        sys.exit(1)
    
    print(f'找到 {len(contacts)} 个联系人')
    for name, email, grade in contacts:
        print(f'  - {name}: {email} (成绩: {grade})')
    
    # 获取发送者邮箱
    sender_email = os.getenv('GMAIL_SENDER', 'peter861005@gmail.com')
    
    # 获取密码 - 优先从环境变量读取
    password = (os.getenv('GMAIL_PASSWORD') or 
                os.getenv('PASSWORD') or 
                os.getenv('EMAIL_PASSWORD') or
                os.getenv('MAIL_PASSWORD'))
    
    # 如果环境变量中没有密码，则提示用户输入
    if not password:
        print(f'\n未找到环境变量中的密码')
        print('请输入Gmail应用密码（用于 {sender_email}）:')
        password = input('密码: ').strip()
    
    if not password:
        print('错误: 未提供密码，无法发送邮件')
        sys.exit(1)
    
    # 确认发送
    print(f'\n准备发送邮件...')
    print(f'发送者: {sender_email}')
    print(f'收件人数量: {len(contacts)}')
    confirm = input('\n确认发送？(y/n): ').strip().lower()
    
    if confirm != 'y':
        print('已取消发送')
        sys.exit(0)
    
    # 发送邮件
    send_emails(contacts, sender_email, password)

if __name__ == '__main__':
    main()

