import yagmail
import os
import re
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
    print(f'尝试从默认位置加载.env')

# 解析收件人列表
def parse_recipients(file_path):
    """从name_list.md读取并解析收件人邮箱地址"""
    recipients = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                # 移除引号和逗号
                email = line.strip('",').strip()
                # 验证邮箱格式
                if re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
                    recipients.append(email)
    except Exception as e:
        print(f'读取收件人列表时出错: {e}')
    return recipients

# 创建HTML邮件正文
def create_html_body(letter_number):
    """创建包含笑话、链接和图片的HTML正文"""
    joke = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
    </head>
    <body>
        <h2 style="color: #2c3e50;">第{letter_number}封信：內部測試</h2>
        <p style="color: #e74c3c; font-weight: bold;">修改html格式</p>
        <hr>
        <h2>测试邮件笑话</h2>
        <p>为什么程序员总是分不清万圣节和圣诞节？</p>
        <p><strong>因为 Oct 31 = Dec 25！</strong></p>
        <p>（八进制31等于十进制25）</p>
        <hr>
        <p>test...</p>
        <p>点击这里访问：<a href="https://google.com">clickme</a></p>
        <br>
        <p style="font-size: 16px; color: #333; line-height: 1.6;">
            老師我終於會用py發mail了，真是高興。太有趣了！
        </p>
        <br>
        <h3>图片展示：</h3>
        <img src="16661" alt="16661.png" style="max-width: 100%; height: auto; display: block; margin: 20px auto;">
    </body>
    </html>
    """
    return joke

# 主函数
def main():
    # 设置路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    name_list_path = os.path.join(script_dir, 'name_list.md')
    pdf_path = os.path.join(script_dir, '木箱裡的名字精靈小隊.pdf')
    # 图片文件名是 "image (16661.png"，注意文件名包含空格和括号
    # 使用绝对路径，并确保路径格式正确
    image_path = os.path.abspath(os.path.join(script_dir, 'image (16661.png'))
    
    # 打印路径用于调试
    print(f'图片路径: {image_path}')
    print(f'图片文件是否存在: {os.path.exists(image_path)}')
    
    # 内部测试：只发送给 peter861005@gmail.com
    test_recipient = "peter861005@gmail.com"
    recipients = [test_recipient]
    print(f'内部测试模式：只发送给 {test_recipient}')
    
    # 检查文件是否存在
    if not os.path.exists(image_path):
        print(f'错误: 图片文件不存在: {image_path}')
        return
    
    if not os.path.exists(pdf_path):
        print(f'错误: PDF文件不存在: {pdf_path}')
        return
    
    # 设置yagmail
    sender_email = "peter861005@gmail.com"
    # 从.env文件读取密码 - 尝试多个可能的变量名
    password = (os.getenv('GMAIL_PASSWORD') or 
                os.getenv('PASSWORD') or 
                os.getenv('EMAIL_PASSWORD') or
                os.getenv('MAIL_PASSWORD'))
    
    if not password:
        print('错误: 未找到密码')
        print(f'请检查.env文件是否存在:')
        print(f'  - {env_path_root}')
        print(f'  - {env_path_script}')
        print('并在.env文件中设置以下任一变量:')
        print('  GMAIL_PASSWORD=你的密码')
        print('  或 PASSWORD=你的密码')
        print('  或 EMAIL_PASSWORD=你的密码')
        print('  或 MAIL_PASSWORD=你的密码')
        return
    
    yag = yagmail.SMTP(sender_email, password)
    
    # 准备附件（PDF作为附件，图片既嵌入又作为附件）
    attachments = [pdf_path, image_path]
    
    # 发送两封测试邮件：第四封和第五封
    # 之前已发送：第一封(yagmail test)、第二封(Python邮件发送测试)、第三封(修正圖片路徑)
    print(f'准备发送邮件，图片路径: {image_path}')
    
    for letter_num in [4, 5]:
        # 创建HTML正文（包含信件编号）
        html_body = create_html_body(letter_num)
        
        try:
            yag.send(
                to=recipients,
                subject=f"Python邮件发送测试 - 第{letter_num}封信",
                contents=[
                    html_body,
                    {"16661": image_path}  # yagmail会自动处理CID，HTML中使用src="16661"即可
                ],
                attachments=attachments  # 图片既作为内联显示，也作为附件
            )
            print(f'第{letter_num}封信发送成功！')
        except Exception as e:
            print(f'发送第{letter_num}封信时出错: {e}')

if __name__ == '__main__':
    main()

