import smtplib
import email.utils
from email.mime.text import MIMEText

message = MIMEText("我是邮件的内容")
message['To'] = email.utils.formataddr(('接收者', '1697444387@qq.com'))
message['From'] = email.utils.formataddr(('发送者', '2196055715@qq.com'))
message['Subject'] = '我是邮件的标题'
server = smtplib.SMTP_SSL('smtp.qq.com', 465)
server.login('2196055715@qq.com','swjtyyurpobidjig')
server.set_debuglevel(True)
try:
    server.sendmail('2196055715@qq.com',['1697444387@qq.com'],msg=message.as_string())
finally:
    server.quit()