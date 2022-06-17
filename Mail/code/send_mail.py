# coding: utf-8 
# @Time    : 2020/8/9
# @Author  : caijinrong
# @Email   : kamwtsai@gmail.com
# @Link    : https://github.com/KamWTsai/AutoSendMsg

from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import parseaddr, formataddr
from email.header import Header
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import json

import schedule
import time
import os

current_path = os.path.abspath(__file__)
dir_path = os.path.dirname(os.path.dirname(current_path))


def loadSettings():
    with open(dir_path + os.sep + "settings.json", 'r', encoding='utf-8') as f:
        settings = json.load(f)
    return settings


class SendMail:
    server = ""
    port = ""
    senderName = ""
    from_addr = ""
    password = ""
    to_addr = ""
    subject = ""
    text_path = ""
    file_path = ""
    sender = ""
    text = ""
    img_name = ""
    log_path = ""
    msg = MIMEMultipart('mixed')

    def format_addr(self, s):
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode(), addr))

    def __init__(self, settings, img_name):
        self.server = settings['服务器地址']
        self.port = settings['端口']
        self.senderName = settings['发件人名称']
        self.from_addr = settings['发件人地址']
        self.password = settings['密码']
        self.to_addr = settings['收件人地址']  # 可以是一个列表
        self.subject = settings['标题']
        self.text_path = dir_path + os.sep + settings['正文路径']
        self.file_path = [dir_path + os.sep + x for x in settings['文件路径']]
        self.sender = '%s <%s>' % (self.senderName, self.from_addr)  # 构造发送者
        self.img_name = img_name
        self.log_path = dir_path + os.sep + "sendMailLog.log"
        # 添加header
        self.msg = MIMEMultipart('mixed')
        self.msg['Subject'] = Header(self.subject, 'utf-8')
        self.msg['To'] = ','.join(self.to_addr)
        self.msg['From'] = self.format_addr(self.sender)

    def writeToLog(self, from_addr, to_addr, subject, file_path, text):
        record = '发件人：%s, 收件人：%s, 标题：%s, 附件：%s' % (from_addr, str(to_addr), subject, str(file_path))
        log_text = '[' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + '] ' + record
        print(log_text)
        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.writelines(log_text+'\n')

    def load_text(self, text_path):
        with open(text_path, 'r', encoding='utf-8') as f:
            return f.read()

    def add_content(self):
        self.text = self.load_text(self.text_path)
                    # % '、'.join([x.replace('\\', '/').split('/')[-1] for x in self.file_path])
                    # % time.strftime("%Y年%m月%d日", time.localtime())
        self.msg.attach(MIMEText(self.text, 'html', 'utf-8'))

        img_file = open(self.img_name, 'rb')
        msg_img = MIMEImage(img_file.read())
        img_file.close()
        msg_img.add_header('Content-ID','<0>')
        self.msg.attach(msg_img)

        # file 附件
        for f in self.file_path:
            file_name = str(f).split(os.sep)[-1]
            msg_file = MIMEBase('application', 'octet-stream')
            msg_file.set_payload(open(f, 'rb').read())
            msg_file.add_header('Content-Disposition', 'attachment', filename=Header(file_name, 'utf-8').encode())
            encoders.encode_base64(msg_file)
            self.msg.attach(msg_file)

    def sendMail(self):
        self.add_content()
        # 普通连接，明文传输
        # smtp = smtplib.SMTP(server, port)
        # SSL加密
        smtp = smtplib.SMTP_SSL(self.server, self.port)
        smtp.login(self.from_addr, self.password)

        # to_addr可以是一个列表；msg.as_string()是email模块所构建的内容
        smtp.sendmail(self.from_addr, self.to_addr, self.msg.as_string())

        self.writeToLog(self.from_addr, self.to_addr, self.subject, self.file_path, self.text)

        # 结束回话
        smtp.quit()

if __name__ == '__main__':
    settings = loadSettings()
    aMail = SendMail(settings)
    aMail.sendMail()
