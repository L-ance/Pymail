# coding = utf-8
import os
import sys
import time
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
# 20190318/sopuy/python3.6
# author;Lance

curtime = time.strftime('%Y%m%d-%H:%M:%S')


class Pymail:
    def __init__(self, **kwargs):
        self._smtp_host = kwargs.get("_smtp_host", '')          # 邮箱SMTP地址
        self._smtp_port = kwargs.get("_smtp_port", 25)          # 邮箱SMTP端口
        self._sender = kwargs.get("_sender", '')                # 发件邮箱
        self._sender_pwd = kwargs.get("_sender_pwd", '')        # 发件邮箱密码
        self._sender_name = kwargs.get("_sender_name", '')      # 发件人名称
        self._receivers = kwargs.get("_receivers", [])          # 收件人
        self._cc = kwargs.get("_cc", [])                        # 抄送地址
        self._bcc = kwargs.get("_bcc", [])                      # 密送地址
        self._subject = kwargs.get("_subject", '')              # 邮箱主题
        self._content = kwargs.get("_content", '')              # 邮箱内容,html格式
        self._img = kwargs.get("_img", '')                      # 正文图片
        self._atts = kwargs.get("_atts", [])                    # 附件
        self._curdate = kwargs.get("_curdate", '')              # 处理附件路径替换字符串

    def send_text(self):
        msg = MIMEMultipart('mixed')
        # msg['From'] = formataddr((Header(self._sender_name, 'utf-8').encode(), self._sender))
        msg['From'] = formataddr((self._sender_name, self._sender))
        msg['To'] = ";".join(self._receivers)
        msg['Subject'] = self._subject
        toaddrs = self._receivers
        if self._content:
            msg.attach(MIMEText(self._content, 'html', _charset='utf-8'))
        # 添加抄送
        if self._cc:
            msg.add_header('Cc', ';'.join(self._cc))
            toaddrs = self._receivers + self._cc
        # 添加密送
        if self._bcc:
            msg.add_header('Bcc', ';'.join(self._bcc))
            toaddrs = toaddrs + self._bcc
        smtp = smtplib.SMTP()
        smtp.connect(self._smtp_host, self._smtp_port)
        smtp.login(self._sender, self._sender_pwd)
        smtp.sendmail(self._sender, toaddrs, msg.as_string())
        smtp.quit()
        print(curtime + ' txt email send ok.')

    @staticmethod
    def compose_mail(**kwargs):
        msg = MIMEMultipart('mixed')
        # msg["Accept-Language"] = "zh-CN"
        # msg["Accept-Charset"] = "ISO-8859-1,utf-8"
        # msg['From'] = formataddr((Header(self._sender_name, 'utf-8').encode(), self._sender))
        msg.add_header('From', formataddr([kwargs.get("_sender_name"), kwargs.get("_sender")]))
        msg.add_header('To', ";".join(kwargs.get("_receivers")))
        msg.add_header('Subject', kwargs.get("_subject"))
        # 正文添加图片
        if kwargs.get("_img"):
            img_path = kwargs.get("_img").replace("YYYYMMDD", kwargs.get("_curdate"))
            if os.path.exists(img_path):
                with open(img_path, 'rb') as f:
                    imgmsg = MIMEImage(f.read())
                imgmsg.add_header('Content-ID', '<image1>')
                msg.attach(imgmsg)
                msg.attach(MIMEText(kwargs.get("_content") + "<br><p><img src='cid:image1'/></p>", 'html', _charset='utf-8'))
            else:
                msg.attach(MIMEText(kwargs.get("_content") +
                                    "<br><font size='6' color=#FF0000>正文图片文件异常,请及时联系发送方处理.</font><br>", 'html', _charset='utf-8'))
        else:
            msg.attach(MIMEText(kwargs.get("_content"), 'html', _charset='utf-8'))
        # 添加抄送
        if kwargs.get("_cc"):
            msg.add_header('Cc', ';'.join(kwargs.get("_cc")))
        # 添加密送
        if kwargs.get("_bcc"):
            msg.add_header('Bcc', ';'.join(kwargs.get("_bcc")))
        # 添加附件
        if kwargs.get("_atts"):
            atts = kwargs.get("_atts")
            for i in range(0, len(atts)):
                file_path = atts[i].replace("YYYYMMDD", kwargs.get("_curdate"))
                if os.path.exists(file_path):
                    att = MIMEApplication(open(file_path, 'rb').read())
                    att.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
                    msg.attach(att)
                else:
                    msg.attach(MIMEText("<br><font size='6' color=#FF0000>附件文件异常,请及时联系发送方处理.</font><br>", 'html'))
                    break
        return msg

    @staticmethod
    def send_mail(**kwargs):
        msg = pymail.compose_mail(**kwargs).as_string()
        toaddrs = kwargs.get("_receivers")
        if kwargs.get("_cc"):
            toaddrs = kwargs.get("_receivers") + kwargs.get("_cc")
        if kwargs.get("_bcc"):
            toaddrs = toaddrs + kwargs.get("_bcc")
        smtp = smtplib.SMTP()
        smtp.connect(kwargs.get("_smtp_host"), kwargs.get("_smtp_port"))
        # smtp.starttls()
        smtp.login(kwargs.get("_sender"), kwargs.get("_sender_pwd"))
        smtp.sendmail(kwargs.get("_sender"), toaddrs, msg)
        smtp.quit()
        print(curtime + ' email send ok.')
