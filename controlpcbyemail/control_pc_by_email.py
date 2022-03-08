# -*- encoding: utf-8 -*-
'''
@File    :   control_pc_by_email.py
@Time    :   2022/03/07 20:02:15
@Author  :   朱峻熠
@Version :   1.0
@Contact :   xxx@qq.com
@Function:   用email实现远程控制电脑
'''

import re
import os
import json
import time
import poplib
import smtplib
from email import encoders
from email.header import Header
from email.parser import Parser
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import decode_header
from email.utils import parseaddr, formataddr
from email.mime.multipart import MIMEMultipart
import zjy_tools

'''邮箱类'''


class EmailClass:
    def __init__(self, options):
        self.options = options
        # POP3协议接收邮件
        self.pop3_server = poplib.POP3(options['receiver']['pop3_server'])
        self.pop3_server.user(options['receiver']['email'])
        self.pop3_server.pass_(options['receiver']['password'])

    '''接收邮件'''

    def get(self, *args):
        res = {}
        # 从收集参数中确定我需要取得邮件信息类型
        for arg in args:
            # 邮箱状态
            if arg == 'stat':
                res[arg] = self.pop3_server.stat()
            # 邮箱邮件的索引列表
            elif arg == 'list':
                res[arg] = self.pop3_server.list()
            # 取得最新的一封邮件信息
            elif arg == 'latest':
                mails = self.pop3_server.list()[1]
                # 邮件是从1开始编号，最新一封邮件的序号就是邮件列表的长度
                resp, lines, octets = self.pop3_server.retr(len(mails))
                # lines为邮件message每一行内容，拼接后，可以得到完整的邮件
                msg = b'\r\n'.join(lines).decode('utf-8')
                # 解析并返回一个email.message对象
                msg = Parser().parsestr(msg)
                # 解析message中具体内容
                result = self.__parse_message(msg)
                res[arg] = result
            # 取得指定的邮件信息
            elif type(arg) == int:
                mails = self.pop3_server.list()[1]
                if arg > len(mails):
                    res[arg] = None
                    continue
                resp, lines, octets = self.pop3_server.retr(arg)
                msg = b'\r\n'.join(lines).decode('utf-8')
                msg = Parser().parsestr(msg)
                result = self.__parse_message(msg)
                res[arg] = result
            else:
                res[arg] = None
        return res

    '''SMTP协议发送邮件'''

    def send(self, subject=None, content=None, attach_path=None):
        options = self.options
        # 实例化MIME邮件对象
        msg = MIMEMultipart()
        # SMTP服务器设置并登陆打开连接
        smtp_server = smtplib.SMTP(options['receiver']['smtp_server'], 25)
        smtp_server.set_debuglevel(1)
        smtp_server.login(options['receiver']['email'], options['receiver']['password'])
        # 发件人设置
        from_name = '远程肉鸡'
        from_addr = options['receiver']['email']
        msg['From'] = formataddr((Header(from_name, 'utf-8').encode(), from_addr))
        # 收件人设置
        to_name = '指令长'
        to_addr = options['sender']['email']
        msg['To'] = formataddr((Header(to_name, 'utf-8').encode(), to_addr))
        # 如果邮件没有主题和正文，也没有附件，则停止发送邮件，返回发送失败
        if attach_path is None and subject is None and content is None:
            return False
        # 如果邮件主题和正文为空，那么给出默认的主题和内容
        if subject is None:
            subject = '肉鸡信息'
        if content is None:
            content = '来自肉鸡的全真影像，哈哈！'
        # 组装邮件主题
        msg['Subject'] = Header(subject)
        # 将邮件正文附上
        msg.attach(MIMEText(content, 'plain', 'utf-8'))
        # 如果有附件，者发送附件
        if attach_path:
            with open(attach_path, 'rb') as f:
                # 构造一个MIME邮件附件
                filename = os.path.basename(attach_path)
                mime = MIMEBase(
                    'attachment', filename.split('.')[-1], filename=filename
                )
                mime.add_header('Content-Disposition', 'attachment', filename=filename)
                mime.add_header('Content-ID', '<0>')
                mime.add_header('X-Attachment-Id', '0')
                mime.set_payload(f.read())
                encoders.encode_base64(mime)
                # 将邮件附件附上
                msg.attach(mime)
        # 发送邮件
        smtp_server.sendmail(from_addr, [to_addr], msg.as_string())
        # 关闭服务器连接
        smtp_server.quit()
        return True

    '''关闭pop3连接'''

    def close_pop(self):
        self.pop3_server.quit()

    '''重置pop3连接'''

    def reset_pop(self):
        options = self.options
        self.close_pop()
        self.pop3_server = poplib.POP3(options['receiver']['pop3_server'])
        self.pop3_server.user(options['receiver']['email'])
        self.pop3_server.pass_(options['receiver']['password'])

    '''解析邮件'''

    def __parse_message(self, msg):
        result = {}
        # 取得message中发件人，收件人，主题等信息
        for header in ['From', 'To', 'Subject']:
            result[header] = None
            # 从message中提取特定的字段信息
            temp = msg.get(header, '')
            if temp:
                if header == 'Subject':
                    # 获取邮件头的指定信息，并得到编码类型
                    value, charset = decode_header(temp)[0]
                    # 根据编码类型解码，避免内容乱码
                    if charset:
                        value = value.decode(charset)
                    result[header] = value
                else:
                    name, addr = parseaddr(temp)
                    value, charset = decode_header(name)[0]
                    if charset:
                        value = value.decode(charset)
                    result[header] = '%s<%s>' % (value, addr)
        result['Text'] = None
        # 暂时不考虑MIMEMultipart对象
        if not msg.is_multipart():
            content_type = msg.get_content_type()
            # 只考虑纯文本/HTML内容
            if content_type == 'text/plain' or content_type == 'text/html':
                # 载入邮件正文
                content = msg.get_payload(decode=True)
                # 尝试取得邮件正文的编码类型
                charset = msg.get_charset()
                # 如果上一步没有取得编码类型，那么通过邮件解析查找charset关键字来取得编码类型
                if charset is None:
                    temp = msg.get('Content-Type', '').lower()
                    pos = temp.find('charset=')
                    if pos >= 0:
                        charset = temp[pos + 8 :].strip()
                if charset:
                    content = content.decode(charset)
                result['Text'] = content
        return result


'''邮件控制电脑'''


class ControlPCbyEmail:
    tool_name = '邮件控制电脑'

    def __init__(self, time_interval=5, **kwargs):
        # 如果类初始化时包括了参数，则使用初始化参数，否则读配置文件
        # 虽然也可以通过字典收集参数配置服务器及命令等设置，但强烈建议用配置文件
        if 'options' in kwargs:
            self.options = kwargs['options']
        else:
            rootdir = os.path.split(os.path.abspath(__file__))[0]
            with open(os.path.join(rootdir, 'resources/config.json'), 'r') as f:
                self.options = json.load(f)
        if 'word2cmd' in kwargs:
            self.word2cmd_dict = kwargs['word2cmd_dict']
        else:
            with open(
                os.path.join(rootdir, 'resources/word2cmd.json'), 'r', encoding='utf-8'
            ) as f:
                self.word2cmd_dict = json.load(f)
        self.email = EmailClass(self.options)
        # 取得目前邮件列表的数量并保存到变量
        self.num_msg = len(self.email.get('list')['list'][1])
        # 服务器检查邮件的间隔时间，单位：秒
        self.time_interval = time_interval

    '''运行服务器'''

    def run(self):
        options, word2cmd_dict = self.options, self.word2cmd_dict
        self.print_info()
        print('[INFO]:服务器成功启动...')
        while True:
            # 重置pop3连接，以防断线
            self.email.reset_pop()
            mails = self.email.get('list')['list'][1]
            # 比较邮件列表数量来判断是否有新邮件
            if len(mails) > self.num_msg:
                for i in range(self.num_msg + 1, len(mails) + 1):
                    # 取得新邮件
                    res = self.email.get(i)
                    # 取得发件人邮箱地址
                    res_from = res[i]['From']
                    res_from = re.findall(r'<(.*?)>', res_from)[0].lower()
                    # 判断是否为合法的发件人邮箱
                    if res_from != options['sender']['email'].lower():
                        continue
                    # 利用邮件主题作为命令
                    command = res[i]['Subject']
                    # 从配置文件中读出真正的命令行字符串
                    if command in word2cmd_dict:
                        command = word2cmd_dict[command]
                    # 截屏为特殊命令，用截屏函数执行
                    if command == 'screenshot':
                        # 保存截屏文件名
                        savename = './screenshot.jpg'
                        # 调用截屏函数
                        self.screenshot(savename)
                        # 邮件发送
                        try:
                            is_success = self.email.send(attach_path=savename)
                            if not is_success:
                                raise RuntimeError('发送截图失败...')
                            print('[INFO]: 截屏图成功发送...')
                        except:
                            print('[Error]: 发送截图失败...')
                    else:
                        self.run_cmd(command)
                self.num_msg = len(mails)
            # 休眠几秒
            time.sleep(self.time_interval)

    '''os.system()运行命令cmd'''

    def run_cmd(self, cmd):
        try:
            os.system(cmd)
            print('[INFO]: 运行 <%s> 命令成功...' % cmd)
            return True
        except:
            print('[Error]: 运行 <%s> 命令失败...' % cmd)
            return False

    '''截屏，默认保存当前目录'''

    def screenshot(self, savename='screenshot.jpg'):
        from PIL import ImageGrab

        # 截屏函数
        img = ImageGrab.grab()
        img.save(savename)
        print('[INFO]: 成功保存截图文件 %s ...' % savename)

    '''打印欢迎信息'''

    def print_info(self):
        print('*' * 20 + '欢迎您使用本程序' + '*' * 20)
        print('[功能]: 用email远程控制你的电脑')
        print('[作者]: 朱峻熠')


'''开始主程序'''
if __name__ == '__main__':
    control = ControlPCbyEmail()
    control.run()
