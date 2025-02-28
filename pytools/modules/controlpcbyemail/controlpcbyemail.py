'''
Function:
    邮件控制电脑
'''
import re
import os
import json
import time
import email
import poplib
import smtplib
import datetime
from email import encoders
from email.header import Header
from email.parser import Parser
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import decode_header
from email.utils import parseaddr, formataddr
from email.mime.multipart import MIMEMultipart


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
        for arg in args:
            if arg == 'stat':
                res[arg] = self.pop3_server.stat()
            elif arg == 'list':
                res[arg] = self.pop3_server.list()
            elif arg == 'latest':
                mails = self.pop3_server.list()[1]
                resp, lines, octets = self.pop3_server.retr(len(mails))
                msg = b'\r\n'.join(lines).decode('utf-8')
                msg = Parser().parsestr(msg)
                result = self.__parsemessage(msg)
                res[arg] = result
            elif type(arg) == int:
                mails = self.pop3_server.list()[1]
                if arg > len(mails):
                    res[arg] = None
                    continue
                resp, lines, octets = self.pop3_server.retr(arg)
                msg = b'\r\n'.join(lines).decode('utf-8')
                msg = Parser().parsestr(msg)
                result = self.__parsemessage(msg)
                res[arg] = result
            else:
                res[arg] = None
        return res

    '''SMTP协议发送邮件'''

    def send(self, content=None, attach_path=None):
        options = self.options
        msg = MIMEMultipart()
        # Smtp server
        if not options['receiver']['enable_ssl']:
            smtp_server = smtplib.SMTP(options['receiver']['smtp_server'], 25)
            smtp_server.set_debuglevel(1)
            smtp_server.login(
                options['receiver']['email'], options['receiver']['password']
            )
        else:
            if options['receiver']['port']:
                try:
                    smtp_server = smtplib.SMTP_SSL(
                        options['receiver']['smtp_server'], options['receiver']['port']
                    )
                except:
                    smtp_server = smtplib.SMTP_SSL(options['receiver']['smtp_server'])
            else:
                smtp_server = smtplib.SMTP_SSL(options['receiver']['smtp_server'])
            smtp_server.set_debuglevel(1)
            smtp_server.ehlo(options['receiver']['smtp_server'])
            smtp_server.login(
                options['receiver']['email'], options['receiver']['password']
            )
        # From
        msg_from = 'Server <%s>' % options['receiver']['email']
        from_name, from_addr = parseaddr(msg_from)
        msg['From'] = formataddr((Header(from_name, 'utf-8').encode(), from_addr))
        # To
        msg_to = 'Controller <%s>' % options['sender']['email']
        to_name, to_addr = parseaddr(msg_to)
        msg['To'] = formataddr((Header(to_name, 'utf-8').encode(), to_addr))
        # Time
        msg['Date'] = Header(
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'utf-8'
        )
        if attach_path is None and content is None:
            return False
        if content is None:
            content = ['Attachment', 'Attachment from your computer...']
        msg['Subject'] = Header(content[0])
        msg.attach(MIMEText(content[1], 'plain', 'utf-8'))
        if attach_path:
            with open(attach_path, 'rb') as f:
                filename = os.path.basename(attach_path)
                mime = MIMEBase(
                    'attachment', filename.split('.')[-1], filename=filename
                )
                mime.add_header('Content-Disposition', 'attachment', filename=filename)
                mime.add_header('Content-ID', '<0>')
                mime.add_header('X-Attachment-Id', '0')
                mime.set_payload(f.read())
                encoders.encode_base64(mime)
                msg.attach(mime)
        smtp_server.sendmail(from_addr, [to_addr], msg.as_string())
        smtp_server.quit()
        return True

    '''关闭pop3连接'''

    def closepop(self):
        self.pop3_server.quit()

    '''重置pop3连接'''

    def resetpop(self):
        options = self.options
        self.closepop()
        self.pop3_server = poplib.POP3(options['receiver']['pop3_server'])
        self.pop3_server.user(options['receiver']['email'])
        self.pop3_server.pass_(options['receiver']['password'])

    '''解析邮件'''

    def __parsemessage(self, msg):
        result = {}
        for header in ['From', 'To', 'Subject']:
            result[header] = None
            temp = msg.get(header, '')
            if temp:
                if header == 'Subject':
                    value, charset = decode_header(temp)[0]
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
        # 不考虑MIMEMultipart对象
        if not msg.is_multipart():
            content_type = msg.get_content_type()
            # 只考虑纯文本/HTML内容
            if content_type == 'text/plain' or content_type == 'text/html':
                content = msg.get_payload(decode=True)
                charset = msg.get_charset()
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
        self.num_msg = len(self.email.get('list')['list'][1])
        self.time_interval = time_interval

    '''运行服务器'''

    def run(self):
        options, word2cmd_dict = self.options, self.word2cmd_dict
        self.printinfo()
        print('[INFO]:Start server successfully...')
        while True:
            self.email.resetpop()
            mails = self.email.get('list')['list'][1]
            if len(mails) > self.num_msg:
                for i in range(self.num_msg + 1, len(mails) + 1):
                    res = self.email.get(i)
                    res_from = res[i]['From']
                    res_from = re.findall(r'<(.*?)>', res_from)[0].lower()
                    print(res_from)
                    if res_from != options['sender']['email'].lower():
                        continue
                    command = res[i]['Subject']
                    if command in word2cmd_dict:
                        command = word2cmd_dict[command]
                    if command == 'screenshot':
                        savename = './screenshot.jpg'
                        self.screenshot(savename)
                        try:
                            is_success = self.email.send(attach_path=savename)
                            if not is_success:
                                raise RuntimeError('Fail to send screenshot...')
                            print('[INFO]: Send screenshot successfully...')
                        except:
                            print('[Error]: Fail to send screenshot...')
                    else:
                        self.runcmd(command)
                self.num_msg = len(mails)
            time.sleep(self.time_interval)

    '''os.system()运行命令cmd'''

    def runcmd(self, cmd):
        try:
            os.system(cmd)
            print('[INFO]: Run <%s> successfully...' % cmd)
            return True
        except:
            print('[Error]: Fail to Run <%s>...' % cmd)
            return False

    '''截屏'''

    def screenshot(self, savename='screenshot.jpg'):
        from PIL import ImageGrab

        img = ImageGrab.grab()
        img.save(savename)
        print('[INFO]: Get %s successfully...' % savename)

    '''打印欢迎信息'''

    def printinfo(self):
        print('*' * 20 + 'Welcome' + '*' * 20)
        print('[Function]: Control your computer by your email')
        print('[Author]: Car')
        print('[微信公众号]: Car的皮皮')
