from tkinter import *
import tkinter.messagebox as messagebox
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

from email.parser import Parser
from email.header import decode_header

import smtplib
import poplib

password = []
email = []
pop3_server = []
smtp_server = []

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def create_bg(widget):
    widget.canvas = Canvas(widget, width=500,height=500, bg='LightGrey', bd=0, highlightthickness=0)
    widget.photo = PhotoImage(file="bg.gif")
    widget.canvas.create_image(200, 200, image = widget.photo)
    widget.canvas.grid(row=0,column=1,rowspan=20,columnspan=10,
        sticky=W+E+N+S, padx=5, pady=5)

class MainApplication(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid(row=0,column=0,sticky=E+N)
        self.createWidgets()
    def createWidgets(self):
        self.canvas = Canvas(self, width=100,height=500, bg='DeepSkyBlue', 
            bd=0, highlightthickness=0)
        self.photo = PhotoImage(file="bg.gif")
        self.canvas.create_image(400, 200, image = self.photo)
        self.canvas.grid(row=0,column=0,rowspan=20,columnspan=4,
            sticky=W+E+N+S, padx=5, pady=5)
        
        self.enrollButton = Button(self, text='登陆', command=self.enroll)
        self.enrollButton.grid(row=0,column=0,sticky=W+N, padx=5, pady=5)          
        self.sendButton = Button(self, text='写信', command=self.send)
        self.sendButton.grid(row=1,column=0,sticky=W+N, padx=5, pady=5)
        self.receiveButton = Button(self, text='收件箱', command=self.receive)
        self.receiveButton.grid(row=2,column=0,sticky=W+N, padx=5, pady=5)
        
        self.App = EnrollApplication()
    def send(self):
        self.App.destroy()
        self.App = SendApplication()
    def receive(self):
        self.App.destroy()
        self.App = ReceiveApplication()
    def enroll(self):
        self.App.destroy()
        self.App = EnrollApplication()
    
class EnrollApplication(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid(row=0,column=1,sticky=E+N)
        self.createWidgets()
    def createWidgets(self):
        create_bg(self)
        self.emailLabel = Label(self, text='邮箱账号')
        self.emailLabel.grid(row=0,column=1,sticky=E)
        self.emailEntry = Entry(self, width=60)
        self.emailEntry.grid(row=0,column=2,sticky=W)
        self.emailEntry.insert(0, 'zhuchengdang@qq.com')
        
        self.pwdLabel = Label(self, text='邮箱密码')
        self.pwdLabel.grid(row=1,column=1,sticky=E)
        self.pwdEntry = Entry(self, width=60, show='*')
        self.pwdEntry.grid(row=1,column=2,sticky=W)

        self.popLabel = Label(self, text='pop3服务器')
        self.popLabel.grid(row=2,column=1,sticky=E)
        self.popEntry = Entry(self, width=60)
        self.popEntry.grid(row=2,column=2,sticky=W)
        self.popEntry.insert(0, 'pop.qq.com')

        self.smtpLabel = Label(self, text='smtp服务器')
        self.smtpLabel.grid(row=3,column=1,sticky=E)
        self.smtpEntry = Entry(self, width=60)
        self.smtpEntry.grid(row=3,column=2,sticky=W)
        self.smtpEntry.insert(0, 'smtp.qq.com')
        
        self.enrollButton = Button(self, text='确认登陆', command=self.save_pwd)
        self.enrollButton.grid(row=4,column=1,sticky=E, padx=5, pady=5)          
    def save_pwd(self):
        global password, email, pop3_server, smtp_server
        password = self.pwdEntry.get()
        email = self.emailEntry.get()
        pop3_server = self.popEntry.get()
        smtp_server = self.smtpEntry.get()
        #print(email, password)
        
class ReceiveApplication(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid(row=0,column=1,sticky=E+N)
        self.createWidgets()
    def createWidgets(self):
        create_bg(self)
        self.emailtopicLabel = Label(self, text='主题')
        self.emailtopicLabel.grid(row=0,column=1,sticky=E)
        self.emailtopicEntry = Entry(self, width=60)
        self.emailtopicEntry.grid(row=0,column=2,sticky=W)

        self.sendaddrLabel = Label(self, text='发件人')
        self.sendaddrLabel.grid(row=2,column=1,sticky=E)
        self.sendaddrEntry = Entry(self, width=60)
        self.sendaddrEntry.grid(row=2,column=2,sticky=W)

        self.timeLabel = Label(self, text='时间')
        self.timeLabel.grid(row=4,column=1,sticky=E)
        self.timeEntry = Entry(self, width=60)
        self.timeEntry.grid(row=4,column=2,sticky=W)

        self.receiveaddrLabel = Label(self, text='收件人')
        self.receiveaddrLabel.grid(row=6,column=1,sticky=E)
        self.receiveaddrEntry = Entry(self, width=60)
        self.receiveaddrEntry.grid(row=6,column=2,sticky=W)
        
        self.emailText = Text(self, width=60)
        self.emailText.grid(row=8, column=1,columnspan=10)
        
        self.receive_email()
    def receive_email(self):
        # 连接到POP3服务器:
        try:
            server = poplib.POP3(pop3_server)
        except BaseException:
            print('pop3 error')
            messagebox.showinfo("提示","POP3服务器错误，在登陆界面重新输入后单击确认")
            return
            
        # 可以打开或关闭调试信息:
        server.set_debuglevel(1)
        # 可选:打印POP3服务器的欢迎文字:
        print(server.getwelcome().decode('utf-8'))

        # 身份认证:
        try:
            server.user(email)
            server.pass_(password)
        except BaseException:
            print('email or password error')
            messagebox.showinfo("提示","邮箱账号或密码错误，在登陆界面重新输入后单击确认")
            return
        
        # stat()返回邮件数量和占用空间:
        print('Messages: %s. Size: %s' % server.stat())
        # list()返回所有邮件的编号:
        resp, mails, octets = server.list()
        # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
        print(mails)

        # 获取最新一封邮件, 注意索引号从1开始:
        index = len(mails)
        resp, lines, octets = server.retr(index)

        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        # 稍后解析出邮件:
        msg = Parser().parsestr(msg_content)
        self.print_info(msg)
        print('\n\n--------------------------')
        print(msg)
        # 可以根据邮件索引号直接从服务器删除邮件:
        # server.dele(index)
        # 关闭连接:
        server.quit()
    # indent用于缩进显示:
    def print_info(self, msg, indent=0):
        if indent == 0:
            for header in ['Date', 'From', 'To', 'Subject']:
                value = msg.get(header, '')
                if value:
                    if header=='Date':
                        self.timeEntry.insert(0, value)
                    elif header=='Subject':
                        value = decode_str(value)
                        self.emailtopicEntry.insert(0, value)
                    elif header=='To':
                        for v in value.split(','):
                            hdr, addr = parseaddr(v)
                            name = decode_str(hdr)
                            v = u'%s <%s>, ' % (name, addr)
                            #print(v)
                            self.receiveaddrEntry.insert(END, v)
                        self.receiveaddrEntry.delete(self.receiveaddrEntry.index(END) - 2, 
                            self.receiveaddrEntry.index(END))                       
                    else:
                        hdr, addr = parseaddr(value)
                        name = decode_str(hdr)
                        value = u'%s <%s>' % (name, addr)                       
                        self.sendaddrEntry.insert(0, value)
                            
                print('%s%s: %s' % ('  ' * indent, header, value))
        if (msg.is_multipart()):
            parts = msg.get_payload()
            for n, part in enumerate(parts):
                print('%spart %s' % ('  ' * indent, n))
                print('%s--------------------' % ('  ' * indent))
                self.print_info(part, indent + 1)
        else:
            content_type = msg.get_content_type()
            if content_type=='text/plain' or content_type=='text/html':
                content = msg.get_payload(decode=True)
                charset = guess_charset(msg)
                if charset:
                    content = content.decode(charset)
                print('%sText: %s' % ('  ' * indent, content))
                if content_type=='text/plain':
                    self.emailText.insert(INSERT, content)
            else:
                print('%sAttachment: %s' % ('  ' * indent, content_type))

class SendApplication(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid(row=0,column=1,sticky=E)
        self.createWidgets()
    def createWidgets(self):
        create_bg(self)

        self.emailaddrLabel = Label(self, text='收件人')
        self.emailaddrLabel.grid(row=0,column=1,sticky=E)
        self.emailaddrEntry = Entry(self, width=60)
        self.emailaddrEntry.grid(row=0,column=2,sticky=W)
        self.emailaddrEntry.insert(0, 'zhuchengdang@outlook.com')
        
        self.emailtopicLabel = Label(self, text='主题')
        self.emailtopicLabel.grid(row=2,column=1,sticky=E)
        self.emailtopicEntry = Entry(self, width=60)
        self.emailtopicEntry.grid(row=2,column=2,sticky=W)
              
        self.emailtxtLabel = Label(self, text='正文')
        self.emailtxtLabel.grid(row=4,column=1,sticky=E+N)        
        self.nameInput = Text(self,width=60)
        self.nameInput.grid(row=4,column=2)
        
        self.alertButton = Button(self, text='发送', command=self.post_email)
        self.alertButton.grid(row=6,column=1,sticky=E)

    def post_email(self):
        to_addr = self.emailaddrEntry.get()
        if len(email) ==0 or len(password) == 0:
            print('email or password null')
            messagebox.showinfo("提示","邮箱账号或密码为空，在登陆界面重新输入后单击确认")
            return        

        #正文以文本plain方式，或者可以用html方式
        msg = MIMEText(self.nameInput.get(0.0, END), 'plain', 'utf-8')
        msg['From'] = _format_addr('%s <%s>' % (email.split('@')[0], email))
        #如果邮件显示需要多人 + ','+ _format_addr('管理员 <493187455@qq.com>')
        msg['To'] = _format_addr('管理员 <%s>' % to_addr)
        msg['Subject'] = Header(self.emailtopicEntry.get(), 'utf-8').encode()       
        print(msg['Subject'])
        
        try:
            server = smtplib.SMTP(smtp_server, 25)#587? 25
            server.starttls()
        except BaseException:
            print('SMTP error')
            messagebox.showinfo("提示","SMTP错误，在登陆界面重新输入后单击确认")
            return        
        
        server.set_debuglevel(1)
        try:
            server.login(email, password)        
        except BaseException:
            print('email or password error')
            messagebox.showinfo("提示","邮箱账号或密码错误，在登陆界面重新输入后单击确认")
            return        
        #[to_addr] list结构，实际接收人
        server.sendmail(email, [to_addr], msg.as_string())
        server.quit()

app = MainApplication()
# 设置窗口标题:
app.master.title('Email post/receive')
# 主消息循环:
app.mainloop()