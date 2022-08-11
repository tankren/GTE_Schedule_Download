# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 21:05:37 2022

@author: REC3WX

vK8FXz+w~UxHR2L

"""


import sys
import time
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtGui import QFont, QAction
from PySide6.QtCore import Slot, Qt, QThread, Signal, QEvent
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import re
import os

class Worker(QThread):
  sinOut = Signal(str)

  def __init__(self, parent=None):
    super(Worker, self).__init__(parent)

  def stop_self(self):
        self.terminate()
        message = f'{datetime.now()} - 进程已终止...'
        self.sinOut.emit(message)
        

  def getdata(self, year, month, user, pwd, rec, chk_dld, once):
    self.year = year
    self.month = month
    self.user = user
    self.pwd = pwd
    self.rec = rec
    self.chk_dld = chk_dld
    self.once = once
  
  def send_mail(self):
    mail_host = 'smtp.163.com'
    mail_user = 'vhcn_lop@163.com'
    mail_pass = 'YQXQMPJDDTSOJLOC'   
    #邮件发送方邮箱地址
    sender = 'vhcn_lop@163.com'  
    #邮件接受方邮箱地址，注意需要[]包裹，这意味着你可以写多个邮件地址群发
    receivers = self.rec
    #设置email信息
    #邮件内容设置
    email = MIMEMultipart()
    email_content = "Dear all,\n\t附件为每周数据，请查收！\n\nBI团队"
    email.attach(MIMEText('email_content ', 'plain', 'utf-8'))
    #邮件主题  
    time = datetime.now().isoformat(' ', 'seconds')
    email['Subject'] = f'no-reply: {time} GTE内示下载结果' 
    #发送方信息
    email['From'] = sender 
    #接受方信息     
    email['To'] = receivers
    folder = 'c:\\temp'
    file_list = [f for f in os.listdir(folder) if f.endswith('.xlsx')]
    for file in file_list:
        part = MIMEApplication(open(f'{folder}\\{file}', 'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename=file)
        email.attach(part)
    try:
        smtpObj = smtplib.SMTP_SSL(mail_host, 465)
        #登录到服务器
        smtpObj.login(mail_user,mail_pass) 
        #发送
        smtpObj.sendmail(
            sender,receivers.split(';'),email.as_string()) 
        #退出
        smtpObj.quit() 
        message = f'{datetime.now()} - 邮件发送成功...'
        self.sinOut.emit(message)
    except smtplib.SMTPException as e:
        print('错误',e) #打印错误

  def to_unicode(self, supplier):
    ret = ''
    for v in supplier:
        ret = ret + hex(ord(v)).upper().replace('0X', '\\u')

    return ret

  def POST(self):
        supplier = self.to_unicode(self.user[0:5])
        ym = self.year+self.month
        url = 'http://192.168.10.33/slcontainer/web.csv'
        body = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">' \
                    '<s:Body>' \
                    '<IPO0704GetInfo xmlns="SysManager">' \
                    '<pagerinfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">' \
                    '<d4p1:pageIndex>1</d4p1:pageIndex>' \
                    '<d4p1:pageSize>10</d4p1:pageSize>' \
                    '</pagerinfo>' \
                    f'<parm>{{"SupplierNum":"{supplier}","InshowYM":"{ym}","Check":"{self.chk_dld}"}}</parm>' \
                    '</IPO0704GetInfo>' \
                    '</s:Body>' \
                    '</s:Envelope>' 
        print(body.encode("utf-8"))            
        r = requests.post(url, data=body.encode("utf-8"))
        print(r.text)

  def download(self):
        pass

  def run(self):
    #主逻辑
    if self.once == '1':
        try:
            message = f'{datetime.now()} - 获取Excel文件清单...'
            self.sinOut.emit(message)

            self.POST()
            message = f'{datetime.now()} - 获取结束, 发送邮件...'
            self.sinOut.emit(message)
            self.send_mail()
        except Exception as e:
            message = f'{datetime.now()} - {e} '
            self.sinOut.emit(message)
    else:
        try:
            message = f'{datetime.now()} - 设置定时任务...'       
        except Exception:
            message = f'{datetime.now()} - 错误...'  

class MyWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread = Worker()
        title = f'GTE内示查询下载工具 v0.1   - Made by REC3WX'
        self.setWindowTitle(title)
        pixmapi = QStyle.SP_FileDialogDetailedView
        icon = self.style().standardIcon(pixmapi)
        self.setWindowIcon(icon)
        self.setFixedSize(700, 300)
        
        self.tray = QSystemTrayIcon()
        self.tray.setIcon(icon)
        self.tray.setToolTip(f'{title} 正在后台运行')
        self.tray.activated.connect(self.on_systemTrayIcon_activated)
        
        self.fld_user = QLabel('用户名:')
        self.line_user = QLineEdit()
        self.line_user.setClearButtonEnabled(True)
        self.line_user.setText("3334A01")  ##测试
        self.fld_pwd = QLabel('密码:')
        self.line_pwd = QLineEdit()
        self.line_pwd.setClearButtonEnabled(True)
        self.line_pwd.setText("123456")  ##测试
        self.line_pwd.setEchoMode(QLineEdit.PasswordEchoOnEdit)
        
        self.fld_year = QLabel('年份:')
        self.cb_year = QComboBox()
        y = datetime.today().year
        self.cb_year.addItems([str(y-1), str(y), str(y+1)])
        self.cb_year.setCurrentText(str(y))
        self.cb_year.currentTextChanged[str].connect(self.get_year_month)

        self.fld_month = QLabel('月份:')
        self.cb_month = QComboBox()
        self.cb_month.addItems(['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'])
        m = datetime.today().strftime('%m')
        self.cb_month.setCurrentText(str(m))
        self.cb_month.currentTextChanged[str].connect(self.get_year_month)

        self.chk_dld = QCheckBox('考虑已下载', self)

        self.btn_start = QPushButton('开始单个任务')
        self.btn_start.clicked.connect(self.execute_once)

        self.fld_sch = QLabel('定时任务时间:')
        self.time_sch = QTimeEdit()
        
        self.fld_email = QLabel('收件人:')
        self.line_email = QLineEdit()
        self.line_email.setClearButtonEnabled(True)
        self.line_email.setPlaceholderText("多个收件人之间用分号;分开") 
        self.line_email.setText("tankren@qq.com;tankrenlive@gmail.com")  ##测试
        self.line_email.editingFinished.connect(self.check_email)
        
        self.fld_result = QLabel('运行日志:')
        self.text_result = QPlainTextEdit()
        self.text_result.setReadOnly(True)

        self.line = QFrame()
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)

        self.btn_schedule = QPushButton('设置定时任务')
        self.btn_schedule.clicked.connect(self.set_schedule)

        self.btn_reset = QPushButton('清空日志')
        self.btn_reset.clicked.connect(self.reset_log)
        self.btn_stop = QPushButton('终止进程')
        self.btn_stop.clicked.connect(self.stop_thread)

        self.layout = QGridLayout()
        self.layout.addWidget((self.fld_user), 0, 0)
        self.layout.addWidget((self.line_user), 0, 1)
        self.layout.addWidget((self.fld_pwd), 0, 2)
        self.layout.addWidget((self.line_pwd), 0, 3)        
        self.layout.addWidget((self.fld_year), 0, 4)
        self.layout.addWidget((self.cb_year), 0, 5)
        self.layout.addWidget((self.fld_month), 1, 4)
        self.layout.addWidget((self.cb_month), 1, 5)
        self.layout.addWidget((self.fld_email), 1, 0)
        self.layout.addWidget((self.line_email), 1, 1, 1, 3)
        self.layout.addWidget((self.fld_sch), 0, 6)
        self.layout.addWidget((self.time_sch), 0, 7)
        self.layout.addWidget((self.chk_dld), 1, 6)


        self.layout.addWidget((self.line), 2, 0, 1, 8)
        self.layout.addWidget((self.fld_result), 3, 0)
        self.layout.addWidget((self.text_result), 4, 0, 6, 6)
        self.layout.addWidget((self.btn_start), 5, 6, 1, 2)
        self.layout.addWidget((self.btn_schedule), 6, 6, 1, 2)
        self.layout.addWidget((self.btn_reset), 7, 6, 1, 2)
        self.layout.addWidget((self.btn_stop), 8, 6, 1, 2)

        self.setLayout(self.layout)
        
        self.thread.sinOut.connect(self.Addmsg)  #解决重复emit

    @Slot()
    def on_systemTrayIcon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            if self.isHidden():
                self.show()
                self.activateWindow()
                self.tray.hide()
                
    def changeEvent(self, event):
        if event.type() == QEvent.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                event.ignore()
                self.hide()
                self.tray.show()
                
    def closeEvent(self, event):
        result = QMessageBox.question(self, "警告", "是否确认退出? ", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            event.accept()
        else:
            event.ignore()

    def check_email(self):
        if not self.line_email.text() == '':
            email_list = self.line_email.text().split(';')
            for email in email_list:
                if not re.match(r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", email):
                    self.msgbox('error', '邮箱格式错误, 请重新输入 !')
                    self.line_email.setText('')
                    self.line_email.setFocus() 

    def Addmsg(self, message):
        self.text_result.appendPlainText(message)

    def stop_thread(self):
        confirm = QMessageBox.question(self, "警告", "是否终止进程? ", QMessageBox.Yes | QMessageBox.No)
        if(confirm == QMessageBox.Yes):
            self.thread.stop_self()

    def get_year_month(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        title = 'GTE内示查询下载工具 v0.1   - Made by REC3WX'
        if not year == '' and not month == '':
            #self.text_result.appendPlainText(f'当前选择的发票年月为: {year}年{month}月')
            self.setWindowTitle(f'{title} - 内示年月: {year}-{month}')
            
    def set_schedule(self):
        print()
    
    def execute_once(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        user = str(self.line_user.text())
        pwd = str(self.line_pwd.text())
        rec = str(self.line_email.text())
        if self.chk_dld.isChecked():
            chk_dld = '1'
        else:
            chk_dld = '0'      
        if user == '' or pwd == '':
            self.msgbox('error', '请输入用户名和密码!! ')
        else:
            if not rec == '':
                self.thread.getdata(year, month, user, pwd, rec, chk_dld, '1')
                self.thread.start()
            else:
                self.msgbox('error', '请输入邮箱地址!! ')


    def msgbox(self, title, text):
        tip = QMessageBox(self)
        if title == 'error':
            tip.setIcon(QMessageBox.Critical)
        elif title == 'DONE' :
            tip.setIcon(QMessageBox.Warning)
        tip.setWindowFlag(Qt.FramelessWindowHint)
        font = QFont()
        font.setFamily("Microsoft YaHei")
        font.setPointSize(9)
        tip.setFont(font)
        tip.setText(text)
        tip.exec()  

    def reset_log(self):
        confirm = QMessageBox.question(self, "警告", "是否清空日志? ", QMessageBox.Yes | QMessageBox.No)
        if(confirm == QMessageBox.Yes):
            self.text_result.clear()          

def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    #app.setStyleSheet(qdarktheme.load_stylesheet(border="rounded"))
    app.setStyle("fusion")
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    widget = MyWidget()
    widget.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()