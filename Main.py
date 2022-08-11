# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 21:05:37 2022

@author: REC3WX

vK8FXz+w~UxHR2L

"""


from re import M
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft  import EdgeChromiumDriverManager
import pandas as pd
import sys
import time
from datetime import datetime
from PySide6.QtWidgets import *
from PySide6.QtGui import QFont, QAction
from PySide6.QtCore import Slot, Qt, QThread, Signal, QEvent
import qdarktheme
import requests
import smtplib
from email.mime.text import MIMEText
import re

class Worker(QThread):
  sinOut = Signal(str)

  def __init__(self, parent=None):
    super(Worker, self).__init__(parent)

  def getdata(self, year, month, user, pwd, email, chk_dld, once):
    self.year = year
    self.month = month
    self.user = user
    self.pwd = pwd
    self.email = email
    self.chk_dld = chk_dld
    self.once = once
  
  def send_mail(self):
    mail_host = 'smtp.163.com'
    mail_user = 'vhcn_lop@163.com'
    mail_pass = 'vK8FXz+w~UxHR2L'   
    #邮件发送方邮箱地址
    sender = 'vhcn_lop@163.com'  
    #邮件接受方邮箱地址，注意需要[]包裹，这意味着你可以写多个邮件地址群发
    receivers = self.email

    #设置email信息
    #邮件内容设置
    message = MIMEText('content','plain','utf-8')
    #邮件主题  
    message['Subject'] = f'NO-REPLY: {}内示下载结果' 
    #发送方信息
    message['From'] = sender 
    #接受方信息     
    message['To'] = receivers[0] 
    try:
        smtpObj = smtplib.SMTP() 
        #连接到服务器
        smtpObj.connect(mail_host,25)
        #登录到服务器
        smtpObj.login(mail_user,mail_pass) 
        #发送
        smtpObj.sendmail(
            sender,receivers,message.as_string()) 
        #退出
        smtpObj.quit() 
        print('success')
    except smtplib.SMTPException as e:
        print('error',e) #打印错误
  def run(self):
    #主逻辑
    if self.once == '0':
        try:
            message = f'{datetime.now()} - 获取未下载文件清单...'
            self.sinOut.emit(message)

            message = 'DONE'
            self.sinOut.emit(message)

        except Exception:
            message = f'{datetime.now()} - 获取失败, 请确认VPN连接是否正常! '
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
        self.fld_pwd = QLabel('密码:')
        self.line_pwd = QLineEdit()
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
        m = datetime.today().month
        self.cb_month.setCurrentText(str(m))
        self.cb_month.currentTextChanged[str].connect(self.get_year_month)

        self.chk_dld = QCheckBox('考虑已下载', self)

        self.btn_start = QPushButton('开始单个任务')
        self.btn_start.clicked.connect(self.execute_once)

        self.fld_sch = QLabel('定时任务时间:')
        self.time_sch = QTimeEdit()
        
        self.fld_email = QLabel('收件邮箱地址:')
        self.line_email = QLineEdit()
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

        self.layout = QGridLayout()
        self.layout.addWidget((self.fld_user), 0, 0)
        self.layout.addWidget((self.line_user), 0, 1)
        self.layout.addWidget((self.fld_pwd), 0, 2)
        self.layout.addWidget((self.line_pwd), 0, 3)        
        self.layout.addWidget((self.fld_year), 0, 4)
        self.layout.addWidget((self.cb_year), 0, 5)
        self.layout.addWidget((self.fld_month), 0, 6)
        self.layout.addWidget((self.cb_month), 0, 7)
        self.layout.addWidget((self.fld_email), 1, 0)
        self.layout.addWidget((self.line_email), 1, 1, 1, 3)
        self.layout.addWidget((self.fld_sch), 1, 4)
        self.layout.addWidget((self.time_sch), 1, 5)
        self.layout.addWidget((self.chk_dld), 1, 7)


        self.layout.addWidget((self.line), 2, 0, 1, 8)
        self.layout.addWidget((self.fld_result), 3, 0)
        self.layout.addWidget((self.text_result), 4, 0, 5, 6)
        self.layout.addWidget((self.btn_start), 5, 6, 1, 2)
        self.layout.addWidget((self.btn_schedule), 6, 6, 1, 2)
        self.layout.addWidget((self.btn_reset), 7, 6, 1, 2)

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
            if not re.match(r'^[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+){0,4}@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+){0,4}$', self.line_email.text()):
                self.msgbox('error', '邮箱格式错误, 请重新输入 !')
                self.line_email.setText('')
                self.line_email.setFocus() 

    def Addmsg(self, message):
        if not message == 'DONE':
            self.text_result.appendPlainText(message)
        else:
            self.msgbox('DONE', '录入完成, 请确认后保存!! ')

    def get_year_month(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        title = 'GTE内示查询下载工具 v0.1   - Made by REC3WX'
        if not year == '' and not month == '':
            #self.text_result.appendPlainText(f'当前选择的发票年月为: {year}年{month}月')
            self.setWindowTitle(f'{title} - 内示年月: {year}-{month}')
            
    def set_schedule(self):
        self.line_user.setText('')
        self.line_pwd.setText('')
        self.cb_year.setCurrentText('')
        self.cb_month.setCurrentText('')
        self.text_result.clear()
        self.btn_start.setEnabled(False)

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

    def execute_once(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        user = str(self.line_user.text())
        pwd = str(self.line_pwd.text())
        email = str(self.line_email.text())
        if self.chk_dld.isChecked():
            chk_dld = '1'
        else:
            chk_dld = '0'      
        if user == '' or user == '':
            self.msgbox('error', '请输入用户名和密码!! ')
        else:
            self.thread.getdata(year, month, user, pwd, email, chk_dld, '0')
            self.thread.start()

def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    #app.setStyleSheet(qdarktheme.load_stylesheet(border="rounded"))
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    widget = MyWidget()
    widget.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()