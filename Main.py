# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 21:05:37 2022

@author: REC3WX

Tool to mass input the VAT on GTE Tax Recon website
CMD:
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\EdgeProfile"
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\ChromeProfile"

ChromeDriverManager-chrome.py
            url: str = "http://npm.taobao.org/mirrors/chromedriver",
            latest_release_url: str = "http://npm.taobao.org/mirrors/chromedriver/LATEST_RELEASE",

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
#from PySide6.QtWidgets import (QWidget, QPushButton, QFileDialog, QApplication, QLineEdit, QGridLayout, QLabel, QMessageBox, QPlainTextEdit, QFrame, QStyle, QComboBox, QSystemTrayIcon)
from PySide6.QtWidgets import *
from PySide6.QtGui import QFont, QAction
from PySide6.QtCore import Slot, Qt, QThread, Signal, QEvent
import qdarktheme
import socket

opt = Options()
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
port = ("localhost", 9333)
open = sock.connect_ex(port)
if open == 0:
    opt.add_experimental_option("debuggerAddress", "localhost:9333")
else:
    opt.add_argument("--remote-debugging-port=9333")
    opt.add_argument("--start-maximized")
    opt.add_argument('user-data-dir=C:\\selenium\\EdgeProfile')
##driver_path = ChromeService(r'./chromedriver.exe')
##driver = webdriver.Chrome(service=driver_path, options=opt)
driver = webdriver.ChromiumEdge(service=EdgeService(EdgeChromiumDriverManager().install()), options=opt)
driver.set_page_load_timeout(5)

class Worker(QThread):
  sinOut = Signal(str)

  def __init__(self, parent=None):
    super(Worker, self).__init__(parent)

  def getdata(self, year, month, user, pwd):
    self.year = year
    self.month = month
    self.user = user
    self.pwd = pwd
  
  def run(self):
    #主逻辑
    try:
        message = f'{datetime.now()} - 打开网站并自动登录...'
        self.sinOut.emit(message)
        
        message = f'{datetime.now()} - 开始录入发票...'
        self.sinOut.emit(message)
        iframe = driver.find_element(By.XPATH, '//iframe[@id="layui-layer-iframe1"]')
        driver.switch_to.frame("layui-layer-iframe1")

        
        self.sinOut.emit(message)
        message = 'DONE'
        self.sinOut.emit(message)

    except Exception:
        driver.execute_script('window.stop()')
        message = f'{datetime.now()} - 网页无法加载, 请确认VPN连接是否正常! '
        self.sinOut.emit(message)
        #driver.quit()

class MyWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread = Worker()
        title = 'GTE内示查询下载工具 v0.1   - Made by REC3WX'
        self.setWindowTitle(title)
        pixmapi = QStyle.SP_FileDialogDetailedView
        icon = self.style().standardIcon(pixmapi)
        self.setWindowIcon(icon)
        self.setFixedSize(700, 300)
        
        self.tray = QSystemTrayIcon()
        self.tray.setIcon(icon)
        self.tray.setToolTip(f'{title} 正在后台运行')
        self.tray.activated.connect(self.on_systemTrayIcon_activated)
        
        self.fld_user= QLabel('用户名:')
        self.line_user = QLineEdit()
        self.fld_pwd= QLabel('密码:')
        self.line_pwd = QLineEdit()
        self.line_pwd.setEchoMode(QLineEdit.PasswordEchoOnEdit)
        
        self.fld_year= QLabel('年份:')
        self.cb_year= QComboBox()
        y = datetime.today().year
        self.cb_year.addItems(['', str(y-1), str(y), str(y+1)])
        self.cb_year.setCurrentText(str(y))
        self.cb_year.currentTextChanged[str].connect(self.get_year_month)

        self.fld_month= QLabel('月份:')
        self.cb_month = QComboBox()
        self.cb_month.addItems(['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'])
        m = datetime.today().month
        self.cb_month.setCurrentText(str(m))
        self.cb_month.currentTextChanged[str].connect(self.get_year_month)

        self.btn_start = QPushButton('开始')
        self.btn_start.clicked.connect(self.execute)

        self.fld_sch = QLabel('定时任务:')
        self.timer = QTimeEdit()
                
        self.fld_result = QLabel('运行日志:')
        self.text_result = QPlainTextEdit()
        self.text_result.setReadOnly(True)

        self.line = QFrame()
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)

        self.btn_reset = QPushButton('设置定时任务')
        self.btn_reset.clicked.connect(self.schedule)

        self.layout = QGridLayout()
        self.layout.addWidget((self.fld_user), 0, 0)
        self.layout.addWidget((self.line_user), 0, 1)
        self.layout.addWidget((self.fld_pwd), 0, 2)
        self.layout.addWidget((self.line_pwd), 0, 3)        
        self.layout.addWidget((self.fld_year), 0, 4)
        self.layout.addWidget((self.cb_year), 0, 5)
        self.layout.addWidget((self.fld_month), 0, 6)
        self.layout.addWidget((self.cb_month), 0, 7)
        self.layout.addWidget((self.fld_sch), 1, 4)
        self.layout.addWidget((self.timer), 1, 5)


        self.layout.addWidget((self.line), 2, 0, 1, 8)
        self.layout.addWidget((self.fld_result), 3, 0)
        self.layout.addWidget((self.text_result), 4, 0, 5, 6)
        self.layout.addWidget((self.btn_start), 5, 6, 1, 2)
        self.layout.addWidget((self.btn_reset), 6, 6, 1, 2)

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
            
    def schedule(self):
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

    def execute(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        user = str(self.line_user.text())
        pwd = str(self.line_pwd.text())
        if user == '' or user == '':
            self.msgbox('error', '请输入用户名和密码!! ')
        else:
            self.thread.getdata(year, month, user, pwd)
            self.thread.start()

def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    app.setStyleSheet(qdarktheme.load_stylesheet())
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    widget = MyWidget()
    widget.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()