# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 21:05:37 2022

@author: REC3WX

vK8FXz+w~UxHR2L

"""


import sys
from datetime import datetime, timedelta
from PySide6.QtWidgets import *
from PySide6.QtGui import QFont
from PySide6.QtCore import Slot, Qt, QThread, Signal, QEvent
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr
import re
import os
import tempfile
from lxml import etree
from apscheduler.schedulers.background import BackgroundScheduler


class Worker(QThread):
    sinOut = Signal(str)

    def __init__(self, parent=None):
        super(Worker, self).__init__(parent)
        self.folder = f"{tempfile.gettempdir()}\GET\Schedule"
        if os.path.exists(self.folder):
            self.purge_file()
        else:
            os.makedirs(self.folder)

    def stop_self(self):
        self.terminate()
        message = f"进程已终止..."
        self.sinOut.emit(message)

    def purge_file(self):
        for xls in os.listdir(self.folder):
            os.remove(f"{self.folder}\\{xls}")

    def getdata(self, year, month, user, pwd, rec, chk_dld, once, timer, chk_workday):
        self.year = year
        self.month = month
        self.user = user
        self.pwd = pwd
        self.rec = rec
        self.chk_dld = chk_dld
        self.once = once
        self.timer = timer
        self.chk_workday = chk_workday

    def send_mail(self):
        mail_host = "smtp.163.com"
        mail_user = "vhcn_lop@163.com"
        mail_pass = "YQXQMPJDDTSOJLOC"
        # 邮件发送方邮箱地址
        sender = "vhcn_lop@163.com"
        # 邮件接受方邮箱地址，注意需要[]包裹，这意味着你可以写多个邮件地址群发
        receivers = self.rec
        # 设置email信息
        # 邮件内容设置
        email = MIMEMultipart()
        # 邮件主题
        time = datetime.now().isoformat(" ", "seconds")
        email["Subject"] = f"no-reply: {time} GTE内示自动下载结果"
        # 发送方信息
        email["From"] = formataddr(["Mail Bot@VHCN", sender])
        # 接受方信息
        email["To"] = receivers

        # add attachments
        file_list = [f for f in os.listdir(self.folder) if f.endswith(".xls")]
        for file in file_list:
            part = MIMEApplication(open(f"{self.folder}\\{file}", "rb").read())
            part.add_header("Content-Disposition", "attachment", filename=file)
            email.attach(part)
        # dynamic email content
        if not file_list and self.once == "1":
            message = f"单次执行无新内示不发送邮件!"
            self.sinOut.emit(message)
            return
        elif not file_list:
            # email_content = f"Dear all,\n\tGTE {self.ym} 没有新的内示, 请知悉! \n\nVHCN ICO"
            email_content = """\
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <title>GTE内示自动下载结果</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
</head>

<body style="margin: 0; padding: 0;">
  <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td style="padding: 20px 0 30px 0;">

        <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="border-collapse: collapse; border: 1px solid #cccccc;">
          <tr>
            <td align="center" bgcolor="#005691" style="padding: 20px 0 20px 0;">
                <h1 style="color: #ffffff; font-size: 24px; margin: 0;  font-family: Microsoft YaHei;">GTE内示自动下载结果</h1>
            </td>
          </tr>  
          <tr>
            <td bgcolor="#ffffff" style="padding: 30px 20px 30px 20px;">
              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                <tr>
                  <td align="center" style="color: #000000; font-family: Microsoft YaHei;">
                    <p style="margin: 0;">没有新的内示, 请知悉! </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td bgcolor="#005691" style="padding: 10px 10px;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                <tr>
                  <td align="center" style="color: #00000; font-family: Microsoft YaHei; font-size: 14px;">
                    <p style="color: #ffffff; margin: 0;">Powered by VHCN ICO</p>
                  </td>
                  <td align="right">
                    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse;">
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

      </td>
    </tr>
  </table>
</body>
</html>
"""
        else:
            pass
            email_content = """ \
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <title>GTE内示自动下载结果</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
</head>

<body style="margin: 0; padding: 0;">
  <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td style="padding: 20px 0 30px 0;">

        <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="border-collapse: collapse; border: 1px solid #cccccc;">
          <tr>
            <td align="center" bgcolor="#005691" style="padding: 20px 0 20px 0;">
                <h1 style="color: #ffffff; font-size: 24px; margin: 0;  font-family: Microsoft YaHei;">GTE内示自动下载结果</h1>
            </td>
          </tr>  
          <tr>
            <td bgcolor="#ffffff" style="padding: 30px 20px 30px 20px;">
              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                <tr>
                  <td align="center" style="color: #000000; font-family: Microsoft YaHei;">
                    <p style="margin: 0;">附件为新的内示, 请查收! </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td bgcolor="#005691" style="padding: 10px 10px;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                <tr>
                  <td align="center" style="color: #00000; font-family: Microsoft YaHei; font-size: 14px;">
                    <p style="color: #ffffff; margin: 0;">Powered by VHCN ICO</p>
                  </td>
                  <td align="right">
                    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse;">
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

      </td>
    </tr>
  </table>
</body>
</html>
"""
        # insert content
        email.attach(MIMEText(email_content, "html", "utf-8"))

        try:
            smtpObj = smtplib.SMTP_SSL(mail_host, 465)
            # 登录到服务器
            smtpObj.login(mail_user, mail_pass)
            # 发送
            smtpObj.sendmail(sender, receivers.split(";"), email.as_string())
            # 退出
            smtpObj.quit()
            message = f"邮件发送成功..."
            self.sinOut.emit(message)
        except smtplib.SMTPException as e:
            message = f"{e}"
            self.sinOut.emit(message)

    def to_unicode(self, supplier):
        ret = ""
        for v in supplier:
            ret = ret + hex(ord(v)).upper().replace("0X", "\\\\u")

        return ret

    def post_download(self):
        supplier = self.to_unicode(self.user[0:5])
        self.ym = self.year + self.month
        url = "http://192.168.10.33/WCFService/WcfService.svc"
        headers = {
            "content-type": "text/xml",
            "Referer": "http://192.168.10.33/ClientBin/SilverlightUI.xap",
            "SOAPAction": '"SysManager/WcfService/IPO0704GetInfo"',
        }
        if self.chk_dld == "1" or self.once == "1":
            scope = f'<parm>{{"SupplierNum":"{supplier}","InshowYM":"{self.ym}","Check":"{self.chk_dld}"}}</parm>'
        else:
            scope = (
                f'<parm>{{"SupplierNum":"{supplier}","Check":"{self.chk_dld}"}}</parm>'
            )

        body = (
            '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">'
            "<s:Body>"
            '<IPO0704GetInfo xmlns="SysManager">'
            '<pagerinfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            "<d4p1:pageIndex>1</d4p1:pageIndex>"
            "<d4p1:pageSize>10</d4p1:pageSize>"
            "</pagerinfo>"
            f"{scope}"
            "</IPO0704GetInfo>"
            "</s:Body>"
            "</s:Envelope>"
        )

        self.response = requests.post(url, data=body, headers=headers)
        self.resp_str = str(self.response.content.decode("utf8"))
        xml = etree.fromstring(self.resp_str)
        message = f"临时下载目录为: {self.folder}"
        self.sinOut.emit(message)
        for filenm in xml.iter("{*}FileNm"):
            try:
                message = f"命中: {filenm.text}"
                self.sinOut.emit(message)
                for ancestor in filenm.iterancestors("{*}IPO0704DTO"):
                    plant_ver = [
                        ancestor.find("./{*}Series").text,
                        ancestor.find("./{*}VersionNum").text,
                    ]
                    dld_path = f"http://192.168.10.33/GTESGFile/Inshow/{self.ym}/{plant_ver[0]}/{self.user[0:5]}/{filenm.text}"
                    # e.g.: http://192.168.10.33/GTESGFile/Inshow/202208/M1/3334A/XXXX.xls
                    message = f"下载Excel {filenm.text}"
                    self.sinOut.emit(message)
                    self.filepath = f"{self.folder}\\{filenm.text}"
                    dld_xls = requests.get(url=dld_path, stream=True)
                    with open(self.filepath, "wb") as xls:
                        xls.write(dld_xls.content)
                    message = f"下载完成! "
                    self.sinOut.emit(message)

            except Exception as e:
                message = f"{e}"
                self.sinOut.emit(message)
        else:
            message = f"没有找到新的内示..."
            self.sinOut.emit(message)

    def chain(self):
        try:
            if self.once == "0":
                message = f"{datetime.now()} - 开始运行定时任务..."
                self.sinOut.emit(message)
            else:
                message = f"{datetime.now()} - 开始运行单次任务..."
                self.sinOut.emit(message)

            message = f"获取Excel文件清单..."
            self.sinOut.emit(message)
            self.post_download()
            message = f"获取结束, 发送邮件..."
            self.sinOut.emit(message)
            self.send_mail()
            self.purge_file()
            if self.once == "0":
                message = f"{datetime.now()} - 定时任务已完成! "
                self.sinOut.emit(message)
                self.time_gap()
                message = f"下次任务时间为: {self.sch_dhm}, 距现在{self.gap_h}小时{self.gap_m}分"
                self.sinOut.emit(message)
            else:
                message = f"{datetime.now()} - 单次任务已完成! "
                self.sinOut.emit(message)

        except Exception as e:
            message = f"{e}"
            self.sinOut.emit(message)

    def stop_scheduler(self):
        message = f"正在终止定时任务..."
        self.sinOut.emit(message)
        try:
            self.scheduler.shutdown()
        except Exception as e:
            message = f"{e}"
            self.sinOut.emit(message)

    def time_gap(self):
        now = datetime.now()
        now_hm = now.strftime("%H:%M")
        now_d = now.strftime("%Y-%m-%d")
        now_dhm = now.strftime("%Y-%m-%d %H:%M")
        # sch_hm =input('输入时间(格式09:00): ')
        if now.weekday() == 4:
            delta = 3
        elif now.weekday() == 5:
            delta = 2
        else:
            delta = 1
        sch_hm = self.timer
        if self.chk_workday == "5":
            if now_hm < sch_hm:
                self.sch_dhm = f"{now_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                self.gap_h = gap.split(":")[0]
                self.gap_m = gap.split(":")[1]
            else:
                sch_d = datetime.strftime((now + timedelta(days=delta)), "%Y-%m-%d")
                self.sch_dhm = f"{sch_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                if "day" in str(gap):
                    self.gap_d = gap[0]
                    self.gap_h = int(gap.split(":")[0][-2:]) + 24 * int(self.gap_d)
                else:
                    self.gap_h = int(gap.split(":")[0])
                self.gap_m = gap.split(":")[1]
        else:
            if now_hm < sch_hm:
                self.sch_dhm = f"{now_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                self.gap_h = gap.split(":")[0]
                self.gap_m = gap.split(":")[1]
            else:
                sch_d = datetime.strftime((now + timedelta(days=delta)), "%Y-%m-%d")
                self.sch_dhm = f"{sch_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                if "day" in str(gap):
                    self.gap_d = gap[0]
                    self.gap_h = int(gap.split(":")[0][-2:]) + 24 * int(self.gap_d)
                else:
                    self.gap_h = gap.split(":")[0]
                self.gap_m = gap.split(":")[1]

    def run(self):
        # 主逻辑
        if self.once == "1":
            try:
                self.chain()
            except Exception as e:
                message = f"{e}"
                self.sinOut.emit(message)
        else:
            try:
                self.scheduler = BackgroundScheduler(timezone="Asia/Shanghai")
                h = self.timer[0:2]
                m = self.timer[3:5]
                if self.chk_workday == "7":
                    day_of_week = "*"
                else:
                    day_of_week = "mon-fri"
                try:
                    self.scheduler.add_job(
                        self.chain,
                        trigger="cron",
                        day_of_week=day_of_week,
                        hour=h,
                        minute=m,
                    )

                    self.scheduler.start()
                    if day_of_week == "*":
                        message = f"已设置定时任务为每天{self.timer} "
                    else:
                        message = f"已设置定时任务为每个工作日的{self.timer} "
                    self.sinOut.emit(message)
                    self.time_gap()
                    message = f"下次任务时间为: {self.sch_dhm}, 距现在{self.gap_h}小时{self.gap_m}分"
                    self.sinOut.emit(message)
                except Exception as e:
                    message = f"{e}"
                    self.sinOut.emit(message)
            except Exception as e:
                message = f"{e}"
                self.sinOut.emit(message)


class MyWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread = Worker()
        title = f"GTE内示查询下载工具 v1.0   - Made by REC3WX"
        self.setWindowTitle(title)
        pixmapi = QStyle.SP_FileDialogDetailedView
        icon = self.style().standardIcon(pixmapi)
        self.setWindowIcon(icon)
        self.setFixedSize(700, 300)

        self.tray = QSystemTrayIcon()
        self.tray.setIcon(icon)
        self.tray.setToolTip(f"{title} 正在后台运行")
        self.tray.activated.connect(self.on_systemTrayIcon_activated)

        self.fld_user = QLabel("用户名:")
        self.line_user = QLineEdit()
        self.line_user.setClearButtonEnabled(True)
        self.line_user.setText("3334A01")  ##测试
        self.fld_pwd = QLabel("密码:")
        self.line_pwd = QLineEdit()
        self.line_pwd.setClearButtonEnabled(True)
        self.line_pwd.setText("123456")  ##测试
        self.line_pwd.setEchoMode(QLineEdit.PasswordEchoOnEdit)

        self.fld_year = QLabel("年份:")
        self.cb_year = QComboBox()
        y = datetime.today().year
        self.cb_year.addItems([str(y - 1), str(y), str(y + 1)])
        self.cb_year.setCurrentText(str(y))
        self.cb_year.currentTextChanged[str].connect(self.get_year_month)

        self.fld_month = QLabel("月份:")
        self.cb_month = QComboBox()
        self.cb_month.addItems(
            ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
        )
        m = datetime.today().strftime("%m")
        self.cb_month.setCurrentText(str(m))
        self.cb_month.currentTextChanged[str].connect(self.get_year_month)

        self.calendar = QDateEdit()
        self.calendar.setCursor(Qt.PointingHandCursor)
        self.calendar.setDisplayFormat("yyyy-MM")

        self.chk_dld = QCheckBox("包含已下载", self)
        self.chk_workday = QCheckBox("仅工作日", self)

        self.btn_start = QPushButton("运行单次任务")
        self.btn_start.clicked.connect(self.execute_once)

        self.fld_sch = QLabel("定时任务时间:")
        self.time_sch = QTimeEdit()
        self.time_sch.setDisplayFormat("HH:mm")
        self.time_sch.setCursor(Qt.PointingHandCursor)

        self.fld_email = QLabel("收件人:")
        self.line_email = QLineEdit()
        self.line_email.setClearButtonEnabled(True)
        self.line_email.setPlaceholderText("多个收件人之间用分号;分开")
        self.line_email.setText(
            "chenlong.ren@cn.bosch.com;feng.he@cn.bosch.com;wenzhuo.gu@cn.bosch.com"
        )  ##测试
        self.line_email.setToolTip(self.line_email.text())
        self.line_email.editingFinished.connect(self.check_email)

        self.fld_result = QLabel("运行日志:")
        self.text_result = QPlainTextEdit()
        self.text_result.setReadOnly(True)

        self.line = QFrame()
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)

        self.btn_schedule = QPushButton("设置定时任务")
        self.btn_schedule.clicked.connect(self.set_schedule)
        self.btn_cnl_sch = QPushButton("取消定时任务")
        self.btn_cnl_sch.clicked.connect(self.cancel_schedule)
        self.btn_cnl_sch.setEnabled(False)

        self.btn_reset = QPushButton("清空日志")
        self.btn_reset.clicked.connect(self.reset_log)
        self.btn_stop = QPushButton("终止任务进程")
        self.btn_stop.setEnabled(False)
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
        self.layout.addWidget((self.chk_workday), 1, 7)

        self.layout.addWidget((self.line), 2, 0, 1, 8)
        self.layout.addWidget((self.fld_result), 3, 0)
        self.layout.addWidget((self.text_result), 4, 0, 6, 6)
        self.layout.addWidget((self.btn_reset), 4, 6, 1, 2)
        self.layout.addWidget((self.btn_start), 6, 6, 1, 2)
        # self.layout.addWidget((self.btn_stop), 7, 6, 1, 2)
        self.layout.addWidget((self.btn_schedule), 8, 6, 1, 2)
        self.layout.addWidget((self.btn_cnl_sch), 9, 6, 1, 2)

        self.setLayout(self.layout)

        self.thread.sinOut.connect(self.Addmsg)  # 解决重复emit

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
        result = QMessageBox.question(
            self, "警告", "是否确认退出? ", QMessageBox.Yes | QMessageBox.No
        )
        if result == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def check_email(self):
        if not self.line_email.text() == "":
            email_list = self.line_email.text().split(";")
            for email in email_list:
                if not re.match(
                    r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", email
                ):
                    self.msgbox("error", "邮箱格式错误, 请重新输入 !")
                    self.line_email.setText("")
                    self.line_email.setFocus()
        self.line_email.setToolTip(self.line_email.text())

    def Addmsg(self, message):
        self.text_result.appendPlainText(message)

    def stop_thread(self):
        confirm = QMessageBox.question(
            self, "警告", "是否终止进程? ", QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            self.thread.stop_self()
            self.btn_stop.setEnabled(False)

    def get_year_month(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        title = "GTE内示查询下载工具 v1.0   - Made by REC3WX"
        if not year == "" and not month == "":
            # self.text_result.appendPlainText(f'当前选择的发票年月为: {year}年{month}月')
            self.setWindowTitle(f"{title} - 内示年月: {year}-{month}")

    def set_schedule(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        user = str(self.line_user.text())
        pwd = str(self.line_pwd.text())
        rec = str(self.line_email.text())
        timer = str(self.time_sch.text())
        if self.chk_dld.isChecked():
            chk_dld = "1"
        else:
            chk_dld = "0"

        if self.chk_workday.isChecked():
            chk_workday = "5"
        else:
            chk_workday = "7"

        if user == "" or pwd == "":
            self.msgbox("error", "请输入用户名和密码!! ")
        else:
            if not rec == "":
                confirm = QMessageBox.question(
                    self,
                    "确认",
                    f"是否设置定时任务: 每天{timer}? ",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if confirm == QMessageBox.Yes:
                    self.btn_cnl_sch.setEnabled(True)
                    self.btn_schedule.setEnabled(False)
                    self.thread.getdata(
                        year, month, user, pwd, rec, chk_dld, "0", timer, chk_workday
                    )
                    self.thread.start()
            else:
                self.msgbox("error", "请输入邮箱地址!! ")

    def cancel_schedule(self):
        confirm = QMessageBox.question(
            self, "警告", "是否取消定时任务? ", QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            self.thread.stop_scheduler()
            self.Addmsg(f"定时任务已取消")
            self.btn_cnl_sch.setEnabled(False)
            self.btn_schedule.setEnabled(True)

    def execute_once(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        user = str(self.line_user.text())
        pwd = str(self.line_pwd.text())
        rec = str(self.line_email.text())
        if self.chk_dld.isChecked():
            chk_dld = "1"
        else:
            chk_dld = "0"
        if self.chk_workday.isChecked():
            chk_workday = "5"
        else:
            chk_workday = "7"
        if user == "" or pwd == "":
            self.msgbox("error", "请输入用户名和密码!! ")
        else:
            if not rec == "":
                self.btn_stop.setEnabled(True)
                self.thread.getdata(year, month, user, pwd, rec, chk_dld, "1", "", "")
                self.thread.start()
            else:
                self.msgbox("error", "请输入邮箱地址!! ")

    def msgbox(self, title, text):
        tip = QMessageBox(self)
        if title == "error":
            tip.setIcon(QMessageBox.Critical)
        elif title == "DONE":
            tip.setIcon(QMessageBox.Warning)
        tip.setWindowFlag(Qt.FramelessWindowHint)
        font = QFont()
        font.setFamily("Microsoft YaHei")
        font.setPointSize(9)
        tip.setFont(font)
        tip.setText(text)
        tip.exec()

    def reset_log(self):
        confirm = QMessageBox.question(
            self, "警告", "是否清空日志? ", QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            self.text_result.clear()


def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    widget = MyWidget()
    # app.setStyleSheet(qdarktheme.load_stylesheet(border="rounded"))
    app.setStyle("fusion")

    widget.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
