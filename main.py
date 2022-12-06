import os
import sys

from pathlib import Path
import sqlite3

from functools import partial
import winreg

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QTextCursor, QMovie
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QDialog, QMessageBox, QTableWidgetItem

import about_python_ui
import edit_ui
import help_ui
import osc_value
import osc_value_ui

sqlite3_file_path = 'user/Config.db'
config_file_path = 'user/Config.ini'
gif_file_path = 'user/images/image_gif.gif'

open_path = ''


def get_desktop():  # 获取桌面路径
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]


def pushbutton_open_click_success(self):  # 获取图像路径
    global open_path

    if open_path == '':
        open_path = get_desktop()  # 获取桌面目录

    image_path_qt = QFileDialog.getExistingDirectory(None, "请选择图像文件夹", open_path)  # 打开选择图像目录
    if image_path_qt != '':
        self.lineEdit_path.setText(image_path_qt)  # 更新输入框目录
        open_path = image_path_qt


def pushbutton_open_2_click_success(self):  # 获取标准路径
    global open_path

    if open_path == '':
        open_path = get_desktop()  # 获取桌面目录

    def_excel_path_qt = QFileDialog.getOpenFileName(None, '请选择标准文件', open_path, 'Excel文件 (*.xlsx *.xls)')  # 打开选择标准文件目录
    def_excel_path_qt = list(def_excel_path_qt)  # 将元组转换成列表
    if def_excel_path_qt[0] != '':
        self.lineEdit_path_2.setText(def_excel_path_qt[0])  # 更新输入框路径
        open_path = def_excel_path_qt[0][:def_excel_path_qt[0].rfind('/')]  # 获取对应的目录


def pushbutton_generate_click_success(self):  # 点击开始按钮，获取配置，子线程
    image_path_qt = self.lineEdit_path.text()  # 获取输入的图像目录
    def_excel_path_qt = self.lineEdit_path_2.text()  # 获取输入的标准文件路径

    if Path(image_path_qt).is_dir() and Path(def_excel_path_qt).is_file():  # 目录正确，文件路径正确
        self.pushButton_generate.setEnabled(False)  # 禁用按钮
        self.pushButton_open.setEnabled(False)  # 禁用按钮
        self.pushButton_open_2.setEnabled(False)  # 禁用按钮

        osc_value.image_path = image_path_qt  # 图像对应目录
        osc_value.def_excel_path = def_excel_path_qt  # 标准文件对应路径
        osc_value.program_path = os.path.dirname(os.path.realpath(sys.argv[0]))  # 获取程序运行路径

        config_res, config_ini = get_config_data()  # 获取配置数据，截图区域
        config_temp = []
        config_get = []
        for i in range(6):
            for j in range(4):
                config_temp.append(config_res[i][j + 1])  # 去掉名
            config_get.append(tuple(config_temp.copy()))  # 复制中间变量数据转换成元组
            config_temp.clear()  # 清空中间变量
        osc_value.config_get = config_get.copy()  # 更新进程全局变量

        osc_value.negative_film = config_ini[1]  # 图像是否需要反相，默认需要
        osc_value.image_matching = config_ini[2]  # 是否需要采用图像匹配方式，默认需要
        osc_value.fast_recognition = config_ini[3]  # 是否需要采用快速识别，默认不需要
        osc_value.super_fast_recognition = config_ini[4]  # 是否需要采用超快速识别，默认需要

        thread_start(self)  # 启动进程
    else:
        ui_up_to_data(self, -1, 'The path of picture directory or excel file is wrong!\n')


def actionabout_python3_click_success():  # 帮助python3页面
    # 因为使用Qt Designer设计的ui是默认继承自object类，不提供show()显示方法，
    # 所以我们需要生成一个QDialog对象来重载我们设计的Ui_Dialog类，从而达到显示效果。
    MainDialog = QDialog()  # 创建一个主窗体（必须要有一个主窗体）
    myDialog = about_python_ui.Ui_Dialog()  # 创建对话框
    myDialog.setupUi(MainDialog)  # 将对话框依附于主窗体
    # 设置窗口的属性为ApplicationModal模态，用户只有关闭弹窗后，才能关闭主界面
    MainDialog.setWindowModality(Qt.ApplicationModal)
    myDialog.pushButton.clicked.connect(MainDialog.close)  # 按钮点击事件信号链接关闭窗口
    MainDialog.show()  # 显示窗口
    MainDialog.exec_()


def actionabout_help_click_success():  # 帮助页面
    MainDialog = QDialog()  # 创建一个主窗体（必须要有一个主窗体）
    myDialog = help_ui.Ui_Dialog()  # 创建对话框
    myDialog.setupUi(MainDialog)  # 将对话框依附于主窗体
    MainDialog.setWindowModality(Qt.ApplicationModal)
    myDialog.pushButton.clicked.connect(MainDialog.close)  # 按钮点击事件信号链接关闭窗口
    MainDialog.show()  # 显示窗口
    MainDialog.exec_()


def get_config_data():  # 获取保存的配置信息
    conn = sqlite3.connect(sqlite3_file_path)  # 创建/链接一个 Connection 对象，它代表数据库
    c = conn.cursor()  # 创建一个 Cursor 游标对象
    c.execute('SELECT * FROM Config')  # 选择Config表
    res = c.fetchall()  # 获取全部内容
    conn.close()  # 关闭数据库链接

    config_res = []
    for s in res:
        config_res.append(list(s))  # 转换成列表

    fr = open(config_file_path, 'r')
    config = fr.readlines()
    fr.close()

    config_ini = [bool(int(config[0].replace('tableWidget_qt=', ''))),
                  bool(int(config[1].replace('negative_film_qt=', ''))),
                  bool(int(config[2].replace('image_matching_qt=', ''))),
                  bool(int(config[3].replace('fast_recognition_qt=', ''))),
                  bool(int(config[4].replace('super_fast_recognition_qt=', '')))]

    return config_res, config_ini


def accept(self, self2):  # 更新保存配置信息
    config_get = []
    config_temp = []
    config_name = ['crop_TK_jpg',
                   'crop_TK_png',
                   'crop_RS_old_jpg',
                   'crop_RS_old_png',
                   'crop_RS_new_jpg',
                   'crop_RS_new_png']

    for i in range(6):
        for j in range(4):
            config_temp.append(int(self.tableWidget.item(i, j).text()))  # 获取表中一行数据
        config_temp.append(config_name[i])  # 组装示波器名
        config_get.append(tuple(config_temp.copy()))  # 转换成元组
        config_temp.clear()

    conn = sqlite3.connect(sqlite3_file_path)  # 创建/链接一个 Connection 对象，它代表数据库
    c = conn.cursor()  # 创建一个 Cursor 游标对象

    for s in config_get:
        c.execute('UPDATE Config SET left=?, up=?, right=?, below=? WHERE Oscilloscope=?', s)  # 写入配置数据库

    conn.commit()  # 执行变更
    conn.close()  # 关闭数据库链接

    tablewidget_qt = self.tableWidget.isEnabled()  # 获取表格是否需要显示
    negative_film_qt = self.checkBox.isChecked()  # 获取图像是否需要反相
    image_matching_qt = self.checkBox_2.isChecked()  # 获取是否采用图像匹配方法
    fast_recognition_qt = self.checkBox_3.isChecked()  # 获取是否采用快速识别
    super_fast_recognition_qt = self.checkBox_4.isChecked()  # 获取是否采用超快速识别

    fo = open(config_file_path, 'w')
    fo.write('tablewidget_qt=%d\n'
             'negative_film_qt=%d\n'
             'image_matching_qt=%d\n'
             'fast_recognition_qt=%d\n'
             'super_fast_recognition_qt=%d\n' % (
                 tablewidget_qt, negative_film_qt, image_matching_qt, fast_recognition_qt, super_fast_recognition_qt))
    fo.close()

    self2.close()  # 关闭窗口


def checkbox_2_statechanged(self):  # 勾选的逻辑
    if self.checkBox_2.isChecked() is False:
        self.tableWidget.setEnabled(True)

        self.checkBox_3.setChecked(False)
        self.checkBox_3.setEnabled(False)

        self.checkBox_4.setChecked(False)
        self.checkBox_4.setEnabled(False)
    else:
        self.tableWidget.setEnabled(False)
        self.checkBox_3.setEnabled(True)
        self.checkBox_4.setEnabled(True)


def checkbox_3_statechanged(self):  # 勾选的逻辑
    if self.checkBox_3.isChecked() is True:
        self.checkBox_4.setChecked(False)


def checkBox_4_stateChanged(self):  # 勾选的逻辑
    if self.checkBox_4.isChecked() is True:
        self.checkBox_3.setChecked(False)


def actionedit_config_click_success():  # 打开设置页面，初始化配置
    MainDialog = QDialog()  # 创建一个主窗体（必须要有一个主窗体）
    myDialog = edit_ui.Ui_Dialog()  # 创建对话框
    myDialog.setupUi(MainDialog)  # 将对话框依附于主窗体
    MainDialog.setWindowModality(Qt.ApplicationModal)

    myDialog.buttonBox.accepted.connect(partial(accept, myDialog, MainDialog))  # OK
    myDialog.buttonBox.rejected.connect(MainDialog.close)  # Cancel

    config_res, config_ini = get_config_data()  # 获取配置数据，截图区域

    for i in range(6):
        for j in range(4):
            myDialog.tableWidget.setItem(i, j, QTableWidgetItem(str(config_res[i][j + 1])))  # 更新表格数据

    myDialog.checkBox_2.stateChanged.connect(partial(checkbox_2_statechanged, myDialog))
    myDialog.checkBox_3.stateChanged.connect(partial(checkbox_3_statechanged, myDialog))
    myDialog.checkBox_4.stateChanged.connect(partial(checkBox_4_stateChanged, myDialog))

    myDialog.tableWidget.setEnabled(config_ini[0])  # 根据配置设置是否勾选
    myDialog.checkBox.setChecked(config_ini[1])  # 根据配置设置是否勾选
    myDialog.checkBox_2.setChecked(config_ini[2])  # 根据配置设置是否勾选
    myDialog.checkBox_3.setChecked(config_ini[3])  # 根据配置设置是否勾选
    myDialog.checkBox_4.setChecked(config_ini[4])  # 根据配置设置是否勾选

    MainDialog.show()  # 显示窗口
    MainDialog.exec_()


def ui_up_to_data(self, p, log):  # 更新主页面，弹出窗口
    if 0 <= p <= 100:
        self.progressBar.setValue(p)  # 更新进度条
    elif p == -1:
        QMessageBox.critical(MainWindow, 'ERROR', log, QMessageBox.Ok)  # 弹出警告

    self.textEdit_log.insertPlainText(log)  # 打印log
    self.textEdit_log.moveCursor(QTextCursor.End)  # 移动光标到最后


def thread_finish(self):  # 进程结束，更新窗口按钮状态
    self.pushButton_generate.setEnabled(True)  # 启用按钮
    self.pushButton_open.setEnabled(True)  # 启用按钮
    self.pushButton_open_2.setEnabled(True)  # 启用按钮


def thread_start(self):  # 启动子线程
    self.thread_1 = osc_value.Worker()  # 创建并启用子线程

    self.thread_1.signal1.connect(partial(ui_up_to_data, self))  # 链接信号
    self.thread_1.finished.connect(partial(thread_finish, self))  # 链接线程结束信号

    self.thread_1.start()  # 启动线程


def actionopen_click_success(self):  # 同时设置
    pushbutton_open_click_success(self)  # 图像路径选择
    pushbutton_open_2_click_success(self)  # 标准文件路径选择


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = osc_value_ui.Ui_MainWindow()
    ui.setupUi(MainWindow)

    ui.gif = QMovie(gif_file_path)  # 配置GIF动图
    ui.label.setMovie(ui.gif)
    ui.gif.start()

    ui.pushButton_open.clicked.connect(partial(pushbutton_open_click_success, ui))  # 链接图像路径选择按钮
    ui.pushButton_open_2.clicked.connect(partial(pushbutton_open_2_click_success, ui))  # 链接标准文件路径选择按钮
    ui.pushButton_exit.clicked.connect(app.quit)  # 链接退出按钮
    ui.pushButton_generate.clicked.connect(partial(pushbutton_generate_click_success, ui))  # 链接执行按钮

    ui.actionopen.triggered.connect(partial(actionopen_click_success, ui))  # 链接图像路径选择选项
    ui.actionexit.triggered.connect(app.quit)  # 链接退出选项

    ui.actionEdit_Config.triggered.connect(actionedit_config_click_success)  # 链接编辑配置选项

    ui.actionAbout_QT5.triggered.connect(app.aboutQt)  # 链接关于选项
    ui.actionAbout_Python3.triggered.connect(actionabout_python3_click_success)  # 链接关于选项
    ui.actionAbout_Help.triggered.connect(actionabout_help_click_success)  # 链接关于选项

    MainWindow.show()  # 显示主窗口
    sys.exit(app.exec())
