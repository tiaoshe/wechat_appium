import time
from PySide2.QtCore import QObject, Signal
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QFileDialog, QMessageBox
import os, shutil, sys
from appium_wechat import read_excl, arrangement_data, SetInfo
from concurrent.futures import ThreadPoolExecutor
from PySide2.QtWidgets import QTableWidgetItem
from queue import Queue
from control_config import ReadConfig, WriteConfig
from PySide2.QtGui import QIcon


class MySignal(QObject):
    list_print = Signal(list)


class MySignalMain(QObject):
    str_print = Signal(str)


class WechatForm(QObject):
    def __init__(self):
        super().__init__()
        self.ui = QUiLoader().load("ui/index.ui")
        self.ui.pushButton_2.clicked.connect(self.handle_work)
        self.ui.pushButton.clicked.connect(self.download_file)
        self.ui.pushButton_3.clicked.connect(self.check_data)
        self.executor = ThreadPoolExecutor(max_workers=3)
        self.ms = MySignal()
        self.ms.list_print.connect(self.print_on_ui)
        self.msm = MySignalMain()
        self.msm.str_print.connect(self.print_on_ui_main)
        self.q = Queue()
        filepath = os.getcwd() + r"\appium_config.ini"
        self.writer = WriteConfig(filepath)
        self.reader = ReadConfig(filepath)

        self.ui.lineEdit.setText(self.reader.get("run_data", "last_load_path"))
        self.ui.lineEdit_2.setText(self.reader.get("run_data", "last_build_name"))
        self.ui.lineEdit_3.setText(self.reader.get("run_data", "last_build_id"))
        self.ui.lineEdit_5.setText(self.reader.get("run_data", "run_num"))
        self.ui.lineEdit_4.setText(self.reader.get("run_data", "start_num"))
        self.arr_list = None
        self.status = False

    # 停止所有进程方法
    def quit_app(self):
        sys.exit()

    # 显示运行状态数据
    def print_on_ui_main(self, msg):
        self.ui.textBrowser.append(msg)

    # 显示表格
    def print_on_ui(self, arr_list):
        self.arr_list = arr_list
        table = self.ui.tableWidget
        table.setRowCount(0)
        row = 0
        for data in arr_list:
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(data['building_name']))
            table.setItem(row, 1, QTableWidgetItem(data['building_id']))
            table.setItem(row, 2, QTableWidgetItem(data['company_name']))
            table.setItem(row, 3, QTableWidgetItem(data['customer_company_name']))
            table.setItem(row, 4, QTableWidgetItem(data['contacts_phone']))
            table.setItem(row, 5, QTableWidgetItem(data['contacts_type']))
            table.setItem(row, 6, QTableWidgetItem(data['competitor_package']))
            table.setItem(row, 7, QTableWidgetItem(str(data['competitor_money'])))
            table.setItem(row, 8, QTableWidgetItem(data['recommend_info']))
            table.setItem(row, 9, QTableWidgetItem(data['estimated_income']))
            table.setItem(row, 10, QTableWidgetItem(data['marketing']))
            table.setItem(row, 11, QTableWidgetItem(data['summary']))
            row += 1

    def check_data(self):
        # 获取文件路径
        file_path = self.ui.lineEdit.text()
        file_copy = self.executor.submit(self.copy_file, file_path)
        self.msm.str_print.emit("正在读取表格…")
        self.ui.pushButton_3.setEnabled(False)
        # 获取楼宇名称
        building_name = self.ui.lineEdit_2.text()
        # 获取楼宇ID
        building_id = self.ui.lineEdit_3.text()
        # 获取表格开始行
        begin_number = self.ui.lineEdit_4.text()
        # 获取录入条数
        run_number = self.ui.lineEdit_5.text()
        self.executor.submit(self.check_data_worker, file_copy, building_name, building_id, begin_number, run_number)

    # 用户检查数据，程序将要上传的数据做展示
    def check_data_worker(self, file_copy, building_name, building_id, begin_number, run_number):
        while not file_copy.done():
            time.sleep(0.5)
        # 读取表格信息
        company_list = read_excl(begin_number, run_number)
        # 整理表格数据
        arr_data_list = arrangement_data(company_list, building_name, building_id)
        self.ms.list_print.emit(arr_data_list)
        self.msm.str_print.emit("读取结束…")
        self.ui.pushButton_3.setEnabled(True)

    def handle_work(self):
        if self.arr_list is None:
            self.msm.str_print.emit("请先点击【检查数据】,再开始上传")
            QMessageBox.critical(
                self.ui,
                '错误',
                '请先检查数据！')

        else:
            # 调用执行方法
            self.executor.submit(self.appium_run, self.arr_list)
            self.msm.str_print.emit("程序开始运行…")
            # 控制程序访问队列
            self.status = True
            self.executor.submit(self.get_queue)
            # 获取楼宇名称
            building_name = self.ui.lineEdit_2.text()
            # 获取楼宇ID
            building_id = self.ui.lineEdit_3.text()
            # 获取录入条数
            run_number = self.ui.lineEdit_5.text()
            # 获取文件路径
            file_path = self.ui.lineEdit.text()
            self.writer.write("run_data", "run_num", run_number)
            self.writer.write("run_data", "last_load_path", file_path)
            self.writer.write("run_data", "last_build_name", building_name)
            self.writer.write("run_data", "last_build_id", building_id)

    # 点击运行执行的方法
    def appium_run(self, arr_list):
        self.q.put("开始运行appium……")
        self.ui.pushButton_2.setEnabled(False)
        SetInfo().run_main(self.q, arr_list)
        self.ui.pushButton_2.setEnabled(True)
        time.sleep(2.2)
        # 控制队列访问关闭
        self.status = False

    # 运行状态通过队列传递，获取队列中数据
    def get_queue(self):
        while self.status:
            a = self.q.get()
            if a is not None:
                self.msm.str_print.emit(a)
                time.sleep(2)

    # 获取文件名称
    def download_file(self):
        file_names = []
        file_dialog = QFileDialog()
        file_dialog.setVisible(QFileDialog.Detail)
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("all file(*)")
        ret = file_dialog.exec_()
        if ret == QFileDialog.Accepted:
            file_names = file_dialog.selectedFiles()
        self.ui.lineEdit.clear()
        self.ui.lineEdit.setText(file_names[0])

    @staticmethod
    def copy_file(from_file):
        """
        将用户给的目录文件复制到工作目录，通过复制文件进行操作
        :param from_file: 用户上传的文件
        """
        path = os.getcwd() + r"\data\temp_data.xlsx"
        if os.path.isfile(path):
            os.remove(path)
        shutil.copy(from_file, path)


app = QApplication()
app.setWindowIcon(QIcon("logo.png"))
windows = WechatForm()
windows.ui.show()
app.aboutToQuit.connect(windows.quit_app)
app.exec_()
