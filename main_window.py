from PyQt5.QtWidgets import QMainWindow, QMessageBox
from PyQt5 import QtWidgets
import pandas as pd
import time
import pathlib
import platform
import os


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = "DataProcess"
        self.top = 300
        self.left = 400
        self.width = 700
        self.height = 250
        self.orderNumberName = "订单编号"

    def init_window(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setMinimumSize(self.width, self.height)
        homePath = os.path.expanduser('~') + "/Desktop/"
        self.centralwidget = QtWidgets.QWidget()
        # self.centralwidget.setGeometry(QtCore.QRect(self.top, self.left, self.width, self.height))
        self.setCentralWidget(self.centralwidget)

        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)

        # 选取csv文件
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setText(" 选择 csv文件")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setText(homePath)
        self.gridLayout.addWidget(self.lineEdit, 0, 1, 1, 1)

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setText("预览文件")
        self.pushButton.clicked.connect(lambda: self._on_btn_click())
        self.gridLayout.addWidget(self.pushButton, 0, 2, 1, 1)

        # 选取xlsx文件
        self.label_1 = QtWidgets.QLabel(self.centralwidget)
        self.label_1.setText(" 选择xlsx文件")
        self.gridLayout.addWidget(self.label_1, 1, 0, 1, 1)

        self.lineEdit_1 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_1.setText(homePath)
        self.gridLayout.addWidget(self.lineEdit_1, 1, 1, 1, 1)

        self.pushButton_1 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_1.setText("预览文件")
        self.pushButton_1.clicked.connect(lambda: self._on_btn1_click())
        self.gridLayout.addWidget(self.pushButton_1, 1, 2, 1, 1)

        # 设置文件保存路径
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setText(" 文件保存路径")
        self.gridLayout.addWidget(self.label_2, 2, 0, 1, 1)

        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setText(homePath)
        self.gridLayout.addWidget(self.lineEdit_2, 2, 1, 1, 1)

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setText("预览文件")
        self.pushButton_2.clicked.connect(lambda: self._on_btn2_click())
        self.gridLayout.addWidget(self.pushButton_2, 2, 2, 1, 1)

        # 开始处理
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setText(" 过滤条件")
        self.gridLayout.addWidget(self.label_3, 3, 0, 1, 1)

        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setText("米色")
        self.gridLayout.addWidget(self.lineEdit_3, 3, 1, 1, 1)

        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setText("开始合并")
        self.pushButton_3.clicked.connect(lambda: self._on_btn3_click())
        self.gridLayout.addWidget(self.pushButton_3, 3, 2, 1, 1)

        label_space = QtWidgets.QLabel(self.centralwidget)
        label_space.setText("处理发货后数据")
        self.gridLayout.addWidget(label_space, 4, 2, 1, 1)

        # 选取xlsx文件
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setText("选择发货后的文件")
        self.gridLayout.addWidget(self.label_4, 5, 0, 1, 1)

        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_4.setText(homePath)
        self.gridLayout.addWidget(self.lineEdit_4, 5, 1, 1, 1)

        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setText("预览文件")
        self.pushButton_4.clicked.connect(lambda: self._on_btn4_click())
        self.gridLayout.addWidget(self.pushButton_4, 5, 2, 1, 1)

        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setText("处理文件")
        self.pushButton_5.clicked.connect(lambda: self._on_btn5_click())
        self.gridLayout.addWidget(self.pushButton_5, 6, 2, 1, 1)
        self.statusBar().showMessage("请根据提示选择相关文件！")


    def _on_btn_click(self):
        cvsPath = QtWidgets.QFileDialog.getOpenFileName(self,
                                                        "浏览",
                                                        self.lineEdit.text(),
                                                        "Files(*.csv)")

        self.lineEdit.setText(cvsPath[0])
        self.statusBar().showMessage("")

    def _on_btn1_click(self):
        xlsxPath = QtWidgets.QFileDialog.getOpenFileName(self,
                                                        "浏览",
                                                        self.lineEdit_1.text(),
                                                        "Files(*.xlsx)")
        self.lineEdit_1.setText(xlsxPath[0])
        self.statusBar().showMessage("")

    def _on_btn2_click(self):
        savePath = QtWidgets.QFileDialog.getExistingDirectory(self,
                                                             "浏览",
                                                             self.lineEdit_2.text())
        self.lineEdit_2.setText(savePath)
        self.statusBar().showMessage("")

    def _on_btn3_click(self):
        csvPath = self.lineEdit.text()
        if len(csvPath) <= 0 or not csvPath.endswith(".csv"):
            QMessageBox.warning(self, "Warning",
                                self.tr("请选择csv文件!"),
                                QMessageBox.Cancel)
            return
        xlsxPath = self.lineEdit_1.text()
        if len(xlsxPath) <= 0 or not xlsxPath.endswith(".xlsx"):
            QMessageBox.warning(self, "Warning",
                                self.tr("请选择xlsx文件!"),
                                QMessageBox.Cancel)
            return
        savePath = self.lineEdit_2.text()
        self.statusBar().showMessage("正在处理中...")
        self._data_merge_process([csvPath, xlsxPath], self.lineEdit_3.text(), savePath)
        self._open_window(savePath)

    def _on_btn4_click(self):
        xlsxPath = QtWidgets.QFileDialog.getOpenFileName(self,
                                                         "浏览",
                                                         self.lineEdit_4.text(),
                                                         "Files(*.xlsx)")
        self.lineEdit_4.setText(xlsxPath[0])
        self.statusBar().showMessage("")

    def _on_btn5_click(self):
        xlsxPath = self.lineEdit_4.text()
        if len(xlsxPath) <= 0 or not xlsxPath.endswith(".xlsx"):
            QMessageBox.warning(self, "Warning",
                                self.tr("请选择xlsx文件!"),
                                QMessageBox.Cancel)
            return
        savePath = self.lineEdit_2.text()
        self._data_delete_process(xlsxPath, savePath)
        self._open_window(savePath)

    def _data_delete_process(self, xlsxPath, outputPath):
        data = pd.read_excel(xlsxPath)
        # 将订单编号类型转换为字符串
        data[self.orderNumberName] = data[self.orderNumberName].astype(str)
        # 删除NAN的列
        data = data.dropna(axis=1)
        # 重命名最后一列
        columnsNames = data.columns.tolist()
        data = data.rename(columns={columnsNames[-1]: '运单号'})
        # 添加一列
        data["物流公司"] = "圆通速递"
        print(data)
        # 判断当前文件有没有输出文件夹，没有创建一个
        absolutePath = pathlib.Path(outputPath)
        if not absolutePath.exists():
            absolutePath.mkdir()

        # 数据写入文件
        now = time.strftime("%Y-%m-%d %H_%M_%S")
        filePath = "{}".format(absolutePath.absolute()) + '/InputOrderList_{}'.format(now) + ".xlsx"
        data.to_excel(filePath, index=False, engine='openpyxl', columns=[self.orderNumberName, '物流公司', '运单号'])
        self.statusBar().showMessage("处理完成，请查看{}".format(filePath))

    def _data_merge_process(self, argv, filterStr, outputPath):
        data1 = None
        data2 = None

        for item in argv:
            if item.endswith(".csv"):
                data1 = pd.read_csv(item, encoding='gbk')
                columnsNames = data1.columns.tolist()
                if "主订单编号" in columnsNames:
                    data1 = data1.rename(columns={"主订单编号": self.orderNumberName})
                data1[self.orderNumberName] = data1[self.orderNumberName].astype(str)
                # 去掉订单编号中的 = 字符
                data1[self.orderNumberName] = data1[self.orderNumberName].str.replace('=', '')
                # 去掉订单编号中的 " 字符
                data1[self.orderNumberName] = data1[self.orderNumberName].str.replace('"', '')
            elif item.endswith(".xlsx"):
                data2 = pd.read_excel(item)
                data2[self.orderNumberName] = data2[self.orderNumberName].astype(str)
                # 过滤收件人是null的行
                data2 = data2[data2['收货人姓名'].notnull()]
            else:
                filterStr = item

        if data1 is not None and data2 is not None:
            # 根据订单号合并表格
            data = pd.merge(data2, data1, on=self.orderNumberName)

            # 根据过滤条件过滤
            if len(filterStr) > 0:
                data3 = data[data['商品属性'].str.contains(filterStr)]
            else:
                data3 = data
            print(data3)

            # 判断当前文件有没有输出文件夹，没有创建一个
            absolutePath = pathlib.Path(outputPath)
            if not absolutePath.exists():
                absolutePath.mkdir()

            # 数据写入文件
            now = time.strftime("%Y-%m-%d %H_%M_%S")
            filePath = "{}".format(absolutePath.absolute()) + '/ExportOrderList_{}'.format(now) + ".xlsx"
            data3.to_excel(filePath, index=False, engine='openpyxl', columns=[self.orderNumberName, "收货人姓名", "联系手机", "收货地址", "商品属性"])
            self.statusBar().showMessage("处理完成，请查看{}".format(filePath))
        else:
            self.statusBar().showMessage("参数输入错误, 请检查参数！")

    def _open_window(self, path):
        # 打开指定路径的文件夹
        sysstr = platform.system()
        if (sysstr == "Windows"):
            path = path.replace('/', '\\')
            os.system("explorer {}".format(path))
        else:
            os.system("open {}".format(path))
