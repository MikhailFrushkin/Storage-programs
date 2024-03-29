import datetime
import os
import subprocess
from pathlib import Path
from sys import argv, executable

import qdarktheme
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from cargo2 import read_file_cargo
from inventory import read_file
from statistic import read_file_tsd


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(470, 400)
        MainWindow.setWindowIcon(QtGui.QIcon("images/icons/склад.ico"))

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 471, 361))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.tab.setObjectName("tab")
        self.pushButton_3 = QtWidgets.QPushButton(self.tab)
        self.pushButton_3.setGeometry(QtCore.QRect(160, 250, 151, 51))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.label_3 = QtWidgets.QLabel(self.tab)
        self.label_3.setGeometry(QtCore.QRect(140, 60, 221, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.label_3.setFont(font)
        self.label_3.setTextFormat(QtCore.Qt.AutoText)
        self.label_3.setObjectName("label_3")
        self.pushButton = QtWidgets.QPushButton(self.tab)
        self.pushButton.setGeometry(QtCore.QRect(10, 60, 121, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.label = QtWidgets.QLabel(self.tab)
        self.label.setEnabled(True)
        self.label.setGeometry(QtCore.QRect(10, 10, 321, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.label.setFont(font)
        self.label.setTextFormat(QtCore.Qt.AutoText)
        self.label.setObjectName("label")
        self.label_5 = QtWidgets.QLabel(self.tab)
        self.label_5.setGeometry(QtCore.QRect(140, 150, 221, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.label_5.setFont(font)
        self.label_5.setTextFormat(QtCore.Qt.AutoText)
        self.label_5.setObjectName("label_5")
        self.checkBox = QtWidgets.QCheckBox(self.tab)
        self.checkBox.setGeometry(QtCore.QRect(20, 200, 281, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.checkBox.setFont(font)
        self.checkBox.setObjectName("checkBox")
        self.pushButton_2 = QtWidgets.QPushButton(self.tab)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 150, 121, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.label_2 = QtWidgets.QLabel(self.tab)
        self.label_2.setGeometry(QtCore.QRect(10, 110, 301, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.label_2.setFont(font)
        self.label_2.setTextFormat(QtCore.Qt.AutoText)
        self.label_2.setObjectName("label_2")
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.pushButton_4 = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_4.setGeometry(QtCore.QRect(160, 250, 151, 51))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setObjectName("pushButton_4")
        self.label_4 = QtWidgets.QLabel(self.tab_2)
        self.label_4.setGeometry(QtCore.QRect(10, 70, 341, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setTextFormat(QtCore.Qt.AutoText)
        self.label_4.setObjectName("label_4")
        self.pushButton_5 = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_5.setGeometry(QtCore.QRect(10, 40, 121, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setObjectName("pushButton_5")
        self.label_6 = QtWidgets.QLabel(self.tab_2)
        self.label_6.setGeometry(QtCore.QRect(10, 10, 301, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setTextFormat(QtCore.Qt.AutoText)
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.tab_2)
        self.label_7.setGeometry(QtCore.QRect(10, 180, 341, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.label_7.setFont(font)
        self.label_7.setTextFormat(QtCore.Qt.AutoText)
        self.label_7.setObjectName("label_7")
        self.pushButton_6 = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_6.setGeometry(QtCore.QRect(10, 140, 121, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.pushButton_6.setFont(font)
        self.pushButton_6.setObjectName("pushButton_6")
        self.label_8 = QtWidgets.QLabel(self.tab_2)
        self.label_8.setGeometry(QtCore.QRect(10, 110, 341, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.label_8.setFont(font)
        self.label_8.setTextFormat(QtCore.Qt.AutoText)
        self.label_8.setObjectName("label_8")
        self.checkBox_2 = QtWidgets.QCheckBox(self.tab_2)
        self.checkBox_2.setGeometry(QtCore.QRect(10, 210, 451, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.checkBox_2.setFont(font)
        self.checkBox_2.setObjectName("checkBox_2")
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.pushButton_7 = QtWidgets.QPushButton(self.tab_3)
        self.pushButton_7.setGeometry(QtCore.QRect(160, 250, 151, 51))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setObjectName("pushButton_7")
        self.label_9 = QtWidgets.QLabel(self.tab_3)
        self.label_9.setGeometry(QtCore.QRect(10, 10, 301, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.label_9.setFont(font)
        self.label_9.setTextFormat(QtCore.Qt.AutoText)
        self.label_9.setObjectName("label_9")
        self.pushButton_8 = QtWidgets.QPushButton(self.tab_3)
        self.pushButton_8.setGeometry(QtCore.QRect(10, 50, 121, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setObjectName("pushButton_8")
        self.label_10 = QtWidgets.QLabel(self.tab_3)
        self.label_10.setGeometry(QtCore.QRect(10, 90, 341, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(10)
        self.label_10.setFont(font)
        self.label_10.setTextFormat(QtCore.Qt.AutoText)
        self.label_10.setObjectName("label_10")
        self.label_11 = QtWidgets.QLabel(self.tab_3)
        self.label_11.setGeometry(QtCore.QRect(10, 130, 301, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.label_11.setFont(font)
        self.label_11.setTextFormat(QtCore.Qt.AutoText)
        self.label_11.setObjectName("label_11")
        self.label_12 = QtWidgets.QLabel(self.tab_3)
        self.label_12.setGeometry(QtCore.QRect(10, 170, 31, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.label_12.setFont(font)
        self.label_12.setTextFormat(QtCore.Qt.AutoText)
        self.label_12.setObjectName("label_12")
        self.label_13 = QtWidgets.QLabel(self.tab_3)
        self.label_13.setGeometry(QtCore.QRect(100, 170, 31, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.label_13.setFont(font)
        self.label_13.setTextFormat(QtCore.Qt.AutoText)
        self.label_13.setObjectName("label_13")
        self.spinBox = QtWidgets.QSpinBox(self.tab_3)
        self.spinBox.setGeometry(QtCore.QRect(30, 175, 50, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.spinBox.setFont(font)
        self.spinBox.setObjectName("spinBox")
        self.spinBox.setRange(0, 24)
        self.spinBox.setValue(6)
        self.spinBox_2 = QtWidgets.QSpinBox(self.tab_3)
        self.spinBox_2.setGeometry(QtCore.QRect(130, 175, 50, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.spinBox_2.setFont(font)
        self.spinBox_2.setObjectName("spinBox_2")
        self.spinBox_2.setRange(0, 24)
        self.spinBox_2.setValue(17)
        self.tabWidget.addTab(self.tab_3, "")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menuBar = QtWidgets.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 469, 21))
        self.menuBar.setObjectName("menuBar")
        self.menu = QtWidgets.QMenu(self.menuBar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menuBar)
        self.action = QtWidgets.QAction(MainWindow)
        self.action.setObjectName("action")
        self.menu.addAction(self.action)
        self.menuBar.addAction(self.menu.menuAction())

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(2)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Складские программы"))
        self.pushButton_3.setText(_translate("MainWindow", "Выполнить"))
        self.label_3.setText(_translate("MainWindow", "Файл не выбран"))
        self.pushButton.setText(_translate("MainWindow", "Выбрать файл"))
        self.label.setText(_translate("MainWindow", "Выберите файл 6.1 Складские лоты \n"
"проверяемых ячеек"))
        self.label_5.setText(_translate("MainWindow", "Файл не выбран"))
        self.checkBox.setText(_translate("MainWindow", "Показывать только расхождения"))
        self.pushButton_2.setText(_translate("MainWindow", "Выбрать файл(ы)"))
        self.label_2.setText(_translate("MainWindow", "Выберите файл(ы) просчета ячейки(ек)"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Инвентаризация"))
        self.pushButton_4.setText(_translate("MainWindow", "Выполнить"))
        self.label_4.setText(_translate("MainWindow", "Файл не выбран"))
        self.pushButton_5.setText(_translate("MainWindow", "Выбрать файл"))
        self.label_6.setText(_translate("MainWindow", "Выберите файл DVL"))
        self.label_7.setText(_translate("MainWindow", "Файл не выбран"))
        self.pushButton_6.setText(_translate("MainWindow", "Выбрать файл(ы)"))
        self.label_8.setText(_translate("MainWindow", "Выберите файл(ы) сканирования R, B с грузомест"))
        self.checkBox_2.setText(_translate("MainWindow", "Использовать файлы CSV программы Сканер Qr"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "Сверка R"))
        self.pushButton_7.setText(_translate("MainWindow", "Выполнить"))
        self.label_9.setText(_translate("MainWindow", "Выберите файл Отслеживаний заданий тсд"))
        self.pushButton_8.setText(_translate("MainWindow", "Выбрать файл"))
        self.label_10.setText(_translate("MainWindow", "Файл не выбран"))
        self.label_11.setText(_translate("MainWindow", "Выберите рабочие часы магазина по МСК:"))
        self.label_12.setText(_translate("MainWindow", "С "))
        self.label_13.setText(_translate("MainWindow", "По"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "Статистика"))
        self.menu.setTitle(_translate("MainWindow", "Help"))
        self.action.setText(_translate("MainWindow", "Опрограмме"))


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.current_dir = Path.cwd()
        self.setupUi(self)

        self.pushButton.clicked.connect(self.evt_btn_open_file_clicked)
        self.pushButton_2.clicked.connect(self.evt_btn_open_file_clicked2)
        self.pushButton_3.clicked.connect(self.evt_btn_clicked)

        self.pushButton_5.clicked.connect(self.evt_btn_open_file_clicked3)
        self.pushButton_6.clicked.connect(self.evt_btn_open_file_clicked4)
        self.pushButton_4.clicked.connect(self.evt_btn_clicked_cargo)

        self.pushButton_8.clicked.connect(self.evt_btn_open_file_clicked5)
        self.pushButton_7.clicked.connect(self.evt_btn_clicked_statistic)

        self.action.triggered.connect(self.about)

    def about(self):
        QMessageBox.information(self, 'О программе', 'Версия 1.1 '
                                                     'от 1.04.2023\n'
                                                     'Разработал Фрышкин Михаил Александрович')

    def evt_btn_open_file_clicked(self):
        res = QFileDialog.getOpenFileName(self, 'Открыть файл', f'{self.current_dir}', 'Лист XLSX (*.xlsx)')
        if res[0] != '':
            self.label_3.setText(res[0])

    def evt_btn_open_file_clicked2(self):
        res = QFileDialog.getOpenFileNames(self, 'Открыть файл(ы)', f'{self.current_dir}', 'Лист XLSX (*.xlsx)')
        if len(res[0]) > 0:
            self.label_5.setText('\n'.join(res[0]))

    def evt_btn_open_file_clicked3(self):
        res = QFileDialog.getOpenFileName(self, 'Открыть файл', f'{self.current_dir}', 'Лист XLSX (*.xlsx)')
        if res[0] != '':
            self.label_4.setText(res[0])

    def evt_btn_open_file_clicked4(self):
        res = QFileDialog.getOpenFileNames(self, 'Открыть файл(ы)', f'{self.current_dir}', "CSV Files (*.csv)")
        if len(res[0]) > 0:
            self.label_7.setText('\n'.join(res[0]))

    def evt_btn_open_file_clicked5(self):
        res = QFileDialog.getOpenFileName(self, 'Открыть файл', f'{self.current_dir}', 'Лист XLSX (*.xlsx)')
        if res[0] != '':
            self.label_10.setText(res[0])

    def evt_btn_clicked(self):
        file_base = None
        file_std = None

        if self.label_3.text() != 'Файл не выбран':
            file_base = self.label_3.text()
        else:
            QMessageBox.critical(self, 'Не выбран файл 6.1', 'Не выбран файл 6.1')
        if self.label_5.text() != 'Файл не выбран':
            file_std = self.label_5.text().split('\n')
        else:
            QMessageBox.critical(self, 'Не выбран файл просчета', 'Не выбран файл просчета')
        if file_base and file_std:
            self.statusBar().showMessage('Выполняется')
            time_start = datetime.datetime.now()
            read_file(self, file_base, file_std, self.checkBox.checkState())
            print('Время сверки инвентаризации: {} секунд(ы)'.format(
                (datetime.datetime.now() - time_start).total_seconds()))
            self.statusBar().showMessage('Готово')
            try:
                open_file(f'{self.current_dir}/Результат инвентаризации.xlsx')
            except:
                ...

    def evt_btn_clicked_cargo(self):
        file_dvl = None
        file_scan = None

        if self.label_4.text() != 'Файл не выбран':
            file_dvl = self.label_4.text()
        else:
            QMessageBox.critical(self, 'Не выбран файл DLVA', 'Не выбран файл DLVA ')
        if self.label_7.text() != 'Файл не выбран':
            file_scan = self.label_7.text().split('\n')
        else:
            QMessageBox.critical(self, 'Не выбран файл(ы) сканирования', 'Не выбран файл(ы) сканирования')
        if file_dvl and file_scan:
            time_start = datetime.datetime.now()
            try:
                print(file_dvl, file_scan, self.checkBox_2.checkState())
                read_file_cargo(self, file_dvl, file_scan, self.checkBox_2.checkState())
                print('Время сверки эрок: {} секунд(ы)'.format((datetime.datetime.now() - time_start).total_seconds()))
                self.statusBar().showMessage('Готово')
            except Exception as ex:
                self.statusBar().showMessage('Ошибка')
                self.restart()
            try:
                open_file(f'{self.current_dir}/Результат сверки R, B.xlsx')
            except:
                ...

    def evt_btn_clicked_statistic(self):
        file_statistic = None

        if self.label_10.text() != 'Файл не выбран':
            file_statistic = self.label_10.text()
        else:
            QMessageBox.critical(self, 'Не выбран файл Отслеживаний заданий тсд',
                                 'Не выбран файл Отслеживаний заданий тсд ')

        if file_statistic:
            time_start = datetime.datetime.now()
            try:
                read_file_tsd(self, file_statistic)
                print('Время формирования статистики: {} секунд(ы)'.format(
                    (datetime.datetime.now() - time_start).total_seconds()))
                self.statusBar().showMessage('Готово')
            except Exception as ex:
                self.statusBar().showMessage('Ошибка')
                self.restart()
            try:
                open_file(f'{self.current_dir}/Статистика.xlsx')
            except:
                ...

    def restart(self):
        os.execl(executable, os.path.abspath(__file__), *argv)


def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    qdarktheme.setup_theme()
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
