from PyQt5 import QtCore, QtGui, QtWidgets
from UI_lke2_1 import Ui_SAFTD
from PyQt5.QtCore import QSettings, QDateTime, QDate
from datetime import datetime
from PyQt5.Qt import *

from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from docxtpl import DocxTemplate
from docx2pdf import convert

from pyqtgraph import PlotWidget, plot
import pyqtgraph as pg

from docx2pdf import convert

import sys
import os

class MainWindow(QtWidgets.QMainWindow, Ui_SAFTD):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        for row in range(self.tableWidget_3.rowCount()):
            date_from = QtWidgets.QDateTimeEdit()
            date_from.setDateTime(
                QtCore.QDateTime(QtCore.QDate(2021, 8, 26),
                                 QtCore.QTime(15, 12, 33))
            )
            self.tableWidget_3.setCellWidget(row, 0, date_from)

        for row in range(self.tableWidget_3.rowCount()):
            date_from = QDateTimeEdit()
            date_from.setDateTime(
                QDateTime(QDate(2021, 9, 2))
            )
            date_from.dateTimeChanged.connect(
                lambda dateTime, row=row: self.calculationTime(dateTime, row))
            self.tableWidget_3.setCellWidget(row, 0, date_from)
        self.dateTime0 = self.tableWidget_3.cellWidget(0, 0).dateTime()

        self.pushButton.clicked.connect(self.buttonSaveProject)
        self.pushButton_2.clicked.connect(self.buttonExport)

        self.toolButton.clicked.connect(self.toolDirectorySave)
        self.toolButton_2.clicked.connect(self.toolDirectoryExport)

        self.pushButton_7.clicked.connect(self.buttonAddtable2)
        self.pushButton_8.clicked.connect(self.buttonDeleteTable2)

        self.pushButton_4.clicked.connect(self.buttonAddtable3)
        self.pushButton_5.clicked.connect(self.buttonDeleteTable3)

    def toolDirectorySave(self):
        directory_file = QFileDialog.getExistingDirectory()
        directory_file = directory_file + "/"
        self.lineEdit_10.setText(directory_file)

    def buttonSaveProject(self):
        directory_file = self.lineEdit_10.text()
        name_folder = self.lineEdit_11.text()

        if len(directory_file) > 0 and len(name_folder) > 0:
            rows = self.tableWidget_3.rowCount()
            cols = self.tableWidget_3.columnCount()

            data_for_word = []
            for row in range(rows):
                tmp = []
                for col in range(cols):
                    if col:
                        item = self.tableWidget_3.item(row, col)
                        if col == 1:
                            item = f'{float(item.text()):.2f}' if item else 'No data'
                        else:
                            item = item.text() if item else 'No data'
                    else:
                        item = self.tableWidget_3.cellWidget(row, 0). \
                            dateTime().toString('dd.MM.yyyy hh:mm')
                    tmp.append(item)

                data_for_word.append(tmp)

            for i in data_for_word:
                print(i)
            self.buttonInstallProject(data_for_word)
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Предупреждение")
            msg.setText("Поля «Директория сохранения файла» и «Имя сохраняемого файла» пустые")
            msg.setIcon(QMessageBox.Warning)

            msg.exec_()

    def buttonInstallProject(self, data):
        data_for_word = []

        for item in data:
            if any(item):
                data_for_word.append({
                    "data": item[0],
                    "time": item[1],
                    "ph": item[2],
                    "ph2": item[3],
                    "fe": item[4],
                    "pm": item[5],
                    "co2": item[6],
                    "pm2": item[7],
                    "pm3": item[8]
                })

        print()
        for i in data_for_word:
            print(i)

        lineEdit = self.lineEdit.text()
        lineEdit_2 = self.lineEdit_2.text()
        lineEdit_3 = self.lineEdit_3.text()
        lineEdit_4 = self.lineEdit_4.text()
        lineEdit_5 = self.lineEdit_5.text()
        lineEdit_6 = self.lineEdit_6.text()
        lineEdit_7 = self.lineEdit_7.text()
        lineEdit_8 = self.lineEdit_8.text()

        lineEdittime = self.lineEdit_dop_1.text()
        lineEdittime_2 = self.lineEdit_dop_2.text()
        lineEdittime_3 = self.lineEdit_dop_3.text()
        lineEdittime_4 = self.lineEdit_dop_4.text()
        lineEdittime_5 = self.lineEdit_dop_5.text()
        lineEdittime_6 = self.lineEdit_dop_6.text()
        lineEdittime_7 = self.lineEdit_dop_7.text()
        lineEdittime_8 = self.lineEdit_dop_8.text()
        lineEdittime_9 = self.lineEdit_dop_9.text()

        dateEdit_1 = self.dateEdit_dop_1.date()
        date_1 = dateEdit_1.toPyDate()
        dateEdit_2 = self.dateEdit_dop_2.date()
        date_2 = dateEdit_2.toPyDate()
        dateEdit_3 = self.dateEdit_dop_3.date()
        date_3 = dateEdit_3.toPyDate()
        dateEdit_4 = self.dateEdit_dop_4.date()
        date_4 = dateEdit_4.toPyDate()
        dateEdit_5 = self.dateEdit_dop_5.date()
        date_5 = dateEdit_5.toPyDate()
        dateEdit_6 = self.dateEdit_dop_6.date()
        date_6 = dateEdit_6.toPyDate()
        dateEdit_7 = self.dateEdit_dop_7.date()
        date_7 = dateEdit_7.toPyDate()
        dateEdit_8 = self.dateEdit_dop_8.date()
        date_8 = dateEdit_8.toPyDate()
        dateEdit_9 = self.dateEdit_dop_9.date()
        date_9 = dateEdit_9.toPyDate()

        doc = DocxTemplate('таблица.docx')
        context = {
            'tbl_contents': data_for_word,
            "namber": lineEdit, "data": lineEdit_2, "namber_2": lineEdit_3, "data_2": lineEdit_4,
            "object": lineEdit_5, "marking": lineEdit_6, "data_start": lineEdit_7, "data_end": lineEdit_8,
            "equipment_1": lineEdittime, "eq_data_1": date_1, "equipment_2": lineEdittime_2,
            "eq_data_2": date_2, "equipment_3": lineEdittime_3, "eq_data_3": date_3,
            "equipment_4": lineEdittime_4,
            "eq_data_4": date_4, "equipment_5": lineEdittime_5, "eq_data_5": date_5,
            "equipment_6": lineEdittime_6,
            "eq_data_6": date_6, "equipment_7": lineEdittime_7, "eq_data_7": date_7,
            "equipment_8": lineEdittime_8,
            "eq_data_8": date_8, "equipment_9": lineEdittime_9, "eq_data_9": date_9
        }
        doc.render(context)

        directory_file = self.lineEdit_10.text()
        name_folder = self.lineEdit_11.text()
        try:
            save_folder = directory_file + name_folder
            os.mkdir(save_folder)
        except:
            msg = QMessageBox()
            msg.setWindowTitle("Предупреждение")
            msg.setText("Папка с именем " + name_folder + " уже существует")
            msg.setIcon(QMessageBox.Warning)

            msg.exec_()

        save_folder = save_folder + "/Protocol No_1-LKI_2_CO2.docx"
        doc.save(save_folder)

    def toolDirectoryExport(self):
        pass

    def buttonExport(self):
        pass

    def buttonAddtable2(self):
        rowPosition = self.tableWidget_2.rowCount()
        self.tableWidget_2.insertRow(rowPosition)

    def buttonDeleteTable2(self):
        if self.tableWidget_2.rowCount() > 0:
            self.tableWidget_2.removeRow(self.tableWidget_2.rowCount() - 1)

    def calculationTime(self, dateTime, row):
        if row == 0:
            self.dateTime0 = self.tableWidget_3.cellWidget(row, 0).dateTime()
            for row in range(1, self.tableWidget_3.rowCount()):
                dateTime2 = self.tableWidget_3.cellWidget(row, 0).dateTime()
                item = QTableWidgetItem()
                item.setData(Qt.DisplayRole,
                             self.dateTime0.secsTo(dateTime2) / 60. / 60.)
                self.tableWidget_3.setItem(row, 1, item)
            return

        dateTime2 = self.tableWidget_3.cellWidget(row, 0).dateTime()
        item = QTableWidgetItem()
        item.setData(Qt.DisplayRole,
                     self.dateTime0.secsTo(dateTime2) / 60. / 60.)
        self.tableWidget_3.setItem(row, 1, item)

    def buttonAddtable3(self):
        rowPosition = self.tableWidget_3.rowCount()
        self.tableWidget_3.insertRow(rowPosition)
        date_from = QtWidgets.QDateTimeEdit()
        dateTime = QtCore.QDateTime().currentDateTime()
        date_from.setDateTime(dateTime)
        date_from.dateTimeChanged.connect(
            lambda dateTime, row=rowPosition:
            self.calculationTime(dateTime, row))
        self.tableWidget_3.setCellWidget(rowPosition, 0, date_from)

    def buttonDeleteTable3(self):
        if self.tableWidget_3.rowCount() > 0:
            self.tableWidget_3.removeRow(self.tableWidget_3.rowCount() - 1)

    def buttonChartPh(self):
        pass

    def buttonChartFe(self):
        pass

    def buttonChartCO2(self):
        pass

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QtGui.QFont("Times", 10))
    w = MainWindow()
    w.setFixedSize(640, 600)
    w.setWindowTitle("SAFTD")
    w.show()
    sys.exit(app.exec_())