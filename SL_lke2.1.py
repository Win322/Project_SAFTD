from PyQt5 import QtCore, QtGui, QtWidgets
from UI_lke2_1 import Ui_SAFTD
from PyQt5.QtCore import QSettings, QDateTime, QDate
from datetime import datetime
from PyQt5.Qt import *

from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from docxtpl import DocxTemplate
from docx2pdf import convert

from docx2pdf import convert

from pyqtgraph import PlotWidget, plot
import pyqtgraph as pg

import sys
import os

CONFIG_FILE_NAME = 'configDateTime.ini'


class TimeAxisItem(pg.AxisItem):
    def tickStrings(self, values, scale, spacing):
        print(values, '\n', scale, '\n', spacing)
        return [datetime.fromtimestamp(value) for value in values]


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

        self._date = QDate.currentDate()
        self.listLineEdit = self.findChildren(QtWidgets.QLineEdit)
        self.listDateEdit = self.findChildren(QtWidgets.QDateEdit)

        self.load_lineEdits = [
            'lineEdit_dop_1',
            'lineEdit_dop_2',
            'lineEdit_dop_3',
            'lineEdit_dop_4',
            'lineEdit_dop_5',
            'lineEdit_dop_6',
            'lineEdit_dop_7',
            'lineEdit_dop_8',
            'lineEdit_dop_9',
        ]

        self.load_settings()

        self.indexColumn = 6
        self.indexColumn_2 = 9

        # self.pushButton.clicked.connect(self.check_date)
        self.pushButton_2.clicked.connect(self.buttonExport)

        self.toolButton.clicked.connect(self.toolDirectorySave)
        self.toolButton_2.clicked.connect(self.toolDirectoryExport)

        self.pushButton_7.clicked.connect(self.buttonAddtable2)
        self.pushButton_8.clicked.connect(self.buttonDeleteTable2)

        self.pushButton_4.clicked.connect(self.buttonAddtable3)
        self.pushButton_5.clicked.connect(self.buttonDeleteTable3)

        self.pushButton_9.clicked.connect(self.buttonChartPh)
        #
        self.pushButton.clicked.connect(self.buttonDialog)
        #

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
                            item = f'{float(item.text()):.0f}' if item else 'No data'
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

        rows = self.tableWidget.rowCount()
        cols = self.tableWidget.columnCount()
        dataTable_1 = []
        for row in range(rows):
            tmp = []
            for col in range(cols):
                try:
                    tmp.append(self.tableWidget.item(row, col).text())
                except:
                    tmp.append('')
            dataTable_1.append(tmp)
        data_for_word_2 = []
        for i in dataTable_1:
            print(i)

        for item in dataTable_1:
            if any(item):
                data_for_word_2.append({
                    "model": item[0],
                    "gaz": item[1],
                    "c": item[2],
                    "v": item[3],
                    "ph_2": item[4],
                    "co": item[5],
                    "co22": item[6],
                    "namb": item[7],
                    "size": item[8]
                })

        rows = self.tableWidget_2.rowCount()
        cols = self.tableWidget_2.columnCount()
        dataTable_3 = []
        for row in range(rows):
            tmp = []
            for col in range(cols):
                try:
                    tmp.append(self.tableWidget_2.item(row, col).text())
                except:
                    tmp.append('')
            dataTable_3.append(tmp)
        data_for_word_3 = []
        for i in dataTable_3:
            print(i)

        for item in dataTable_3:
            if any(item):
                data_for_word_3.append({
                    "mark": item[0],
                    "nam": item[1],
                    "mark2": item[2],
                    "ves": item[3],
                    "ves2": item[4],
                    "mass2": item[5],
                    "mass3": item[6],
                    "ves3": item[7],
                    "mass4": item[8],
                    "speed": item[9]
                })

        rows = self.tableWidget_2.rowCount()

        if not self.indexColumn:
            return

        _sum = 0
        for row in range(rows):
            item = self.tableWidget_2.item(row, self.indexColumn)
            item = item.text() if item else '0'
            _sum += float(item) if item else 0

        average = _sum / rows
        print(average)

        if not self.indexColumn_2:
            return

        _sum2 = 0
        for row in range(rows):
            item = self.tableWidget_2.item(row, self.indexColumn_2)
            item = item.text() if item else '0'
            _sum2 += float(item) if item else 0

        average_2 = _sum2 / rows
        print(average_2)

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

        doc = DocxTemplate('lke2_co2.docx')
        context = {
            'tbl_contents': data_for_word,
            'tbl_contents_2': data_for_word_2,
            'tbl_contents_3': data_for_word_3,
            "namber": lineEdit, "data": lineEdit_2, "namber_2": lineEdit_3, "data_2": lineEdit_4,
            "object": lineEdit_5, "marking": lineEdit_6, "data_start": lineEdit_7, "data_end": lineEdit_8,
            "equipment_1": lineEdittime, "eq_data_1": date_1, "equipment_2": lineEdittime_2,
            "eq_data_2": date_2, "equipment_3": lineEdittime_3, "eq_data_3": date_3,
            "equipment_4": lineEdittime_4,
            "eq_data_4": date_4, "equipment_5": lineEdittime_5, "eq_data_5": date_5,
            "equipment_6": lineEdittime_6,
            "eq_data_6": date_6, "equipment_7": lineEdittime_7, "eq_data_7": date_7,
            "equipment_8": lineEdittime_8,
            "eq_data_8": date_8, "equipment_9": lineEdittime_9, "eq_data_9": date_9,
            "medium": average, "medium2": average_2
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

    def check_date(self):
        self.x_y()
        listLineEdit = self.findChildren(QtWidgets.QLineEdit)
        _dict = {}
        for lineEdit in listLineEdit:
            obj = lineEdit.objectName()
            if obj[0:13] == 'lineEdit_dop_':
                key = obj[13:]
                _dict[key] = lineEdit.text()
        listDateEdit = self.findChildren(QtWidgets.QDateEdit)
        dateEnd = 0
        for dateEdit in listDateEdit:
            if dateEdit.objectName()[0:13] == 'dateEdit_dop_':
                if self._date >= dateEdit.date():
                    key = dateEdit.objectName()[13:]
                    msgBox = QtWidgets.QMessageBox(self)
                    msgBox.setWindowTitle("Уведомление")
                    msgBox.setIcon(QtWidgets.QMessageBox.Warning)
                    msgBox.setText(
                        'Срок действия данных истек, замените даннные<br>'
                        f'Дата окончания: {dateEdit.date().toString("dd:MM:yyyy")}<br>'
                        f'Информация: {_dict[key]}<br>'
                    )
                    msgBox.move(self.x, self.y)
                    self.y += 100
                    msgBox.show()
                else:
                    dateEnd += 1
        if dateEnd == 9:
            self.buttonSaveProject()

    def x_y(self):
        self.x = self.pos().x() + 650
        self.y = self.pos().y()

    def save_settings(self):
        settings = QSettings(CONFIG_FILE_NAME, QSettings.IniFormat)
        settings.setValue('Geometry', self.saveGeometry())
        settings.setValue('WindowState', self.saveState())

        for lineEdit in self.listLineEdit:
            try:
                obj = lineEdit.objectName()
                settings.setValue(obj, lineEdit.text())
            except:
                pass

        for dateEdit in self.listDateEdit:
            obj = dateEdit.objectName()
            settings.setValue(obj, dateEdit.date())

    def load_settings(self):
        settings = QSettings(CONFIG_FILE_NAME, QSettings.IniFormat)
        geometry = settings.value('Geometry')
        if geometry:
            self.restoreGeometry(geometry)
        state = settings.value('WindowState')
        if state:
            self.restoreState(state)

        for lineEdit in self.listLineEdit:
            obj = lineEdit.objectName()
            if obj in self.load_lineEdits:
                lineEdit.setText(settings.value(obj, ""))

        for dateEdit in self.listDateEdit:
            obj = dateEdit.objectName()
            _date = settings.value(obj, QDate.currentDate())
            dateEdit.setDate(_date)

    def closeEvent(self, e):
        self.save_settings()
        super().closeEvent(e)

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
        data_for_word = []

        rows = self.tableWidget_3.rowCount()
        for row in range(rows):
            item_N2 = self.tableWidget_3.item(row, 1)
            _data_N2 = item_N2.data(Qt.DisplayRole) if item_N2 else '0'
            item_N3 = self.tableWidget_3.item(row, 2)
            _data_N3 = item_N3.data(Qt.DisplayRole) if item_N3 else '0'
            data_for_word.append([_data_N2, _data_N3])

        self.x, self.y = [], []
        for x, y in data_for_word:
            print(f'x={x}; y={y}')
            self.x.append(float(x))
            self.y.append(float(y))

        self.widget = pg.PlotWidget()
        pen = pg.mkPen(color=(255, 0, 0), width=2)
        self.widget.plot(
            x=self.x,
            y=self.y, pen=pen, symbol='+'
        )

        w = self.findChild(PlotWidget, 'widget')

        print(f'w --> {w}')
        if w:
            w.deleteLater()
        self.gridLayout_8.addWidget(
            self.widget,
            0, 0, 1, 1,
            alignment=Qt.AlignCenter
        )
        self.widget.setFixedSize(280, 300)
        self.widget.setBackground('w')
        self.widget.setTitle("Ph", color="b", size="12pt")
        styles = {"color": "#f00", "font-size": "12px"}
        self.widget.setLabel("left", "Ph, ед.Ph", **styles)
        self.widget.setLabel("bottom", "Hour(H)", **styles)

        self.widget.showGrid(x=True, y=True)

        self.buttonChartFe()

    def buttonChartFe(self):
        data_for_word = []

        rows = self.tableWidget_3.rowCount()
        for row in range(rows):
            item_N2 = self.tableWidget_3.item(row, 1)
            _data_N2 = item_N2.data(Qt.DisplayRole) if item_N2 else '0'
            item_N3 = self.tableWidget_3.item(row, 4)
            _data_N3 = item_N3.data(Qt.DisplayRole) if item_N3 else '0'
            data_for_word.append([_data_N2, _data_N3])

        self.x, self.y = [], []
        for x, y in data_for_word:
            print(f'x={x}; y={y}')
            self.x.append(float(x))
            self.y.append(float(y))

        self.widget_2 = pg.PlotWidget()
        pen = pg.mkPen(color=(0, 255, 111), width=2)
        self.widget_2.plot(
            x=self.x,
            y=self.y, pen=pen, symbol='+'
        )

        w = self.findChild(PlotWidget, 'widget')

        print(f'w --> {w}')
        if w:
            w.deleteLater()
        self.gridLayout_10.addWidget(
            self.widget_2,
            0, 0, 1, 1,
            alignment=Qt.AlignCenter
        )
        self.widget_2.setFixedSize(269, 269)
        self.widget_2.setBackground('w')
        self.widget_2.setTitle("Fe", color="b", size="12pt")
        styles = {"color": "#f00", "font-size": "12px"}
        self.widget_2.setLabel("left", "C(Fe общ.), мг/дм3", **styles)
        self.widget_2.setLabel("bottom", "Hour(H)", **styles)

        self.widget_2.showGrid(x=True, y=True)

        self.buttonChartCO2()

    def buttonChartCO2(self):
        data_for_word = []

        rows = self.tableWidget_3.rowCount()
        for row in range(rows):
            item_N2 = self.tableWidget_3.item(row, 1)
            _data_N2 = item_N2.data(Qt.DisplayRole) if item_N2 else '0'
            item_N3 = self.tableWidget_3.item(row, 6)
            _data_N3 = item_N3.data(Qt.DisplayRole) if item_N3 else '0'
            data_for_word.append([_data_N2, _data_N3])

        self.x, self.y = [], []
        for x, y in data_for_word:
            print(f'x={x}; y={y}')
            self.x.append(float(x))
            self.y.append(float(y))

        self.widget_3 = pg.PlotWidget()
        pen = pg.mkPen(color=(0, 0, 255), width=2)
        self.widget_3.plot(
            x=self.x,
            y=self.y, pen=pen, symbol='+'
        )

        w = self.findChild(PlotWidget, 'widget')

        print(f'w --> {w}')
        if w:
            w.deleteLater()
        self.gridLayout_11.addWidget(
            self.widget_3,
            0, 0, 1, 1,
            alignment=Qt.AlignCenter
        )
        self.widget_3.setFixedSize(269, 269)
        self.widget_3.setBackground('w')
        self.widget_3.setTitle("CO2", color="b", size="12pt")
        styles = {"color": "#f00", "font-size": "12px"}
        self.widget_3.setLabel("left", "C(CO2), мг/дм3", **styles)
        self.widget_3.setLabel("bottom", "Hour(H)", **styles)

        self.widget_3.showGrid(x=True, y=True)

    def buttonDialog(self):
        self.dialog = Dialog()
        self.dialog.show()

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(320, 240)
        self.gridLayout = QtWidgets.QGridLayout(Dialog)
        self.gridLayout.setObjectName("gridLayout")
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setObjectName("lineEdit")
        self.gridLayout.addWidget(self.lineEdit, 1, 0, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 3, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 2, 0, 1, 1)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 7, 0, 1, 2)
        self.checkBox = QtWidgets.QCheckBox(Dialog)
        self.checkBox.setObjectName("checkBox")
        self.gridLayout.addWidget(self.checkBox, 1, 1, 1, 1)
        self.checkBox_2 = QtWidgets.QCheckBox(Dialog)
        self.checkBox_2.setObjectName("checkBox_2")
        self.gridLayout.addWidget(self.checkBox_2, 3, 1, 1, 1)
        self.line = QtWidgets.QFrame(Dialog)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.gridLayout.addWidget(self.line, 6, 0, 1, 2)
        self.line_2 = QtWidgets.QFrame(Dialog)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.gridLayout.addWidget(self.line_2, 4, 0, 1, 2)
        self.checkBox_3 = QtWidgets.QCheckBox(Dialog)
        self.checkBox_3.setObjectName("checkBox_3")
        self.gridLayout.addWidget(self.checkBox_3, 5, 0, 1, 2)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label_2.setText(_translate("Dialog", "Имя документа типа \"Журнал\""))
        self.label.setText(_translate("Dialog", "Имя документа типа \"Протокол\""))
        self.pushButton.setText(_translate("Dialog", "Загрузить данные"))
        self.checkBox.setText(_translate("Dialog", "Заполнить документ"))
        self.checkBox_2.setText(_translate("Dialog", "Заполнить документ"))
        self.checkBox_3.setText(_translate("Dialog", "Проверить на истечение срока действия"))

class Dialog(QDialog, Ui_Dialog):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def checkBox(self):
        pass

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QtGui.QFont("Times", 10))
    w = MainWindow()
    w.setFixedSize(640, 640)
    w.setWindowTitle("AutomaticDocuments-CO2 corrosion")
    w.show()
    sys.exit(app.exec_())