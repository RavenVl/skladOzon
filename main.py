import sys

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QUrl
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox, QApplication, QFileDialog
import pandas as pd
import dbf
from pathlib import Path

import MainWindow
class OzonApp(QtWidgets.QMainWindow, MainWindow.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton_file.clicked.connect(self.select_file)
        self.pushButton_naklad.clicked.connect(self.save_dbf)
        self.pushButton_sklad.clicked.connect(self.select_sklad)
        self.file_name = None
        self.rez_arr = []
        self.sklad_dir = None

    def select_sklad(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.Directory)
        if file_dialog.exec_():
            select_folder = file_dialog.selectedFiles()[0]
            self.lineEdit_2.setText(select_folder)
            self.sklad_dir = select_folder

    def select_file(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel files (*.xls *.xlsx)")
        file_dialog.setFileMode(QFileDialog.ExistingFile, )
        if file_dialog.exec_():
            file_name = file_dialog.selectedFiles()[0]
            self.lineEdit.setText(file_name)
            self.file_name = file_name
            self.process_excel_file(file_name)

    def process_excel_file(self, file_name):
        try:
            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(file_name)
            start_row = 14
            while True:
                data = df.iloc[start_row:start_row+1, 0:19]

                kodpr = data.iloc[0, 2]
                try:
                    kodpr = int(kodpr)
                except:
                    break
                itogo = data.iloc[0, 13]

                ret = data.iloc[0, 17]
                try:
                    ret = int(ret)
                except:
                    ret = 0

                kol = data.iloc[0, 8]
                try:
                    kol = int(kol)
                except:
                    pass
                kol -= ret
                if kol<=0:
                    start_row += 1
                    continue

                cena = itogo/kol




                self.rez_arr.append([kodpr, cena, kol])
                start_row += 1

        except Exception as e:
            print(f"Error reading or processing Excel file: {e}")

    def save_dbf(self):
        path = Path(self.sklad_dir)
        filename = "output.dbf"
        file_name = path / filename

        table = dbf.Table(str(file_name), 'kodpr N(10,0); cena N(10,2); kol N(10,0)', codepage='cp1251')
        table.open(mode=dbf.READ_WRITE)
        # Write the rez_arr array to the DBF file
        for row in self.rez_arr:
            table.append(tuple(row))

        # Close the DBF file
        table.close()

        QMessageBox.information(self, "Успех", f"Накладная сохранена в {file_name}")

def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = OzonApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()