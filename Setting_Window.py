from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QDesktopWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QFileDialog
from PyQt5.uic import loadUi
import cv2
import json
import os

class Ui_SecondWindow(QMainWindow):
    dataSaved = QtCore.pyqtSignal()  # Signal emitted after saving data
    def __init__(self):
        super().__init__()
        try:
            # super(Ui_SecondWindow, self).__init__()
            loadUi(r"./Setting.ui", self)
            self.new_row_Btn.clicked.connect(self.addNewRow)
            self.delete_row_Btn.clicked.connect(self.deleteLastRow)
            self.set_password.clicked.connect(self.saveDataToFile)

            # station_name input
            self.station_input.addItems(["01", "02", "03", "04","05", "06", "07", "08","09", "10", "11", "12","13","14","15","16","17","18","19","20"])  # Set initial items#
            self.station_input.currentIndexChanged.connect(self.station_name_change)
            self.recipe_input.addItems(["2.3 Kwh mooving", "Lectrix 2.3kwh", "Triangular _8S4P", "Smart Battery","Triangular _16S2P"])  # Set initial items#
            self.recipe_input.currentIndexChanged.connect(self.recipe_name_change)
            self.station_name = "01"
            self.recipe_name = "01"
            self.paths_data = None
            self.recipe_no = 1
            # Load paths.json file
            with open('paths.json', 'r') as json_file:
                self.paths_data = json.load(json_file)

            self.loadDataFromFile()
        except Exception as e:
            print("Error setting window:", e)

    def station_name_change(self, index):
        self.station_name = self.station_input.itemText(index)

    def recipe_name_change(self, index):
        self.recipe_name = self.recipe_input.itemText(index)
        self.recipe_no = index+1
        self.loadDataFromFile(index+1)

    def addNewRow(self):
        currentRowCount = self.User_table.rowCount()
        self.User_table.setRowCount(currentRowCount + 1)

    def deleteLastRow(self):
        currentRowCount = self.User_table.rowCount()
        if currentRowCount > 0:
            self.User_table.removeRow(currentRowCount - 1)

    def loadDataFromFile(self, recipe_no=1):
        try:
            # Load table data from JSON
            table_data = self.paths_data["table_data"][f"recipe_0{recipe_no}"]

            self.User_table.setRowCount(len(table_data))
            for row, line in enumerate(table_data):
                columns = line.split(',')
                for col, text in enumerate(columns):
                    item = QtWidgets.QTableWidgetItem(text)
                    self.User_table.setItem(row, col, item)

        except FileNotFoundError:
            print("No saved data found. Starting with empty table.")

    def saveDataToFile(self):
        try:
            # Read the existing data from paths.json
            with open('paths.json', 'r') as json_file:
                self.paths_data = json.load(json_file)

            # Prepare the table data
            rows = self.User_table.rowCount()
            cols = self.User_table.columnCount()

            table_data = []
            for row in range(rows):
                row_data = []
                for col in range(cols):
                    item = self.User_table.item(row, col)
                    row_data.append(item.text() if item else "")
                table_data.append(",".join(row_data))
            # Update the JSON object with the new table data
            self.paths_data["table_data"][f"recipe_0{self.recipe_no}"] = table_data
            # Write the updated data back to the JSON file
            with open('paths.json', 'w') as json_file:
                json.dump(self.paths_data, json_file, indent=4)
                json_file.flush()
                os.fsync(json_file.fileno())

            QtWidgets.QMessageBox.information(self, "Table Updated", "Table data saved successfully.")
            self.dataSaved.emit()  # Emit the signal after saving

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error Saving Data", f"Error saving data: {e}")


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    MainWindow = Ui_SecondWindow()
    MainWindow.setObjectName("MainWindow")
    screen_geometry = QDesktopWidget().screenGeometry()
    # MainWindow.setGeometry(screen_geometry)
    MainWindow.show()
    sys.exit(app.exec_())