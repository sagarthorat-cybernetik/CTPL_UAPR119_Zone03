import time
import os
import cv2
import sys
import fitz
import json
import threading
import Livguard_resource
#from OEE import OLE_Charts
from PyQt5.uic import loadUi
from Call_Buttons import Calling_Butn
from PyQt5 import QtCore, QtGui, QtWidgets
from Setting_Window import Ui_SecondWindow
from pyModbusTCP.client import ModbusClient
from PyQt5.QtWebEngineWidgets import QWebEngineSettings
from PyQt5.QtGui import QPixmap, QPainter, QColor, QImage
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QLabel, QDialog, QVBoxLayout, QApplication, QWidget, QScrollArea
from PyQt5.QtCore import QUrl, QTimer, QDateTime, Qt, QPoint, QThread, pyqtSignal, QPropertyAnimation, QRect, QSize
import requests
import re
import struct

class PDFViewer(QWidget):
    def __init__(self, pdf_path, parent=None):
        super().__init__(parent)

        # Scroll area for zoomable PDF
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)

        # QLabel to render PDF page
        self.pdf_label = QLabel(self)
        self.pdf_label.setAlignment(Qt.AlignCenter)
        self.scroll_area.setWidget(self.pdf_label)

        # Layout
        layout = QVBoxLayout(self)
        layout.addWidget(self.scroll_area)

        # PDF logic
        self.pdf_path = pdf_path
        self.doc = None
        self.zoom_factor = 1.0
        self.load_pdf()

    def load_pdf(self):
        """Load PDF and show first page"""
        if os.path.exists(self.pdf_path):
            self.doc = fitz.open(self.pdf_path)
            self.zoom_factor = 1.0
            self.display_page()
        else:
            print(f"❌ PDF not found: {self.pdf_path}")

    def display_page(self, page_no=0):
        """Render and display given page"""
        if self.doc:
            page = self.doc[page_no]
            mat = fitz.Matrix(self.zoom_factor, self.zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            self.pdf_label.setPixmap(pixmap)
            self.pdf_label.adjustSize()
    def wheelEvent(self, event):
        """Mouse wheel zoom in/out inside the scrollable PDF area."""
        if event.angleDelta().y() > 0:  # Scroll up
            self.zoom_in()
        else:  # Scroll down
            self.zoom_out()

    def zoom_in(self):
        """Zoom in on the PDF."""
        if self.doc:
            self.zoom_factor *= 1.2
            self.display_page()

    def zoom_out(self):
        """Zoom out on the PDF."""
        if self.doc:
            self.zoom_factor /= 1.2
            self.display_page()


class ModbusWorker(QThread):
    # Signal to update the GUI
    update_gui_signal = pyqtSignal(list)

    def __init__(self, host, port):
        super(ModbusWorker, self).__init__()
        self.host = host
        self.port = port
        self.client = ModbusClient(host=self.host, port=self.port, auto_open=True, timeout=1)
        self._running = True

    def run(self):
        while True:
            try:
                if self.client:
                    # Load paths.json file
                    with open('paths.json', 'r') as json_file:
                        self.paths_data = json.load(json_file)
                    # Load station_no from paths.json instead of inputs
                    station_no = self.paths_data["station_name"]
                    
                    if not station_no:
                        station_no = "01"
                    start_reg = 0
                    if station_no == "01":
                        start_reg = 0
                    elif station_no == "02":
                        start_reg = 100
                    elif station_no == "03":
                        start_reg = 200
                    elif station_no == "04":
                        start_reg = 300
                    elif station_no == "05":
                        start_reg = 400
                    elif station_no == "06":
                        start_reg = 500
                    elif station_no == "07":
                        start_reg = 600
                    elif station_no == "08":
                        start_reg = 700
                    elif station_no == "09":
                        start_reg = 800
                    elif station_no == "10":
                        start_reg = 900
                    elif station_no == "11":
                        start_reg = 1000
                    elif station_no == "12":
                        start_reg = 1100
                    elif station_no == "13":
                        start_reg = 1200
                    elif station_no == "14":
                        start_reg = 1300
                    elif station_no == "15":
                        start_reg = 1400
                    elif station_no == "16":
                        start_reg = 1500
                    elif station_no == "17":
                        start_reg = 1600
                    elif station_no == "18":
                        start_reg = 1700
                    elif station_no == "19":
                        start_reg = 1800
                    elif station_no == "20":
                        start_reg = 1900

                    values = [self.client.read_holding_registers(start_reg, 99)]

                # Emit the signal with the read values
                self.update_gui_signal.emit(values)
            except Exception as e:
                print(f"Error During read holding resister:{e}")

            time.sleep(0.5)
    def stop(self):
        self._running = False
        self.client.close()
        self.quit()
        self.wait()
        self.terminate()


class GUI_load(QMainWindow):
    def __init__(self):
        # Load paths.json file
        self.station_name = None
        self.new_img_path = None
        with open('paths.json', 'r') as json_file:
            self.paths_data = json.load(json_file)

        super(GUI_load, self).__init__()
        # Load station_no from paths.json instead of inputs
        station_no = self.paths_data["station_name"]
        if not station_no:
            station_no = "01"
        start_reg = 0
        if station_no == "08":
            print("station 08")
            ui_file = os.path.join("./Station_GUI_livgaurd_Z02_S08.ui")
        else:
            ui_file = os.path.join("./Station_GUI_Livgaurd_v02.ui")

        loadUi(ui_file, self)

        # Ensure SOP_img_lbl is a QScrollArea
        self.scroll_area = self.SOP_img_lbl
        self.scroll_area.setWidgetResizable(True)

        # Create a QLabel to Display PDF inside the ScrollArea
        self.pdf_label = QLabel(self)
        self.pdf_label.setAlignment(Qt.AlignCenter)
        self.scroll_area.setWidget(self.pdf_label)

        # Load PDF path from JSON
        self.pdf_path = self.paths_data["last_pdf_path"]
        self.doc = None
        self.zoom_factor = 1.0


        # #### login password ###########
        self.is_logged_in = False
        self.user_passwords = {"Operator_01": "111", "Manager": "mgr", "Admin": "123", }
        self.user_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.longin_btn.clicked.connect(self.toggle_login_logout)

        # ###### PDF and Excel btns call ########
        self.opl_pdf_btn.clicked.connect(self.open_pdf_file_2)
        self.open_excel_btn.clicked.connect(self.open_excel_file)
        self.Call_btn.clicked.connect(self.call_btn_func)
        self.setting_btn.clicked.connect(self.open_setting)
        #self.OLE_btn.clicked.connect(self.open_OLE_Charts)
        self.Minimize.clicked.connect(self.minimize_window)

        font = self.tableWidget.font()
        font.setPointSize(12)  # Adjust the font size as needed
        self.tableWidget.setFont(font)
        self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        self.tableWidget.setSelectionMode(QtWidgets.QTableWidget.NoSelection)
        self.tableWidget.setFocusPolicy(Qt.NoFocus)
        header = self.tableWidget.horizontalHeader()
        header.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Plain)
        header.setLineWidth(1)
        self.tableWidget.setHorizontalHeader(header)
        item = QtWidgets.QTableWidgetItem()
        item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsEnabled)
        self.tableWidget.horizontalHeader().setVisible(True)
        self.tableWidget.verticalHeader().setSortIndicatorShown(False)

        #self.show_pdf_widget.setStyleSheet("QWidget { border: 1px solid black; }")



        '''Call Windows Button access'''
        self.Calling_window = Calling_Butn()
        self.Calling_window.Team_lead_call.clicked.connect(self.Team_lead_call)
        self.Calling_window.maintainance_call.clicked.connect(self.maintainance_call)
        self.Calling_window.Engineer_call.clicked.connect(self.Engineer_call)

        '''OLE class import for showing charts window'''
        #self.OLE_Charts = OLE_Charts()



        # Create Setting Window Obj
        # ===================== Instantiate the Ui_SecondWindow class to access its widgets
        '''Setting Window Class Button Call'''
        self.ui_second_window = Ui_SecondWindow()
        self.ui_second_window.img_brows_Btn.clicked.connect(self.Open_IMG)
        self.ui_second_window.video_btn.clicked.connect(self.open_video_file)
        self.ui_second_window.pdf_Brows_btn.clicked.connect(self.Open_PDF)
        self.ui_second_window.input_setting_changes.clicked.connect(self.save_Input_setting_data)
        self.ui_second_window.set_password.clicked.connect(self.loadDataFromFile)


        # Add Extra Material Here ---------
        # ===== Date Time Timer ============================
        self.timer_1 = QTimer()
        self.timer_1.timeout.connect(self.update_datetime)
        self.timer_1.start(1000)  # Update every 1000 milliseconds (1 second)

        # ===== Table Row Blink ============================
        self.current_highlighted_row = -1
        self.blink_timer = None

        # ===== Msg BOX Widget =============================
        self.msgBox = QtWidgets.QMessageBox()
        self.msgBox.setIcon(QtWidgets.QMessageBox.Information)

        # All Button Connection with Function -------------
        self.Exit_btn.clicked.connect(self.Exit_window)
        self.setting_btn.setEnabled(False)


        # Video Buttons
        self.pushButton_file.clicked.connect(self.open_video_file)
        self.pushButton_playpause.clicked.connect(self.toggle_playpause)
        self.pushButton_replay.clicked.connect(self.replay_video)
        # open pdf file in maximize screen
        self.Maximize_pdf.clicked.connect(self.openpopup)

        # video timer
        self.video_timer = QTimer(self)
        self.video_timer.timeout.connect(self.next_video_frame)
        self.video_capture = None
        self.video_fps = 0
        self.playing = False

        # Add Default Functions Call here -------------------
        self.loadDataFromFile()
        self.IMG_load()

        # Load and display the PDF
        self.load_pdf()

        self.Input_Data_Load()
        self.load_previous_video()

        # PLC connection class call pass ip and port
        self.modbus_worker = ModbusWorker(host="192.168.205.161", port=502)
        self.modbus_worker.update_gui_signal.connect(self.update_gui)
        self.modbus_worker.start()

        self.tool_data = True
        self.temp = 0
        self.on_user_input_changed('00')

    def open_setting(self):
        self.ui_second_window.show()

    def call_btn_func(self):
        try:
            self.Calling_window.show()
        except Exception as e:
            print(f"calling Button Function Error {e}")

    # Update GUI value
    def update_gui(self, values):
        #
        # with open('dummydata.json', 'r') as json_file:
        #     values = json.load(json_file)
        # values=values["values"]
        # print(len(values[0]))
        # print(values)
        try:
            if values[0] == None:
                self.PLC_Connection_sts.setStyleSheet("background-color: rgb(255, 34, 16);")
                self.PLC_Connection_sts.setText("PLC Disconnected")

            if values[0][0] >= 0:
                self.PLC_Connection_sts.setStyleSheet("background-color: rgb(16, 235, 16);")
                self.PLC_Connection_sts.setText("PLC Connected")

            if values[0][27] == 1:
                self.Manual_lbl.setStyleSheet("background-color: rgb(16, 235, 16);")
                self.Manual_lbl.setText("System Ready")
            else:
                self.Manual_lbl.setStyleSheet("background-color: rgb(128, 128, 128);")
                self.Manual_lbl.setText("System Fault")
            if values[0][28] == 1:
                self.Auto_mnual_sts.setStyleSheet("background-color: rgb(16, 235, 16);")
                self.Auto_mnual_sts.setText("Auto mode")
            else:
                self.Auto_mnual_sts.setStyleSheet("background-color: rgb(128, 128, 128);")
                self.Auto_mnual_sts.setText("Manual mode")

            #Battery ID update
            battery_id=""
            for i in range(1,26):
                convert_id=self.dword_to_chars(values[0][i])
                battery_id+=convert_id
            self.Battery_id_lbl.setText("Battery ID: "+battery_id)

            # Table step change data
            self.on_user_input_changed(text = str(values[0][29]))

            # current cycle time
            self.Current_Cycle_Time_lbl.setText(f" Current Cycle Time : {str(values[0][30 ])} Sec ")

            # PLan_prod_count
            self.PLan_prod_count.setText(str(values[0][33]))

            # Actual_prod_count
            self.Actual_prod_count.setText(str(values[0][34]))

            if values[0][36] != self.temp:
                self.loadDataFromFile(str(values[0][36]))
                self.IMG_load(str(values[0][36]))
                self.load_pdf(str(values[0][36]))
                self.temp = values[0][36]
            # print(values[0][95])
            # values[0][95] = 1
            if values[0][95] == 1:
                threading.Thread(self.readwritemeterdata()).start()

            # Alarm status
            '''
            self.Alarm_Status(alarm = values[0][36])
            if self.OLE_Charts.isVisible():
                threading.Thread(self.OLE_Charts.update_charts, values[0]).start()
                # self.OLE_Charts.update_charts(values[0])
            '''
        except Exception as e:
            print(f"Error During GUI values update{e}")


    # Save Input setting data------------------------------
    def save_Input_setting_data(self):
        try:
            # Validate inputs
            prev_station = self.paths_data.get("station_name", None)
            self.station_name = self.ui_second_window.station_name
            index = 1

            # Update paths.json with new data
            recipe_no = self.ui_second_window.recipe_no
            self.paths_data["station_name"] = self.station_name
            self.paths_data["inputs"]["index"] = index

            # Save updated JSON data to file
            with open('paths.json', 'w') as json_file:
                json.dump(self.paths_data, json_file, indent=4)

            # Show success message
            self.msgBox.setText("Successfully saved changes.")
            self.msgBox.setWindowTitle("Success")
            self.msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
            self.msgBox.exec_()

            # Reload input data and update UI
            self.Input_Data_Load()
            # Check restart condition
            if self.station_name == "08" or prev_station == "08":
                self.msgBox.setText("Please Restart the Application.")
                self.msgBox.setWindowTitle("Success")
                self.msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
                self.msgBox.exec_()
                return
        except Exception as e:
            print(f"Error in the save input setting data: {e}")

    # input Data Load --------------------------------------
    def Input_Data_Load(self):
        try:
            # Load data from paths.json instead of inputs
            img_path = self.paths_data["inputs"]["image_path"]
            self.station_name = self.paths_data["station_name"]
            index = self.paths_data["inputs"]["index"]

            if not self.station_name:
                self.station_name = "01"
            if self.station_name == "01":
                self.Station_lbl.setText("Station : HRD Testing Station")
            elif self.station_name == "02":
                self.Station_lbl.setText("Station : Housing Insertion Station")
            elif self.station_name == "03":
                self.Station_lbl.setText("Station : PCM Filling Station")
            elif self.station_name == "04":
                self.Station_lbl.setText("Station : BMS Connection Station")
            elif self.station_name == "05":
                self.Station_lbl.setText("Station : Top Cover Attachment Station")
            elif self.station_name == "06":
                self.Station_lbl.setText("Station : Routing And Glueing Station ")
            elif self.station_name == "07":
                self.Station_lbl.setText("Station : Top Cover Closing Station")
            elif self.station_name == "08":
                self.Station_lbl.setText("Station : Leak Testing Station")
            elif self.station_name == "09":
                self.Station_lbl.setText("Station : Manually Housing Insertion Station")
            elif self.station_name == "10":
                self.Station_lbl.setText("Station : Laser Marking Station")
            elif self.station_name == "11":
                self.Station_lbl.setText("Station : EOL Station")
            elif self.station_name == "12":
                self.Station_lbl.setText("Station : PDI Station")
            elif self.station_name == "13":
                self.Station_lbl.setText("Station : Battery Pack 1 Station")
            elif self.station_name == "14":
                self.Station_lbl.setText("Station : Battery Pack 2 Station")
            elif self.station_name == "15":
                self.Station_lbl.setText("Station : Battery Pack 3 Station")
            elif self.station_name == "16":
                self.Station_lbl.setText("Station : Battery Pack 4 Station")
            elif self.station_name == "17":
                self.Station_lbl.setText("Station : xxxxxxxxxxxxxxxxxxxxxxxxxx")
            elif self.station_name == "18":
                self.Station_lbl.setText("Station : xxxxxxxxxxxxxxxxxxxxxxxxxx")
            elif self.station_name == "19":
                self.Station_lbl.setText("Station : xxxxxxxxxxxxxxxxxxxxxxxxxx")
            elif self.station_name == "20":
                self.Station_lbl.setText("Station : xxxxxxxxxxxxxxxxxxxxxxxxxx")

        except Exception as e:
            print(f"Error loading image path from file: {e}")

    # ------ brows IMG and Load and save -----------------------------------------------------
    def Open_IMG(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_dialog = QFileDialog()
            file_dialog.setDirectory('C:/Users/STN09/Desktop')
            file_dialog.setNameFilter("Image Files (*.png *.jpg *.jpeg *.bmp *.gif *.webp)")
            if file_dialog.exec_() == QFileDialog.Accepted:
                filenames = file_dialog.selectedFiles()
                if filenames:
                        self.new_img_path = filenames[0]
                        # Update image path in JSON data
                        recipe_no = self.ui_second_window.recipe_no
                        self.paths_data["inputs"]["image_path"][f"recipe_0{recipe_no}"] = self.new_img_path
                        # Save updated JSON data to file
                        with open('paths.json', 'w') as json_file:
                            json.dump(self.paths_data, json_file, indent=4)
                        # update the image
                        self.im = QtGui.QPixmap(self.new_img_path)
                        self.img_lbl.setPixmap(self.im)

        except Exception as e:
            print(f"error in Open Image Function {e}")

    def IMG_load(self, recipe_no="1"):
        try:
            # Load image path from JSON
            recipe_no = str(recipe_no)
            recipes = ["1", "2", "3", "4", "5"]
            if recipe_no not in recipes:
                print("recipe_no not in recipes")
                recipe_no = self.ui_second_window.recipe_no
                img_path = os.path.join(self.paths_data["Root_Path"], "./invalid.png")
            else:
                img_path = self.paths_data["inputs"]["image_path"][f"recipe_0{recipe_no}"]
            if os.path.exists(img_path):
                self.img_path = img_path
                self.im = QtGui.QPixmap(self.img_path)
                self.img_lbl.setPixmap(self.im)
                self.img_lbl.setScaledContents(True)
            else:
                img_path = os.path.join(self.paths_data["Root_Path"],"./demo.png")
                self.img_path = img_path
                self.im = QtGui.QPixmap(self.img_path)
                self.img_lbl.setPixmap(self.im)
                self.img_lbl.setScaledContents(True)
                print("No image path found in the file.")
        except Exception as e:
            print(f"Error loading image load from file: {e}")

    # ===== Date time----------------------------
    def update_datetime(self):
        try:
            currentDateTime = QDateTime.currentDateTime()
            formattedDateTime = currentDateTime.toString("dd/MM/yyyy hh:mm:ss AP")
            self.Auto_lbl.setText(formattedDateTime)
        except:
            print("Facing Error During Data Time Update")

    # ========= Exit Window =========================
    def Exit_window(self):
        self.timer.stop()
        self.timer_1.stop()
        sys.exit(app.exec_())

    # Blink Table Row
    def on_user_input_changed(self, text):
        if text=='0':
            text="00"
        try:
            if self.blink_timer:
                self.blink_timer.stop()
                self.unhighlight_row(self.current_highlighted_row)
                self.blink_timer = None

            matching_row = -1

            # Find the row where the text matches user input
            for row in range(self.tableWidget.rowCount()):
                item = self.tableWidget.item(row, 0)
                if item and item.text() == text:
                    matching_row = row

                # Highlight the matching row and start blinking effect
                if matching_row != -1:
                    self.current_highlighted_row = matching_row
                    self.highlight_row(matching_row)
                    self.start_blink_effect(matching_row)
                    break

            # Highlight rows above the matching row
            for row in range(self.tableWidget.rowCount()):
                if row <= matching_row:
                    self.highlight_row(row)
                else:
                    self.unhighlight_row(row)
        except:
            print("Error During fill data exception")

    def highlight_row(self, row):
        try:
            for col in range(self.tableWidget.columnCount()):
                self.tableWidget.item(row, col).setBackground(QtGui.QColor(16, 235, 16))  # Set background color
        except:
            print("Error : highlight row")

    def unhighlight_row(self, row):
        try:
            for col in range(self.tableWidget.columnCount()):
                self.tableWidget.item(row, col).setBackground(QtGui.QColor(255, 255, 255))  # Reset background color
        except:
            print("Error : unhighlight row")

    def start_blink_effect(self, row):
        try:
            self.blink_timer = QtCore.QTimer()
            self.blink_timer.timeout.connect(lambda: self.toggle_highlight(row))
            self.blink_timer.start(300)  # 500 ms sleep time
        except:
            print("Error : start blink effect row")

    def toggle_highlight(self, row):
        try:
            if self.tableWidget.rowCount() > row:
                if self.tableWidget.item(row, 0):
                    current_color = self.tableWidget.item(row, 0).background().color()

                    self.current_item = self.tableWidget.item(row, 1).text()

                    self.show_current_sts_lbl.setText(' {}'.format(self.current_item))

                    if current_color == QtGui.QColor(16, 235, 16):  # If already highlighted, Unhighlight
                        self.unhighlight_row(row)
                    else:
                        self.highlight_row(row)
        except:
            print("Error : toggle row")

        # # Ensure the timer is stopped when closing the application

    def closeEvent(self, event):
        try:
            if self.blink_timer and self.blink_timer.isActive():
                self.blink_timer.stop()
                self.blink_timer = None
                self.unhighlight_row(self.current_highlighted_row)
            super().closeEvent(event)
        except:
            print("close event")

    # user Login -------------------
    def toggle_login_logout(self):
        try:
            if self.is_logged_in:
                self.logout()
            else:
                self.login()
        except Exception as e:
            print(f"Error in toggle_login_logout{e}")

    # USER LOGIN HERE AND ALLOW TO ACCESS FOR CONTROL USER WISE
    def login(self):
        try:
            selected_user = self.User_ComboBox.currentText()
            entered_password = self.user_password.text()
            # Check if the entered password matches the stored password
            if self.user_passwords.get(selected_user) == entered_password:
                self.is_logged_in = True
                QtWidgets.QMessageBox.information(self, "Login Successful", f"Welcome, {selected_user}!")
                self.user_password.clear()  # Clear password field after login
                self.longin_btn.setText("Logout")
                if selected_user == "Manager" or selected_user == "Admin":
                    self.setting_btn.setEnabled(True)
            else:
                QtWidgets.QMessageBox.warning(self, "Login Failed", "Invalid password. Please try again.")
        except Exception as e:
            print("Error during user login", e)

    # USER LOGOUT AND REMOVE SETTING OPTION FOR CONTROLLING
    def logout(self):
        try:
            if self.is_logged_in:
                self.is_logged_in = False
                self.user_password.clear()  # Clear password field on logout
                QtWidgets.QMessageBox.information(self, "Logged Out", "You have been logged out.")
                self.longin_btn.setText("Login")
                self.setting_btn.setEnabled(False)
                self.stackedWidget.setCurrentIndex(0)
            else:
                QtWidgets.QMessageBox.warning(self, "Logout Failed", "You are not logged in.")
        except Exception as e:
            print("error during user logout", e)

    # Load Step Table data
    def loadDataFromFile(self, recipe_no="1"):
        try:
            # Load table data from JSON
            recipe_no = str(recipe_no)
            recipes = ["1", "2", "3", "4", "5"]
            if recipe_no not in recipes:
                print("recipe_no not in recipes")
                recipe_no = self.ui_second_window.recipe_no
            table_data = self.paths_data["table_data"][f"recipe_0{recipe_no}"]
            self.tableWidget.setRowCount(len(table_data))
            for row, line in enumerate(table_data):
                columns = line.split(',')
                for col, text in enumerate(columns):
                    item = QtWidgets.QTableWidgetItem(text)
                    item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    self.tableWidget.setItem(row, col, item)

        except FileNotFoundError:
            self.msgBox.setText("'table_data' file not found ")
            self.msgBox.setWindowTitle("Message Box")
            self.msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
            self.msgBox.exec_()
            print("No saved data found. Starting with empty table.'table_data' file not found ")

    # Battery-Id convert in char
    def dword_to_chars(self, dword):
        try:
            # Ensure the number is within the range of a 32-bit unsigned integer
            if not (0 <= dword <= 0xFFFFFFFF):
                raise ValueError("The number is out of the range for a 32-bit unsigned integer.")
            # Extract each byte from the DWORD
            byte1 = (dword >> 24) & 0xFF
            byte2 = (dword >> 16) & 0xFF
            byte3 = (dword >> 8) & 0xFF
            byte4 = dword & 0xFF

            # Convert each byte to its corresponding character
            char1 = chr(byte2)
            char2 = chr(byte1)
            char3 = chr(byte4)
            char4 = chr(byte3)

            # Combine the characters into a string
            result = char1 + char2 + char3 + char4

            return result
        except ValueError as e:
            return f"Error: {e}"

    # ===== Brows PDF =====================================
    def Open_PDF(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_dialog = QFileDialog()
            file_dialog.setDirectory('C:/Users/STN09/Desktop')
            file_dialog.setNameFilter("PDF Files (*.pdf)")

            if file_dialog.exec_() == QFileDialog.Accepted:
                filenames = file_dialog.selectedFiles()
                if filenames:
                    SOP_pdf_path = filenames[0]

                    # Save PDF path to JSON
                    recipe_no = self.ui_second_window.recipe_no
                    self.paths_data["last_pdf_path"][f"recipe_0{recipe_no}"] = SOP_pdf_path

                    # Save updated JSON data to file
                    with open('paths.json', 'w') as json_file:
                        json.dump(self.paths_data, json_file, indent=4)
                    
                    SOP_pdf_path = self.paths_data["last_pdf_path"][f"recipe_0{recipe_no}"]
                    # Load the first page of the PDF as an image
                    doc = fitz.open(SOP_pdf_path)
                    page = doc.load_page(0)  # Load first page
                    pix = page.get_pixmap()  # Render page as image

                    # Convert the image to a QPixmap
                    img = QtGui.QImage(pix.samples, pix.width, pix.height, pix.stride, QtGui.QImage.Format_RGB888)
                    pixmap = QPixmap.fromImage(img)

                    # Ensure pdf_label is inside QScrollArea
                    if isinstance(self.pdf_label, QLabel):
                        self.pdf_label.setPixmap(pixmap)
                        self.pdf_label.setScaledContents(True)
                        self.pdf_label.adjustSize()
                    else:
                        print("Error: self.pdf_label is not a QLabel.")

        except Exception as e:
            print("Error in the open PDF: ", e)

    def load_pdf(self, recipe_no ="1"):
        """Load a PDF file and display the first page."""
        try:
            # Load PDF path from JSON
            recipes = ["1", "2", "3", "4", "5"]
            if recipe_no not in recipes:
                print("recipe_no not in recipes")
                recipe_no = "0"
                self.pdf_path = os.path.join(self.paths_data["Root_Path"],"./invalid.pdf")
            else:
                self.pdf_path = self.paths_data["last_pdf_path"][f"recipe_0{recipe_no}"]
            if os.path.exists(self.pdf_path):
                self.doc = fitz.open(self.pdf_path)
                self.zoom_factor = 1.0
                self.display_page()
            else:
                self.pdf_path = os.path.join(self.paths_data["Root_Path"],"./demo.pdf")
                if os.path.exists(self.pdf_path):
                    self.doc = fitz.open(self.pdf_path)
                    self.zoom_factor = 1.0
                    self.display_page()
                print(f"Error: File {self.pdf_path} does not exist.")
        except Exception as e:
            print(f"Error loading PDF: {e}")

    def display_page(self):
        """Render and display the current page inside SOP_img_lbl."""
        if self.doc:
            page = self.doc[0]  # Display first page
            mat = fitz.Matrix(self.zoom_factor, self.zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            self.pdf_label.setPixmap(pixmap)
            self.pdf_label.adjustSize()

    def wheelEvent(self, event):
        """Mouse wheel zoom in/out inside the scrollable PDF area."""
        if event.angleDelta().y() > 0:  # Scroll up
            self.zoom_in()
        else:  # Scroll down
            self.zoom_out()

    def zoom_in(self):
        """Zoom in on the PDF."""
        if self.doc:
            self.zoom_factor *= 1.2
            self.display_page()

    def zoom_out(self):
        """Zoom out on the PDF."""
        if self.doc:
            self.zoom_factor /= 1.2
            self.display_page()

    # Step 2: Convert floats to Modbus register format
    def float_to_modbus(self, value):
        # Convert float to 4-byte big-endian format
        byte_data = struct.pack('>f', value)
        # Convert bytes to two 16-bit Modbus registers
        registers = [int.from_bytes(byte_data[i:i + 2], byteorder='big') for i in range(0, len(byte_data), 2)]
        return registers

    def readwritemeterdata(self):
        url = "http://192.168.205.185/report_test_res_last_1.txt"
        try:
            respose = requests.get(url, timeout=5)
            if respose.status_code == 200:
                data = respose.text
                match = re.search(r"DL\s+([-+]?\d*\.\d+)", data)
                if match:
                    leak_data = float(match.group(1))
                    self.Leakvalue.setText(f"Leak Value: {leak_data} △Pa")
                    self.modbus_worker.client.write_single_register(796, 1)
                    modbus_registers = [self.float_to_modbus(leak_data)]
                    # print("this is modubus register value",modbus_registers)
                    # Print results
                    self.modbus_worker.client.write_single_register(797, modbus_registers[0][1])
                    self.modbus_worker.client.write_single_register(798, modbus_registers[0][0])
                    time.sleep(1)
                    self.modbus_worker.client.write_single_register(795, 0)
                else:
                    print("leak value not fount")
            else:
                print(f"Failed to read data {respose.status_code}")

        except Exception as e:
            print(f"failed to read write {e}")

    def open_pdf_file(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_dialog = QFileDialog()
            # file_dialog.setDirectory('C:/Users/STN09/Desktop')
            file_dialog.setNameFilter("PDF Files (*.pdf)")

            if file_dialog.exec_() == QFileDialog.Accepted:
                filenames = file_dialog.selectedFiles()
                if filenames:
                    pdf_path = filenames[0]
                    # Open the PDF file with the default viewer
                    os.startfile(pdf_path)
        except Exception as e:
            print("Error opening PDF file: ", e)

    def open_pdf_file_2(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_dialog = QFileDialog()
            # file_dialog.setDirectory('C:/Users/STN09/Desktop')
            file_dialog.setNameFilter("PDF Files (*.pdf)")

            if file_dialog.exec_() == QFileDialog.Accepted:
                filenames = file_dialog.selectedFiles()
                if filenames:
                    pdf_path = filenames[0]
                    # Open the PDF file with the default viewer
                    os.startfile(pdf_path)
        except Exception as e:
            print("Error opening PDF file: ", e)

    def open_excel_file(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_dialog = QFileDialog()
            # file_dialog.setDirectory('C:/Users/STN09/Desktop')
            file_dialog.setNameFilter("PDF Files (*.xlsx)")

            if file_dialog.exec_() == QFileDialog.Accepted:
                filenames = file_dialog.selectedFiles()
                if filenames:
                    pdf_path = filenames
                    # Open the PDF file with the default viewer
                    os.startfile(pdf_path)
        except Exception as e:
            print("Error opening PDF file: ", e)

    # =========== Show Alarm Status -----------------------------------
    def Alarm_Status(self, alarm):
        try:
            number = int(alarm)
            if number == 0:
                self.show_alarm_sts_lbl.setText(' No Alarm')
            elif number == 1:
                self.show_alarm_sts_lbl.setText(' Wait For Battery ID Scan')
            elif number == 2:
                self.show_alarm_sts_lbl.setText(' Battery Present Sensor Absent')
            elif number == 3:
                self.show_alarm_sts_lbl.setText(' Tool Cycle Fail ')
            elif number == 4:
                self.show_alarm_sts_lbl.setText(' Station Operation Time Complete')
            elif number == 5:
                self.show_alarm_sts_lbl.setText(' E005-Space')
            elif number == 6:
                self.show_alarm_sts_lbl.setText(' E006-Space')
            elif number == 7:
                self.show_alarm_sts_lbl.setText(' E007-Space')
            elif number == 8:
                self.show_alarm_sts_lbl.setText(' E008-Space')
            elif number == 9:
                self.show_alarm_sts_lbl.setText(' E009-Space')
            elif number == 10:
                self.show_alarm_sts_lbl.setText(' E010-Space')
            else:
                self.show_alarm_sts_lbl.setText('')
        except ValueError:
            self.show_alarm_sts_lbl.setText('')

    # =========== Show Video -----------------------------------

    def open_video_file(self):
        try:
            filename, _ = QFileDialog.getOpenFileName(None, "Select video", 'C:/Users/STN09/Desktop')
            if filename:
                self.video_capture = cv2.VideoCapture(filename)
                self.video_fps = self.video_capture.get(cv2.CAP_PROP_FPS)
                try:
                     # Save PDF path to JSON
                    self.paths_data["video_path"] = filename

                    # Save updated JSON data to file
                    with open('paths.json', 'w') as json_file:
                        json.dump(self.paths_data, json_file, indent=4)
                except Exception as e:
                    print("Error saving image path: ", e)
                self.play_video()
        except Exception as e:
            print(f"Error During Open Video {e}")

    def load_previous_video(self):
        try:
            # Load video path from paths.json
            video_path = self.paths_data["video_path"] # Use .get() to avoid KeyError if the key doesn't exist
            filename = video_path
            if filename:
                self.video_capture = cv2.VideoCapture(filename)
                self.video_fps = self.video_capture.get(cv2.CAP_PROP_FPS)
                self.play_video()
            else:
                print(f"The file {filename} does not exist.")
        except Exception as e:

            print(e)

    def play_video(self):
        try:
            if not self.video_timer.isActive() and self.video_capture.isOpened():
                self.video_timer.start(int(1000 / self.video_fps))
                self.pushButton_playpause.setText('Pause')
                self.playing = True
        except Exception as e:
            print("Error During play video", e)
            
    def pause_video(self):
        try:
            if self.video_timer.isActive():
                self.video_timer.stop()
                self.pushButton_playpause.setText('Play')

                self.playing = False
        except Exception as e:
            print("Error During pause video", e)

    def toggle_playpause(self):
        try:
            if self.playing:
                self.pause_video()
            else:
                self.play_video()
        except Exception as e:
            print("Error During toggle playpause ", e)

    def replay_video(self):
        try:
            if self.video_capture.isOpened():
                self.video_capture.set(cv2.CAP_PROP_POS_FRAMES, 0)
                self.play_video()
        except Exception as e:
            print("Error During replay video", e)

    def next_video_frame(self):
        try:
            if self.video_capture.isOpened():
                ret, frame = self.video_capture.read()
                if ret:
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    height, width, channel = frame.shape
                    bytes_per_line = channel * width
                    q_image = QImage(frame.data, width, height, bytes_per_line, QImage.Format_RGB888)
                    pixmap = QPixmap.fromImage(q_image)
                    self.label_video.setPixmap(pixmap)
                    self.label_video.setScaledContents(True)
                else:
                    # End of video, restart from the beginning
                    self.video_capture.set(cv2.CAP_PROP_POS_FRAMES, 0)
        except Exception as e:
            print("Error During next video frame ", e)

    def openpopup(self):
        popup = QDialog(self)
        popup.setWindowFlags(Qt.FramelessWindowHint | Qt.Popup)  # close when clicking outside
        popup.setAttribute(Qt.WA_DeleteOnClose)  # auto-cleanup

        layout = QVBoxLayout(popup)

        # ✅ Create a duplicate PDFViewer with same path
        pdf_path = self.paths_data["last_pdf_path"]["recipe_01"]  # adjust key as per your JSON
        pdf_widget_copy = PDFViewer(pdf_path, parent=popup)

        layout.addWidget(pdf_widget_copy)
        popup.setLayout(layout)

        # 🔹 Animate to full screen
        screen_geometry = QApplication.primaryScreen().geometry()
        start_rect = QRect(450, 100, 1460, 850)  # start size
        end_rect = screen_geometry  # expand to full screen

        anim = QPropertyAnimation(popup, b"geometry")
        anim.setDuration(500)
        anim.setStartValue(start_rect)
        anim.setEndValue(end_rect)
        anim.start()

        popup.show()

    def animate_resize(self, target_width, target_height):
        print("animated")
        self.table_pdf_video_frm.raise_()  # bring above other elements
        self.table_pdf_video_frm.show()

        anim = QPropertyAnimation(self.table_pdf_video_frm, b"maximumSize")
        anim.setDuration(500)
        anim.setStartValue(self.table_pdf_video_frm.size())
        anim.setEndValue(QSize(target_width, target_height))
        anim.start()
        self.anim = anim

    def Team_lead_call(self):
        try:
            self.modbus_worker.client.write_single_register(64, 1)
        except Exception as e:
            print('Error during team call')

    def maintainance_call(self):
        try:
            self.modbus_worker.client.write_single_register(64, 2)
        except Exception as e:
            print('Error during to maintenance_call')

    def Engineer_call(self):
        try:
            self.modbus_worker.client.write_single_register(64, 3)
        except Exception as e:
            print('Error during Engineer_call')

    def minimize_window(self):
        # Minimize the window
        self.showMinimized()

if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = GUI_load()
    MainWindow.setObjectName("MainWindow")
    MainWindow.resize(1028, 905)
    icon = QtGui.QIcon()
    icon.addPixmap(QtGui.QPixmap(":/newPrefix/images_ICON/Cybernetik_logo.jpg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
    MainWindow.setWindowIcon(icon)
    MainWindow.setStyleSheet("background-color: rgb(214, 214, 214);")
    MainWindow.setDockOptions(QtWidgets.QMainWindow.AllowTabbedDocks | QtWidgets.QMainWindow.AnimatedDocks)
    MainWindow.setUnifiedTitleAndToolBarOnMac(False)
    MainWindow.setWindowFlags(Qt.FramelessWindowHint)
    screen_geometry = QtWidgets.QDesktopWidget().screenGeometry()
    MainWindow.setGeometry(screen_geometry)
    MainWindow.show()
    sys.exit(app.exec_())