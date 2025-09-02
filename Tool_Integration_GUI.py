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
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QLabel
from PyQt5.QtCore import QUrl, QTimer, QDateTime, Qt, QPoint, QThread, pyqtSignal



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
                    
                    if station_no:
                        print("Station_name:",station_no)
                    else:
                        station_no = "01"
                    start_reg = 0
                    if station_no == "01":
                        start_reg = 0
                    elif station_no == "02":
                        start_reg = 100
                    if station_no == "03":
                        start_reg = 200
                    if station_no == "04":
                        start_reg = 300
                    if station_no == "05":
                        start_reg = 400
                    if station_no == "06":
                        start_reg = 500
                    if station_no == "07":
                        start_reg = 600
                    if station_no == "08":
                        start_reg = 700
                    if station_no == "09":
                        start_reg = 800
                    if station_no == "10":
                        start_reg = 900

                    values = [self.client.read_holding_registers(start_reg, 99)]
                    # print(self.client.read_holding_registers(130, 1))
                    # print(start_reg)

                # Emit the signal with the read values
                self.update_gui_signal.emit(values)

            except Exception as e:
                print(f"Error in ModbusWorker: {e}")
                print("Error During read holding resister")
            time.sleep(0.5)
    def stop(self):
        self._running = False
        self.client.close()
        self.quit()
        self.wait()
        self.terminate()


class GUI_load(QMainWindow):
    def __init__(self):
        super(GUI_load, self).__init__()
        ui_file = os.path.join("./Station_GUI_Livgaurd_v02.ui")
        loadUi(ui_file, self)

        # Load paths.json file
        with open('paths.json', 'r') as json_file:
            self.paths_data = json.load(json_file)

        # Use image path from JSON
        self.img_path = self.paths_data["inputs"]["image_path"]
        self.im = QtGui.QPixmap(self.img_path)
        # Set up QLabel to display the image
        self.img_lbl.setPixmap(self.im)
        self.img_lbl.setScaledContents(True)  # Scale image to fit QLabel
        #self.img_lbl.setFixedSize(self.im.size())  # Set QLabel size to match image size    

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

        # Load and display the PDF
        self.load_pdf()



        # #### login password ###########
        self.is_logged_in = False
        self.user_passwords = {"Operator_01": "111", "Manager": "mgr", "Admin": "123", }
        self.user_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.longin_btn.clicked.connect(self.toggle_login_logout)

        # ###### PDF and Excel btns call ########
        self.open_pdf_btn.clicked.connect(self.open_pdf_file)
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

        # video timer
        self.video_timer = QTimer(self)
        self.video_timer.timeout.connect(self.next_video_frame)
        self.video_capture = None
        self.video_fps = 0
        self.playing = False

        # Add Default Functions Call here -------------------
        self.loadDataFromFile()
        self.IMG_load()
        self.Input_Data_Load()
        self.load_previous_video()


        # PLC connection class call pass ip and port
        self.modbus_worker = ModbusWorker(host="192.168.10.81", port=502)
        self.modbus_worker.update_gui_signal.connect(self.update_gui)
        self.modbus_worker.start()


        self.tool_data = True
        self.station_and_recipe_data = True

        self.on_user_input_changed('00')

    def open_setting(self):
        self.ui_second_window.show()

    def call_btn_func(self):
        try:
            self.Calling_window.show()
        except Exception as e:
            print(e)

    # Update GUI value
    def update_gui(self, values):
        # values=[[234,243,23,324,423,423,324,243,423,324,423,243,432,432,342,432,435,435,564,543,4,4,4,54,5,5,4,4,4,3,43,45,3,34,4325,43,45,5,453,43,324,432,324,432,432,342,3,32,2,2,32,432,432,432,42,432,432,432,432,432,432,432,432,432,45,54,5,54,45,456,46,34,34,43,432,23,23,23,543,564,65,75,7,7,78,87,8,8]]
        print("Station wise Read Holding Reg : ", values)
        print("total list length : ", len(values))

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
            print(f"mode: {values[0][28]}")
            if values[0][28] == 1:
                self.Auto_mnual_sts.setStyleSheet("background-color: rgb(16, 235, 16);")
                self.Auto_mnual_sts.setText("Auto mode")
            else:
                self.Auto_mnual_sts.setStyleSheet("background-color: rgb(128, 128, 128);")
                self.Auto_mnual_sts.setText("Manual mode")

            battery_id=""
            for i in range(1,26):
                convert_id=self.dword_to_chars(values[0][i])
                battery_id+=convert_id
            #Battery ID update
            self.Battery_id_lbl.setText("Battery ID: "+battery_id)
            # self.Battery_id_lbl.setText(str(19493948))

            # battery_id=str(308540348509)
            # Table step change data
            # print(values[0][28])
            self.on_user_input_changed(text = str(values[0][29]))
            # self.on_user_input_changed(text = str(60))
            # current cycle time
            self.Current_Cycle_Time_lbl.setText(f" Current Cycle Time : {str(values[0][30 ])} Sec ")


            # PLan_prod_count
            self.PLan_prod_count.setText(str(values[0][33]))
            # self.PLan_prod_count.setText(str(123211312))

            # Actual_prod_count
            self.Actual_prod_count.setText(str(values[0][34]))
            # self.Actual_prod_count.setText("345354")

            # Alarm status
            '''
            self.Alarm_Status(alarm = values[0][36])
            if self.OLE_Charts.isVisible():
                threading.Thread(self.OLE_Charts.update_charts, values[0]).start()
                # self.OLE_Charts.update_charts(values[0])
            '''

            print("All GUI Value Update")

        except Exception as e:
            print(e)
            print("Error During GUI values update")


    # Save Input setting data------------------------------
    def save_Input_setting_data(self):
        try:
            img_path = self.new_img_path
            index = 1
            self.station_name = self.ui_second_window.station_name

            # Check for empty fields
            has_error = False
            if not img_path:
                has_error = True
                self.ui_second_window.img_brows_Btn.setStyleSheet("border: 2px solid red;")
            else:
                self.ui_second_window.img_brows_Btn.setStyleSheet("")

            if not self.station_name:
                has_error = True
                self.ui_second_window.station_input.setStyleSheet("border: 2px solid red;")
            else:
                self.ui_second_window.station_input.setStyleSheet("")

            # Show error if there's any missing input
            if has_error:
                self.msgBox.setText("Please fill in all fields.")
                self.msgBox.setWindowTitle("Input Error")
                self.msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
                self.msgBox.exec_()
                return

            # Update paths.json with new data
            self.paths_data["inputs"]["image_path"] = img_path
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
            self.Station_and_recipe_data_change()
        except Exception as e:
            print(f"Error in the save input setting data: {e}")

    # input Data Load --------------------------------------
    def Input_Data_Load(self):
        try:
            # Load data from paths.json instead of inputs
            img_path = self.paths_data["inputs"]["image_path"]
            self.station_name = self.paths_data["station_name"]
            index = self.paths_data["inputs"]["index"]
            if self.station_name == "01":
                self.Station_lbl.setText("Station : Bottom Cell Holder Loading")
            elif self.station_name == "02":
                self.Station_lbl.setText("Station : NTC Assembly and Checking")
            elif self.station_name == "03":
                self.Station_lbl.setText("Station : TOP Cell Holder Loading & Fitting")
            elif self.station_name == "04":
                self.Station_lbl.setText("Station : Welding Fixture & Loading")
            elif self.station_name == "05":
                self.Station_lbl.setText("Station : Visual Inspection Station")
            elif self.station_name == "06":
                self.Station_lbl.setText("Station : Wire Harness Fixing Station")
            elif self.station_name == "07":
                self.Station_lbl.setText("Station : Routing Station 01")
            elif self.station_name == "08":
                self.Station_lbl.setText("Station : Routing Station 02")
            elif self.station_name == "09":
                self.Station_lbl.setText("Station : Routing Station 03")
            elif self.station_name == "10":
                self.Station_lbl.setText("Station : Engraving on battery pack")
            elif self.station_name == "11":
                self.Station_lbl.setText("Station : xxxxxxxxxxxxxxxxxxxxx")
            elif self.station_name == "12":
                self.Station_lbl.setText("Station : xxxxxxxxxxxxxxxxxxxxxxxx")
            elif self.station_name == "13":
                self.Station_lbl.setText("Station : xxxxxxxxxxxxxxxxxxxxxxxxxx")

            if img_path:
                    self.img_path = img_path
                    self.im = QtGui.QPixmap(self.img_path)
                    self.img_lbl.setPixmap(self.im)
            else:
                print("No image path found in the file.")

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
                    self.save_path_to_file(self.new_img_path)
                    self.Update_IMG(self.new_img_path)
        except Exception as e:
            print(e)

    def save_path_to_file(self, img_path):
        try:
            # Update image path in JSON data
            self.paths_data["inputs"]["image_path"] = img_path

            # Save updated JSON data to file
            with open('paths.json', 'w') as json_file:
                json.dump(self.paths_data, json_file, indent=4)
        except Exception as e:
            print(f"Error saving image path: {e}")

    def Update_IMG(self, filepath):
        try:
            self.img_path = filepath
            self.im = QtGui.QPixmap(self.img_path)
            self.img_lbl.setPixmap(self.im)
            print(filepath)
        except:
            print("Img Error handled")

    def IMG_load(self):
        try:
            # Load image path from JSON
            img_path = self.paths_data["inputs"]["image_path"]
            print("::::::station", img_path)
            if img_path:
                self.img_path = img_path
                self.im = QtGui.QPixmap(self.img_path)
                self.img_lbl.setPixmap(self.im)
                self.img_lbl.setScaledContents(True)
            else:
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


    # ==== Resize the window to match the screen size =======================
    def resize_to_screen(self):
        try:
            screen_geometry = QtWidgets.QDesktopWidget().screenGeometry()
            MainWindow.setGeometry(screen_geometry)
        except:
            print("Facing Error during resize window")

    # ========= Exit Window =========================
    def Exit_window(self):
        self.timer.stop()
        self.timer_1.stop()
        sys.exit(app.exec_())

    # Blink Table Row
    def on_user_input_changed(self, text):
        if text=='0':
            text="00"
        print("step value : ", text)

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
            # print("toggle call")
            if self.tableWidget.rowCount() > row:
                if self.tableWidget.item(row, 0):
                    current_color = self.tableWidget.item(row, 0).background().color()
                    # print("current row :", row)
                    self.current_item = self.tableWidget.item(row, 1).text()
                    # print(self.current_item)
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
            print(e)

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
    def loadDataFromFile(self):
        try:
            # Load table data from JSON
            table_data = self.paths_data["table_data"]["general"]
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

    # Recipe name convert in char
    def recipe_to_chars(self, dword):
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
            char1 = chr(byte1)
            char2 = chr(byte2)
            char3 = chr(byte3)
            char4 = chr(byte4)

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
                    self.paths_data["last_pdf_path"] = SOP_pdf_path

                    # Save updated JSON data to file
                    with open('paths.json', 'w') as json_file:
                        json.dump(self.paths_data, json_file, indent=4)
                    
                    SOP_img_path = self.paths_data["last_pdf_path"]  
                    print("SOP_img_path:", SOP_img_path)
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
            print("Error: ", e)
        

    # ======= Load Pdf On Widget ===========================
    def load_pdf_file(self, file_path):
        try:
            url = QUrl.fromLocalFile(file_path)
            self.webview.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)
            self.webview.settings().setAttribute(QWebEngineSettings.PdfViewerEnabled, True)
            self.webview.load(url)
        except:
            print("Handled Exception during Load PDF file")

    def load_pdf(self):
        """Load a PDF file and display the first page."""
        try:
            # Load PDF path from JSON
            self.pdf_path = self.paths_data["last_pdf_path"]
            if os.path.exists(self.pdf_path):
                self.doc = fitz.open(self.pdf_path)
                self.zoom_factor = 1.0
                self.display_page()
            else:
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
                    print("Opened PDF File:", pdf_path)
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
                    print("Opened PDF File:", pdf_path)
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
                    print("Opened PDF File: ", pdf_path)
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

    # ========================================================Img circle Colors change ======================================================

    def Station_and_recipe_data_change(self):
        try:
            recipe_name = str(self.Three_char_Recipe_Name)
            # print(self.Recipe_Name)
            station_name = self.station_name

            if recipe_name == "2P4S":
                # Load image paths from paths.json instead of Recipe_01_img_path
                recipe_img_paths = self.paths_data["recipe_img_path"]["Recipe_01"]

                if station_name == "01":
                    img_path = recipe_img_paths[0]  # First image path for station 0
                    #img_path = r"D:\Program\Software_Project_GUI\SOP\Recipe_01\Station_img_01.png"
                    self.circle_colors = [Qt.red] * 9
                    self.num_circles = 9
                    self.radius = 20
                    self.circle_positions = [QPoint(230, 303), QPoint(278, 412), QPoint(377, 297), QPoint(473, 412),
                                             QPoint(522, 297), QPoint(740, 417), QPoint(836, 295), QPoint(936, 410),
                                             QPoint(984, 305)]

                if station_name == "02":
                    img_path = recipe_img_paths[1]  # Second image path for station 1
                    #img_path = r"D:\Program\Software_Project_GUI\SOP\Recipe_01\Station_img_02.png"
                    self.circle_colors = [Qt.red] * 8
                    self.num_circles = 8
                    self.radius = 20
                    self.circle_positions = [QPoint(251, 244), QPoint(307, 130), QPoint(414, 253), QPoint(523, 131),
                                             QPoint(762, 251), QPoint(760, 250), QPoint(818, 128), QPoint(925, 250)]

                if station_name == "03":
                    img_path = recipe_img_paths[2]
                    img_path = ""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 20
                    self.circle_positions = []

                if station_name == "04":
                    img_path = recipe_img_paths[3]
                    #img_path = r"D:\Program\Software_Project_GUI\SOP\Recipe_01\Station_img_04.png"
                    self.circle_colors = [Qt.red] * 16
                    self.num_circles = 16
                    self.radius = 20
                    self.circle_positions = [QPoint(710, 127), QPoint(225, 302), QPoint(225, 126), QPoint(711, 302),
                                             QPoint(759, 300), QPoint(1246, 126), QPoint(758, 126), QPoint(1245, 304),
                                             QPoint(1245, 314), QPoint(759, 512), QPoint(759, 339), QPoint(1243, 509),
                                             QPoint(710, 509), QPoint(225, 338), QPoint(225, 510), QPoint(709, 335)]

                if station_name == "05":
                    img_path = recipe_img_paths[5]
                    #img_path = ""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 15
                    self.circle_positions = []

                if station_name == "06":
                    img_path = recipe_img_paths[6]
                    #img_path = r"D:\Program\Software_Project_GUI\SOP\Recipe_01\Station_img_06.png"
                    self.circle_colors = [Qt.red] * 8
                    self.num_circles = 8
                    self.radius = 15
                    self.circle_positions = [QPoint(541, 210), QPoint(569, 280), QPoint(618, 253), QPoint(493, 237),
                                             QPoint(353, 321), QPoint(380, 391), QPoint(426, 362), QPoint(305, 346)]

                if station_name == "07":
                    img_path = recipe_img_paths[7]
                    #img_path = r"D:\Program\Software_Project_GUI\SOP\Recipe_01\Station_img_07.png"
                    # self.circle_colors = [Qt.red] * 10
                    # self.num_circles = 10
                    # self.radius = 15
                    # self.circle_positions = [QPoint(1244, 393), QPoint(1244, 240), QPoint(1244, 361), QPoint(1244, 275), QPoint(223, 396), QPoint(223, 363), QPoint(132, 282), QPoint(223, 273), QPoint(223, 245), QPoint(127, 202)]

                if station_name == "08":
                    print("enter in 8")
                    img_path = recipe_img_paths[8]
                    #img_path = r"D:\Program\Software_Project_GUI\SOP\Recipe_01\Station_img_08.png"
                    self.circle_colors = [Qt.red] * 28
                    self.num_circles = 28
                    self.radius = 15

                    self.circle_positions = [QPoint(201, 174), QPoint(170, 106), QPoint(977, 409), QPoint(974, 311),
                                             QPoint(736, 460), QPoint(782, 457), QPoint(355, 19), QPoint(433, 15),
                                             QPoint(268, 42), QPoint(869, 407), QPoint(900, 278), QPoint(247, 168),
                                             QPoint(670, 412), QPoint(485, 33), QPoint(837, 233), QPoint(316, 208),
                                             QPoint(599, 370), QPoint(555, 71), QPoint(769, 184), QPoint(387, 250),
                                             QPoint(523, 325), QPoint(624, 114), QPoint(698, 153), QPoint(459, 287),
                                             QPoint(220, 73), QPoint(923, 376), QPoint(825, 430), QPoint(314, 23)]

                if station_name == "09":
                    img_path = recipe_img_paths[9]
                    #img_path = ""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 15
                    self.circle_positions = []

                if station_name == "10":
                    img_path = recipe_img_paths[10]
                    #img_path = ""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 15
                    self.circle_positions = []

                if img_path:
                    print(img_path)
                    self.img_path = img_path
                    self.im = QtGui.QPixmap(self.img_path)
                    self.img_lbl.setPixmap(self.im)

            # ==================== For Recipe TRRA ==============================
            if recipe_name == "TERR":
                # Load image paths from paths.json instead of Recipe_01_img_path
                recipe_img_paths = self.paths_data["recipe_img_path"]["Recipe_01"]

                if station_name == "01":
                    # img_path = lines[0].strip()
                    img_path = r"D:\Program\Software_Project_GUI\SOP\TERRA\TERRA_STN_01.png"
                    self.circle_colors = [Qt.red] * 9
                    self.num_circles = 9
                    self.radius = 25
                    self.circle_positions = [QPoint(399, 656), QPoint(403, 861), QPoint(621, 658), QPoint(676, 863),
                                             QPoint(892, 656), QPoint(948, 861), QPoint(1169, 656), QPoint(1225, 861),
                                             QPoint(1313, 652)]

                if station_name == "02":
                    # img_path = lines[1].strip()
                    img_path = r"D:\Program\Software_Project_GUI\SOP\TERRA\TERRA_STN_02.png"
                    self.circle_colors = [Qt.red] * 9
                    self.num_circles = 9
                    self.radius = 25
                    self.circle_positions = [QPoint(424, 436), QPoint(516, 195), QPoint(577, 418), QPoint(804, 193),
                                             QPoint(865, 415), QPoint(1092, 195), QPoint(1153, 415), QPoint(1382, 193),
                                             QPoint(1382, 416)]
                # BY PASS
                if station_name == "03":
                    # img_path = lines[2].strip()
                    img_path = ""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 25
                    self.circle_positions = []

                if station_name == "04":
                    # img_path = lines[3].strip()
                    img_path = r"D:\Program\Software_Project_GUI\SOP\TERRA\TERRA_STN_04_AND_07.png"
                    self.circle_colors = [Qt.red] * 8
                    self.num_circles = 8
                    self.radius = 25
                    self.circle_positions = [QPoint(1439, 110), QPoint(480, 445), QPoint(481, 109), QPoint(1442, 452),
                                             QPoint(1442, 499), QPoint(478, 832), QPoint(479, 496), QPoint(1441, 835)]
                # USE 2 TOOLS
                if station_name == "05":
                    # img_path = lines[4].strip()
                    img_path = r"D:\Program\Software_Project_GUI\SOP\TERRA\TERRA_STN_05.png"
                    # self.circle_colors = [Qt.red] * 6
                    # self.num_circles = 6
                    # self.radius = 15
                    # self.circle_positions = [QPoint(203, 546), QPoint(328, 446), QPoint(452, 525), QPoint(498, 499), QPoint(637, 464), QPoint(683, 37)]

                if station_name == "06":
                    # img_path = lines[5].strip()
                    img_path = r"D:\Program\Software_Project_GUI\SOP\TERRA\TERRA_STN_06.png"
                    self.circle_colors = [Qt.red] * 4
                    self.num_circles = 4
                    self.radius = 25
                    self.circle_positions = [QPoint(411, 448), QPoint(598, 554), QPoint(654, 588), QPoint(365, 424)]

                if station_name == "07":
                    # img_path = lines[6].strip()
                    img_path = r"D:\Program\Software_Project_GUI\SOP\TERRA\TERRA_STN_07.png"
                    self.circle_colors = [Qt.red] * 24
                    self.num_circles = 24
                    self.radius = 25
                    self.circle_positions = [QPoint(36, 294), QPoint(56, 214), QPoint(1071, 434), QPoint(1056, 321),
                                             QPoint(638, 562), QPoint(698, 561), QPoint(518, 15), QPoint(400, 16),
                                             QPoint(233, 111), QPoint(863, 470), QPoint(963, 273), QPoint(174, 298),
                                             QPoint(546, 514), QPoint(586, 54), QPoint(870, 221), QPoint(264, 349),
                                             QPoint(451, 456), QPoint(684, 111), QPoint(787, 162), QPoint(365, 406),
                                             QPoint(151, 158), QPoint(959, 416), QPoint(782, 518), QPoint(334, 52)]

                if station_name == "08":
                    print("enter in 8")
                    # img_path = lines[7].strip()
                    img_path = r""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 25

                    self.circle_positions = []

                if station_name == "09":
                    # img_path = lines[8].strip()
                    img_path = ""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 25
                    self.circle_positions = []

                if station_name == "10":
                    # img_path = lines[9].strip()
                    img_path = ""
                    self.circle_colors = [Qt.red] * 0
                    self.num_circles = 0
                    self.radius = 25
                    self.circle_positions = []

                if img_path:
                    print(img_path)
                    self.img_path = img_path
                    self.im = QtGui.QPixmap(self.img_path)
                    self.img_lbl.setPixmap(self.im)

                # self.change_circle_color()

        except Exception as e:
            print(e)
            print("Error During station wise and recipe wise img")

    # Draw circle -----------------------------------------------
    def change_circle_color(self, count):
        print("count : ", count)
        print("enter in circle color")
        # self.circle_colors = [Qt.red] * 28
        # self.num_circles = 28
        user_input = count

        try:
            circle_index = int(user_input) - 1  # Convert to 0 index
            if 0 <= circle_index < self.num_circles:
                self.circle_colors[circle_index] = QColor(16, 235, 16)
                # self.circle_colors[circle_index] = Qt.green

                if circle_index + 1 < self.num_circles:
                    self.circle_colors[circle_index + 1] = QColor(255, 255, 0)  # Yellow color

                self.draw_circles_on_images()
                self.img_lbl.setPixmap(self.im)
                # print("done")
                user_input = user_input + 1

            if int(user_input) == 0:
                for i in range(0, self.num_circles):
                    self.circle_colors[i] = Qt.red
                    # Redraw the circles with updated colors
                    self.draw_circles_on_images()
                    self.img_lbl.setPixmap(self.im)

        except ValueError:
            print("Handled color exception")

    def draw_circles_on_images(self):
        try:
            painter = QtGui.QPainter(self.im)  # self.img_lbl.pixmap()
            painter.setRenderHint(QPainter.Antialiasing)
            # Define circle positions FOR 2P4S AND STATION 7
            if str(self.Three_char_Recipe_Name) == "2P4S" and self.station_name == "07":
                if self.active_63_tool_no <= 1:
                    self.circle_positions = [QPoint(133, 182), QPoint(132, 258)]

                if self.active_63_tool_no == 2:
                    self.circle_positions = [QPoint(1228, 274), QPoint(1225, 366), QPoint(1227, 397), QPoint(1227, 242),
                                             QPoint(242, 272), QPoint(241, 243), QPoint(241, 362),
                                             QPoint(241, 394)]

                if self.tabl_step == "30":
                    self.circle_positions = [QPoint(133, 182), QPoint(132, 258), QPoint(1228, 274), QPoint(1225, 366),
                                             QPoint(1227, 397), QPoint(1227, 242),
                                             QPoint(242, 272), QPoint(241, 243), QPoint(241, 362),
                                             QPoint(241, 394)]

            # Define circle positions FOR TERRA AND STATION 05
            if str(self.Three_char_Recipe_Name) == "TERR" and self.station_name == "05":
                if self.active_63_tool_no <= 1:
                    self.circle_positions = [QPoint(354, 474), QPoint(261, 522)]

                if self.active_63_tool_no == 2:
                    self.circle_positions = [QPoint(530, 573), QPoint(593, 598), QPoint(667, 481), QPoint(706, 459)]

                if self.tabl_step == "30":
                    self.circle_positions = [QPoint(354, 474), QPoint(261, 522), QPoint(530, 573), QPoint(593, 598),
                                             QPoint(667, 481), QPoint(706, 459)]

            '''
                              #    Screw_1         Screw_2           screw_3            Screw_4          Screw_5            Screw_6     """
            circle_positions = [[QPoint(130, 145), QPoint(160, 185), QPoint(190, 225), QPoint(270, 200), QPoint(350, 175), QPoint(430, 145), QPoint(350, 275), QPoint(430, 345)],
                                [QPoint(10, 145), QPoint(100, 145), QPoint(200, 145), QPoint(10, 200), QPoint(100, 200), QPoint(200, 200)],
                                [QPoint(130, 40), QPoint(130, 80), QPoint(130, 120), QPoint(200, 40), QPoint(200, 80), QPoint(200, 120)]
                             ]
            '''
            circle_positions = self.circle_positions

            for i, center in enumerate(circle_positions):
                # self.radius = 10
                color = self.circle_colors[i]
                painter.setPen(color)
                painter.setBrush(QColor(color))
                painter.drawEllipse(center, self.radius, self.radius)
            painter.end()
        except Exception as e:
            print("Error During Draw circle on img : ", e)
    
    def open_video_file(self):
        try:
            filename, _ = QFileDialog.getOpenFileName(None, "Select video", 'C:/Users/STN09/Desktop')
            if filename:
                self.video_capture = cv2.VideoCapture(filename)
                self.video_fps = self.video_capture.get(cv2.CAP_PROP_FPS)
                try:
                    
                    # Save PDF path to JSON
                    self.paths_data["video_path"] = self.video_fps

                    # Save updated JSON data to file
                    with open('paths.json', 'w') as json_file:
                        json.dump(self.paths_data, json_file, indent=4)
                except Exception as e:
                    print("Error saving image path: ", e)
                self.play_video()
        except Exception as e:
            print("Error During Open Video")
            print(e)

    
    
    def load_previous_video(self):
        try:
            # Load video path from paths.json
            video_path = self.paths_data["video_path"] # Use .get() to avoid KeyError if the key doesn't exist
            print("Video path from JSON:", video_path)
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
                print("self.video_capure is Opened")
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

    def resizeEvent(self, event):
        try:
            self.label_video.setGeometry(10, 10, self.width() - 20, self.height() - 70)
            self.pushButton_file.setGeometry(10, self.height() - 50, 100, 40)
            self.pushButton_playpause.setGeometry(120, self.height() - 50, 100, 40)
            self.pushButton_replay.setGeometry(230, self.height() - 50, 100, 40)
        except Exception as e:
            print("Error During resize event", e)

    def Team_lead_call(self):
        try:
            self.modbus_worker.client.write_single_register(64, 1)
            print('Successfully call to team ')
        except Exception as e:
            print('Error during team call')
    def maintainance_call(self):
        try:
            self.modbus_worker.client.write_single_register(64, 2)
            print('Successfully call to maintenance_call ')
        except Exception as e:
            print('Error during to maintenance_call')
    def Engineer_call(self):
        try:
            self.modbus_worker.client.write_single_register(64, 3)
            print('Successfully call to Engineer_call ')
        except Exception as e:
            print('Error during Engineer_call')

    '''
    def open_OLE_Charts(self):
        self.OLE_Charts.show()'
    '''

    def rejectFunctionality(self):
        self.modbus_worker.client.write_single_register(436, 1)
        print("-----------------------------------------------rejected succesfully------------------------------------------------------------------------------------")

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