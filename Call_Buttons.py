from PyQt5.QtWidgets import QDialog
from PyQt5.uic import loadUi
from PyQt5.uic.properties import QtWidgets
import os
import json

class Calling_Butn(QDialog):
    def __init__(self):
        try:
            with open('paths.json', 'r') as json_file:
                self.paths_data = json.load(json_file)
            # Load station_no from paths.json instead of inputs
            self.root_Path = self.paths_data["Root_Path"]
            super(Calling_Butn, self).__init__()
            self.ui_path = self.root_Path
            ui_file = os.path.join(self.ui_path, "Call_btns_dialog.ui")
            loadUi(ui_file, self)
        except Exception as e:
            print(e)
