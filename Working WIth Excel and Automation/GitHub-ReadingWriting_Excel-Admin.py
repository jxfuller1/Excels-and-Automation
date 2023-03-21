import os
import win32com.client as win32
import time
import pyperclip
import ctypes
from distutils.dir_util import copy_tree
import getpass
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys
import win32gui
import win32con
import Admin_Helper_setup_go_to
import Admin_Helper_windows
import shutil
import pyautogui
import pandas as pd
import easygui
import win32api

#OVERVIEW OF PROGRAM:
#program finds a window/program on desktop, that program is able to output an excel file which i then read
#based on what's in that excel, I read for file directories at certain locations and output results based on what is found
#in the directories

#for finding all child windows in a window
def get_child_windows(parent):
    if not parent:
        return
    hwndChildList = []
    win32gui.EnumChildWindows(parent, lambda hwnd, param: param.append(hwnd), hwndChildList)
    return hwndChildList

#instructions window part of GUI
class Instructions_Window(QMainWindow):

    def __init__(self, x, y):
        super().__init__()
        self.x = int(x) + 50
        self.y = int(y) + 100
        self.initUI()

    def initUI(self):
        self.setGeometry(self.x, self.y, 460, 100)
        self.setFixedSize(self.size())
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        self.setWindowTitle("INSTRUCTIONS")

        self.label = QLabel("- Make sure Admin is not minimized and in main monitor.", self)
        self.label.adjustSize()
        self.label.move(10, 15)

        self.label2 = QLabel("- Program will check if any on QP waiting list can <b>potentially</b> be completed", self)
        self.label2.adjustSize()
        self.label2.move(10, 35)

        self.label3 = QLabel("- Any results returned, check FAI to make sure it's filled out and done properly for new rev!!!", self)
        self.label3.adjustSize()
        self.label3.move(10, 55)

#main GUI
class Actions(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def instructions(self):
        a = str(self.pos())
        b = a.split('QPoint')
        c = b[1].split(',')
        x = ''.join(filter(str.isdigit, c[0]))
        y = ''.join(filter(str.isdigit, c[1]))

        self.w = Instructions_Window(int(x), int(y))
        self.w.show()

    def initUI(self):
        self.setGeometry(400,400,315,375)
        self.setFixedSize(self.size())
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        self.setWindowTitle("Auto Check QP Waiting")

        self.menubar = QMenuBar()
        self.setMenuBar(self.menubar)

        actionFile = self.menubar.addMenu("Help")
        actionFile.addAction("Instructions", self.instructions)

        myfont = QFont()
        myfont.setPointSize(7)

        self.about = QLabel("About: This Program will check to see if any on QP\n            "
                            "waiting list can be finished.\n\n           "
                            "       ***See Help for Instructions***", self)
        self.about.adjustSize()
        self.about.setFont(myfont)
        self.about.move(40, 23)

        self.endresult = QLabel("Potential for completion!", self)
        self.endresult.adjustSize()
        self.endresult.move(105, 135)

        self.running = QLabel("                               ", self)
        self.running.adjustSize()
        self.running.move(90, 110)

        self.button = QPushButton(self)
        self.button.setText("Start")
        self.button.adjustSize()
        self.button.clicked.connect(self.export_cert_conformity)
        self.button.move(120, 80)

        self.complete = QTextEdit(self)
        self.complete.setGeometry(60, 150, 200, 200)

        self.tool = QLabel("<b>NOTE:</b> CHECK manually for tools/tooling validation!", self)
        self.tool.adjustSize()
        self.tool.move(40, 353)

        # showing all the widgets
        self.show()

    #method for finding a specific window and using it's export feature to export to an excel file
    def export_cert_conformity(self):

        user = getpass.getuser()

        Quality_window = win32gui.FindWindow(None, "Epic Quality Admin")
        win32gui.SetForegroundWindow(Quality_window)
        time.sleep(.5)

        x, y, w, h = win32gui.GetWindowRect(Quality_window)

        pyautogui.click(x + 100, y + 60)
        time.sleep(1)
        pyautogui.click(x + 60, y + 130)
        time.sleep(2)

        admin_inspection_childs = get_child_windows(Quality_window)

        for i in range(len(admin_inspection_childs)):
            parentwindow = (win32gui.GetWindowText(admin_inspection_childs[i]))
            if "Export" in parentwindow:
                x, y, w, h = win32gui.GetWindowRect(admin_inspection_childs[i])
                pyautogui.click(x+10, y+10)

        time.sleep(2)

        location_qpwaiting = "pathway to save excel file

        try:
            pyautogui.typewrite(location_qpwaiting)
            time.sleep(.5)
            pyautogui.press('enter')
            time.sleep(.5)
            pyautogui.press('enter')

            time.sleep(3)

            os.system('TASKKILL /F /IM excel.exe')
        except:
            easygui.msgbox(msg="ERROR: WRONG! NO ERROR", title="ERROR")

        self.onButtonClick()

    def onButtonClick(self):
        #start worker thread for reading directories
        self.calc = Admin()
        self.calc.completeChanged.connect(self.oncompleteChanged)
        self.calc.scanChanged.connect(self.onscanChanged)
        self.calc.start()

    #for updating GUI box text
    def oncompleteChanged(self, total):
        self.complete.setText(total)

    #for updating text in GUI for what the program is currectly checking
    def onscanChanged(self, part):
        self.running.setText(part)
        self.running.adjustSize()

class Admin(QThread):
    completeChanged = pyqtSignal(str)
    scanChanged = pyqtSignal(str)

    def __init__(self):
        super(Admin, self).__init__()

    #method that automatically starts when Qthread is executed, a pyqt5 feature
    def run(self):
        user = getpass.getuser()
        self.qp_number = []

        try:
            location_qpwaiting = "pathway to excel file"

            #read excel file
            qpwaiting_excel = pd.read_excel(location_qpwaiting)

            k = 0
            while k < len(qpwaiting_excel):
                if "nan" not in str(qpwaiting_excel.iloc[k, 0]):
                    qp_exists = self.check_qp_fai_folder(str(qpwaiting_excel.iloc[k, 0]))
                    #check to make sure QP wasn't accidentally put in FAI folder already

                    if qp_exists == True:
                        self.qp_number.append("ERROR: " + str(qpwaiting_excel.iloc[k, 0]) + " QP IN FOLDER!")

                    fai_exists, jobnumber = self.find_fai(str(qpwaiting_excel.iloc[k, 0]))
                    if fai_exists == True:
                        self.qp_number.append(str(qpwaiting_excel.iloc[k, 0]) + "  " + jobnumber)
                k+=1

            #output results to GUI
            if len(self.qp_number) > 0:
                tostringqpnumber = '\n\n'.join(self.qp_number)
                self.completeChanged.emit(tostringqpnumber)
            else:
                self.completeChanged.emit("0")

            self.scanChanged.emit("      Done")
        except:
            easygui.msgbox(msg="ERROR: Something stupid happened!", title="ERROR")

    #check if a certain file is in a folder
    def check_qp_fai_folder(self, qp_number):
        yes_qp = False

        folderfai = "O:pathway to a specific directory"
        faipartpath = "O:pathway to a specific directory"

        #if file exists then return yes_qp to original method that claled this method
        if os.path.exists(faipartpath):
            readfolder = os.listdir(faipartpath)
            k = 0
            while k < len(readfolder):
                if partnumber_withrev in readfolder[k]:
                    if "QP" in readfolder[k]:
                        yes_qp = True
                k +=1

        return(yes_qp)


    def find_fai(self, qpnumber):
        scan_text = "Scanning" + " " + qpnumber + "....."
        self.scanChanged.emit(scan_text)
        fai_exists = False

        folderscans = "O:pathway to specific directory
        scansparthpath = "O:pathway to specific directory

        possible_folders = []

        #if pathway exists, walk through ALL folders in directory and if certain files contained
        #in those folders contain certain verbage or don't contain certain verbage, then append the file name
        #to a list for outputting to GUI
        if os.path.exists(scansparthpath):
            for path, subdirs, files in os.walk(scansparthpath):
                if scansparthpath != path:
                    for i in files:
                        if partnumber_withrev in i:
                            if "TRAVELER" not in i.upper():
                                split_path = path.split("\\")
                                folder = split_path[-1]
                                if "PROT" not in folder:
                                    if "PO" in folder:
                                        folder = folder[0:8]
                                        if folder not in possible_folders:
                                            possible_folders.append(folder)
                                    else:
                                        folder = folder[0:9]
                                        if folder not in possible_folders:
                                            possible_folders.append(folder)



        if len(possible_folders) > 0:
            total_folders = '  '.join(possible_folders)
            fai_exists = True

            return(fai_exists, total_folders)
        else:
            return(fai_exists, "0")



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Actions()
    sys.exit(app.exec_())