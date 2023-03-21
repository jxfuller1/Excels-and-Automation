from PyQt5.QtWidgets import * 
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
import sys
import os
from xlrd import open_workbook
import time
import pyautogui
import easygui
from distutils.dir_util import copy_tree
import win32gui
import win32con
import getpass
from PIL import Image

#overview of Program:
#finds a window on desktop in a specific program
#that program outputs an excel which i then read
#collect data from excel and put into lists
#change lists based on data in specific folders
#with resulting lists, go to a specific program and hit certain buttons and click on certain locations
#update GUI as it hits the buttons in the specific program


#setup some lists to be used by the worker thread
drawinglist = []
joblist = []
idlist = []
ncr = []

#def resource_path(relative_path):
#    try:
#        base_path = sys._MEIPASS
#    except Exception:
#        base_path = os.path.abspath(".")

#    return os.path.join(base_path, relative_path)

#legacy code, keeping if i ever want to revert the images back to relative resource path and include them in compile of exe file

#for getting all windows of parent window (used for determining proper spots on parent window to click)
def get_child_windows(parent):
    if not parent:
        return
    hwndChildList = []
    win32gui.EnumChildWindows(parent, lambda hwnd, param: param.append(hwnd), hwndChildList)
    return hwndChildList


#main window GUI
class Actions(QDialog):
    """
    Simple dialog that consists of a Progress Bar and a Button.
    Clicking on the button results in the start of a timer and
    updates the progress bar.
    """
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
# setting window geometry
        self.setGeometry(400,400,615,640)
        self.setFixedSize(self.size())
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
  
        # setting window action
        self.setWindowTitle("Auto-Conformity")

        self.label = QLabel("Instructions:\n- Go to Conformity and Select plane! \n- Set EpicEnquiries in main monitor \n- Put this program in other monitor \n- Hit Initiate!", self)
        self.label.adjustSize()
        self.label.move(10,25)
        
        self.about = QLabel("About: This Program will auto-complete the easy ones in conformity and return the NCR's to add manually", self)
        self.about.adjustSize()
        self.about.move(50,5)
               
        self.pbarlabel = QLabel("                                                                    ", self)
        self.pbarlabel.adjustSize()
        self.pbarlabel.move(255,115)

        self.endresult = QLabel("Ones to complete", self)
        self.endresult.adjustSize()
        self.endresult.move(10,165)

        self.endresult = QLabel("Completed", self)
        self.endresult.adjustSize()
        self.endresult.move(205,165)

        self.ncrlabel = QLabel("NCR's --    (ADD THESE MANUALLY)", self)
        self.ncrlabel.adjustSize()
        self.ncrlabel.move(395,195)

        self.arrow = QLabel("------>", self)
        self.arrow.adjustSize()
        self.arrow.move(162,340)

        self.totaltocomplete = QLabel("Total:", self)
        self.totaltocomplete.adjustSize()
        self.totaltocomplete.move(10,560)

        self.numbertocomplete = QLabel("0     ", self)
        self.numbertocomplete.adjustSize()
        self.numbertocomplete.move(40,560)

        self.totalcompleted = QLabel("Total:", self)
        self.totalcompleted.adjustSize()
        self.totalcompleted.move(200,560)

        self.numbercompleted = QLabel("0     ", self)
        self.numbercompleted.adjustSize()
        self.numbercompleted.move(230,560)
        
        self.button = QPushButton(self)
        self.button.setText("Initiate")
        self.button.clicked.connect(self.onButtonClick)
        self.button.move(275,90)
        
        self.date = QTextEdit(self)
        self.date.setGeometry(10,180,150,380)

        self.complete = QTextEdit(self)
        self.complete.setGeometry(200,180,150,380)

        self.ncrresult = QTextEdit(self)
        self.ncrresult.setGeometry(395,210,190,245)
        
        self.prog_bar = QProgressBar(self)
        self.prog_bar.setGeometry(10, 130, 600, 10)
        self.prog_bar.setRange(0, 100)
        self.prog_bar.setValue(0)

        self.button5 = QPushButton(self)
        self.button5.setText("Pause")
        #self.button5.clicked.connect(self.onPauseClick)
        self.button5.move(240,590)

        self.button6 = QPushButton(self)
        self.button6.setText("Resume")
        #self.button6.clicked.connect(self.onPauseClick)
        self.button6.move(315,590)

        self.shortcut_open_pause = QShortcut(QKeySequence('Alt+P'), self)
        self.shortcut_open_unpause = QShortcut(QKeySequence('Alt+R'), self)
        #self.shortcut_open.activated.connect(self.onPauseClick)

        # showing all the widgets
        self.show()
        
    #when start button blicked , read parent window i want to read and it exports data into an excel file for reading
    def onButtonClick(self):
        try:
            user = getpass.getuser()
            usersdesktop = "your desktop path"
            conformityexcel = "your excel path, temp place to save excel"

            #look for window on desktop and bring it to front and hit buttons on it (this is purely for my needs as
            #the program im hitting the buttons on has the capability to export data to an excel file that i can then read
            enquiries = win32gui.FindWindow(None, "********")
            win32gui.SetForegroundWindow(enquiries)

            enquirieschilds = get_child_windows(enquiries)

            total_parents = []
            for i in range(len(enquirieschilds)):
                a = win32gui.GetWindowText(enquirieschilds[i])
                if a == "Conformity":
                    total_parents.append(i)

            x, y, z, d = win32gui.GetWindowRect(enquirieschilds[total_parents[-1]])
            x = x + 25
            y = y + 10
            pyautogui.click(x, y)

            time.sleep(1)

            resfreshs = win32gui.FindWindowEx(enquiries, 0, None, "Export")
            position1 = win32gui.GetWindowRect(resfreshs)
            x = position1[0] + 20
            y = position1[1] + 10
            pyautogui.click(x, y)

            time.sleep(2)
            pyautogui.write("conformity")
            time.sleep(.5)
            pyautogui.press('enter')
            time.sleep(.7)
            pyautogui.press('enter')

            #after excel file saved, kill excel to close it down (don't need it on screen)
            time.sleep(3)
            os.system('TASKKILL /F /IM excel.exe')
            time.sleep(2)

            #start worker thread
            self.calc = External()
            self.calc.countChanged.connect(self.onCountChanged)
            self.calc.textChanged.connect(self.onTextChanged)
            self.calc.dateChanged.connect(self.ondateChanged)
            self.calc.completeChanged.connect(self.oncompleteChanged)
            self.calc.ncrresultChanged.connect(self.onncrresultChanged)
            self.calc.numbertocompleteChanged.connect(self.onnumbertocompleteChanged)
            self.calc.numbercompletedChanged.connect(self.onnumbercompletedChanged)

            self.calc.start()

            #for a pause/resume buttons in the GUI to stop the worker thread when i want it to stop
            self.button5.clicked.connect(self.calc.pausefunc)
            self.button6.clicked.connect(self.calc.pausestart)
            #shortcut is for shortcut keys... which aren't working at the moment
            self.shortcut_open_pause.activated.connect(self.calc.pausefunc)
            self.shortcut_open_unpause.activated.connect(self.calc.pausestart)
        except:
            easygui.msgbox(msg="ERROR 2002", title="ERROR")

        
    def onCountChanged(self, value):
        self.prog_bar.setValue(value)

    def onTextChanged(self, value1):
        self.pbarlabel.setText(value1)

    def ondateChanged(self, value2):
        self.date.setPlainText(value2)

    def oncompleteChanged(self, value3):
        self.complete.setPlainText(value3)

    def onncrresultChanged(self, value4):
        self.ncrresult.setPlainText(value4)

    def onnumbertocompleteChanged(self, value8):
        self.numbertocomplete.setText(value8)

    def onnumbercompletedChanged(self, value9):
        self.numbercompleted.setText(value9)
        

#worker thread
class External(QThread):
    """
    Runs a thread.
    """
    countChanged = pyqtSignal(int)
    textChanged = pyqtSignal(str)
    dateChanged = pyqtSignal(str)
    completeChanged = pyqtSignal(str)
    ncrresultChanged = pyqtSignal(str)
    numbertocompleteChanged = pyqtSignal(str)
    numbercompletedChanged = pyqtSignal(str)
    
    
    def __init__(self):
        super().__init__()
        self.pausevalue = True

    def pausefunc(self):
        self.pausevalue = False
        
    def pausestart(self):
        self.pausevalue = True
            
    def run(self):
        try:
            user = getpass.getuser()
            usersdesktop = "your desktop path"
            conformityexcel = "path to excel file to read"

            #read excel file
            book = open_workbook(conformityexcel, formatting_info=True)
            a = book.sheets()
            #reading first sheet only
            b = a[0]

            totaldrawinglist = []

            self.textChanged.emit("Reading Conformity")

            #iterate through first sheet to obtain data.  The following code is setup in a way
            #to collect the data from my specific excel files
            for row in range(1, b.nrows-1):

                while self.pausevalue == False:
                    time.sleep(0)

                drawingnumberlist = b.cell(row, 2).value
                totaldrawinglist.append(drawingnumberlist)

                #for progress bar calculation , emit to progress bar GUI
                fieldraw = round(((row / b.nrows)*100) / 9 )
                self.countChanged.emit(fieldraw)
                #print("test")


            for row in range(1, b.nrows-1):

                while self.pausevalue == False:
                    time.sleep(0)

                #get data from row
                conformed = b.cell(row, 7).value
                NA = b.cell(row, 8).value
                HAVE = b.cell(row, 9).value
                jobfilled = b.cell(row, 13).value
                drawingnumber = b.cell(row, 2).value
                idnumber = b.cell(row, 14).value


                fieldraw1 = round(((row / b.nrows)*100) / 9 )
                field9 = fieldraw + fieldraw1
                self.countChanged.emit(field9)


                if "False" in conformed:
                    if "False" in NA:
                        if "False" in HAVE:
                            testjobcell = len(jobfilled)
                            if testjobcell > 1:
                                drawinglist.append(drawingnumber)
                                joblist.append(jobfilled)
                                idlist.append(idnumber)
                #print("test1")
                    #put data into lists

            times = len(drawinglist)
            k = 0
            time.sleep(1)
            self.textChanged.emit("Filtering Conformity")


            while k < times:

                while self.pausevalue == False:
                    time.sleep(0)

                totaloftotaldrawinglist = totaldrawinglist.count(drawinglist[k])
                totalofbaselist = drawinglist.count(drawinglist[k])

                fieldraw = round(((k / (times-1))*100) / 9)

                field1 = field9 + fieldraw
                self.countChanged.emit(field1)

                if totalofbaselist != totaloftotaldrawinglist:
                    indexlocation = [idx for idx, s in enumerate(drawinglist) if drawinglist[k] in s][0]
                    del drawinglist[indexlocation]
                    del joblist[indexlocation]
                    del idlist[indexlocation]
                    times -= 1
                    k-=1
                k +=1

    # this code removes items from lists if they don't match the number in the raw list

            #print(drawinglist)
            #print(joblist)
            #print(idlist)

            times = len(drawinglist)
            k = 0

            while k < times:

                while self.pausevalue == False:
                    time.sleep(0)

                if times == 1:
                    times2 = times + 1
                    fieldraw = round(((k / (times2-1))*100) / 9)
                else:
                    fieldraw = round(((k / (times-1))*100) / 9)


                field2 = field1 + fieldraw
                self.countChanged.emit(field2)

                totalofbaselist = drawinglist.count(drawinglist[k])

                if totalofbaselist > 1:
                    indices = []
                    for i in range(len(drawinglist)):
                        if drawinglist[i] == drawinglist[k]:
                            indices.append(i)
                    totalofindices = len(indices)
                    base = 0
                    valid = "None"
                    while base < totalofindices:
                        stringvalue = joblist[indices[base]]
                        if "JOB" in stringvalue:
                            valid = "True"
                        if "RWK" in stringvalue:
                            valid = "True"
                        if "ECN" in stringvalue:
                            valid = "True"
                        base += 1

                    if "True" in valid:
                        base = 0
                        while base < totalofindices:
                            indexlocation = [idx for idx, s in enumerate(drawinglist) if drawinglist[k] in s][0]

                            del joblist[indexlocation]
                            del drawinglist[indexlocation]
                            del idlist[indexlocation]
                            totalofindices -=1
                            times -=1
                        k-=1

                k +=1

    # if multiple line items for same drawing removes them if they have a job

            times = len(drawinglist)
            k = 0
            while k < times:

                while self.pausevalue == False:
                    time.sleep(0)

                if times == 1:
                    times2 = times + 1
                    fieldraw = round(((k / (times2-1))*100) / 9)
                else:
                    fieldraw = round(((k / (times-1))*100) / 9)

                field3 = field2 + fieldraw
                self.countChanged.emit(field3)

                totalofbaselist = drawinglist.count(drawinglist[k])
                valid = "None"
                if totalofbaselist == 1:
                    stringvalue = joblist[k]
                    if "JOB" in stringvalue:
                        valid = "True"
                    POcheckup = stringvalue[0:2]
                    if "PO" in POcheckup:
                        valid = "True"
                    if "None" in valid:
                        del joblist[k]
                        del drawinglist[k]
                        del idlist[k]
                        times -=1
                        k -=1
                k +=1

    #if drawing has only one line item and not a job or po removes it from list


            times = len(drawinglist)
            k = 0
            while k < times:

                while self.pausevalue == False:
                    time.sleep(0)

                if times == 1:
                    times2 = times + 1
                    fieldraw = round(((k / (times2-1))*100) / 9)
                else:
                    fieldraw = round(((k / (times-1))*100) / 9)


                field4 = field3 + fieldraw
                self.countChanged.emit(field4)

                totalofbaselist = drawinglist.count(drawinglist[k])
                delete = "None"
                if totalofbaselist > 1:
                    indices = []
                    for i in range(len(drawinglist)):
                        if drawinglist[i] == drawinglist[k]:
                            indices.append(i)
                    indicestotal = len(indices)
                    enumerateindice = 0
                    while enumerateindice < indicestotal:
                        if "PO" not in joblist[indices[enumerateindice]]:
                            enumerateid =0
                            idenumerate = []
                            while enumerateid < indicestotal:
                                idenumerate.append(idlist[indices[enumerateid]])
                                enumerateid +=1
                            find = any(s for s in idenumerate if joblist[indices[enumerateindice]] in s)
                            if find == False:
                                delete = "True"
                        enumerateindice +=1

                if "True" in delete:
                    iterateindice = 0
                    while iterateindice < indicestotal:
                        del joblist[indices[0]]
                        del drawinglist[indices[0]]
                        del idlist[indices[0]]
                        iterateindice += 1
                        times -= 1
                        k -=1

                k +=1


    #if drawing has multiple line items and contains a lot number instead of po number and that
    #and that lot number doesn't appear in idlist , deletes all the line items from lists

    #print(drawinglist)
    #print(joblist)
    #print(idlist)


            times = len(drawinglist)
            k = 0
            self.textChanged.emit("Filtering Thru Folders")
            while k < times:

                while self.pausevalue == False:
                    time.sleep(0)

                if times == 1:
                    times2 = times + 1
                    fieldraw = round(((k / (times2-1))*100) / 9)
                else:
                    fieldraw = round(((k / (times-1))*100) / 9)

                field5 = field4 + fieldraw
                self.countChanged.emit(field5)


                #here I am checking to see if there is more than 1 folder for the
                #value im looking for in a specific pathway.  this is specific to my needs
                #as if there's more than 1 folder in the pathway i want to remove it from my lists
                pathscans = "O:path to read where file should exist"

                totalofbaselist = drawinglist.count(drawinglist[k])

                if totalofbaselist == 1:
                    readfolder = os.listdir(pathscans)
                    numberoftimes = sum(jobnumber in s for s in readfolder)
                    if numberoftimes != 1:
                        del joblist[k]
                        del drawinglist[k]
                        del idlist[k]
                        times -= 1
                        k-=1

                if totalofbaselist > 1:
                    if "PO" in jobnumber:
                        readfolder = os.listdir(pathscans)
                        numberPO = sum(jobnumber in s for s in readfolder)
                        if numberPO != 1:
                            indicesPO = []
                            for i in range(len(drawinglist)):
                                if drawinglist[i] == drawinglist[k]:
                                    indicesPO.append(i)
                            deletetotal = len(indicesPO)
                            m = 0
                            while m < deletetotal:
                                del joblist[indicesPO[0]]
                                del drawinglist[indicesPO[0]]
                                del idlist[indicesPO[0]]
                                times -= 1
                                k-=1
                                m +=1
                k +=1

    #this code checks if none or more than 1 folder in inspection scans , if not , deletes from lists

    #print(drawinglist)
    #print(joblist)
    #print(idlist)

            times = len(drawinglist)
            k = 0
            while k < times:

                while self.pausevalue == False:
                    time.sleep(0)

                if times == 1:
                    times2 = times + 1
                    fieldraw = round(((k / (times2-1))*100) / 9)
                else:
                    fieldraw = round(((k / (times-1))*100) / 9)


                field6 = field5 + fieldraw
                self.countChanged.emit(field6)

                #honestly i can't remember what this does something to do with checking folders at a specific
                #location in that is specific to my needs
                pathscans = "O:path to folder"
                readfolder = os.listdir(pathscans)

                totalofbaselist = drawinglist.count(drawinglist[k])

                if totalofbaselist == 1:
                    indicesread = []
                    for i in range(len(readfolder)):
                        if jobnumber in readfolder[i]:
                            indicesread.append(i)

                    res = readfolder[indicesread[0]]

                    #for checking folders at another location that is specific to my needs
                    jobscans = "O:pathway to folders"
                    if checkforqp not in res:
                        readfolderjob = os.listdir(jobscans)
                        iteratefolder = len(readfolderjob)
                        dup = 0
                        value = "None"
                        while dup < iteratefolder:
                            if partnumber in readfolderjob[dup]:
                                #if jobnumber in readfolderjob[dup]:
                                value = "True"
                            dup +=1
                        if "None" in value:
                            del drawinglist[k]
                            del joblist[k]
                            del idlist[k]
                            times -=1
                            k-=1


                if totalofbaselist > 1:
                    if "PO" in jobnumber:
                        indicesread = []
                        for i in range(len(readfolder)):
                            if jobnumber in readfolder[i]:
                                indicesread.append(i)
                        res = readfolder[indicesread[0]]

                        jobscans = '\\'.join([folderscans, chapterscans, partnumber, res])
                        if checkforqp not in res:
                            readfolderjob = os.listdir(jobscans)
                            iteratefolder = len(readfolderjob)
                            dup = 0
                            value = "None"
                            while dup < iteratefolder:
                                if partnumber in readfolderjob[dup]:
                                    #if jobnumber in readfolderjob[dup]:
                                    value = "True"
                                dup +=1
                            if "None" in value:
                                indicesPO = []
                                for i in range(len(drawinglist)):
                                    if drawinglist[i] == drawinglist[k]:
                                        indicesPO.append(i)
                                iterateindice = 0
                                totalidices = len(indicesPO)
                                while iterateindice < totalidices:
                                    del joblist[indicesPO[0]]
                                    del drawinglist[indicesPO[0]]
                                    del idlist[indicesPO[0]]
                                    iterateindice += 1
                                    times -= 1
                                    k-=1
                k +=1

    #this block of code checks to see if FAI in folder if folder is not labeled with QP
    #if not QP and No fai, deletes it from lists

    #print(drawinglist)
    #print(joblist)
    #print(idlist)

            ncrlist = []

            times = len(drawinglist)
            k = 0
            while k < times:

                while self.pausevalue == False:
                    time.sleep(0)

                if times == 1:
                    times2 = times + 1
                    fieldraw = round(((k / (times2-1))*100) / 9 + 1)
                else:
                    fieldraw = round(((k / (times-1))*100) / 9 + 1)

                field7 = field6 + fieldraw
                self.countChanged.emit(field7)

                #for checking folders at specific directory again, this is specific for my needs in looking
                #for specific file in those folders
                pathscans = "O: pathway to folder"

                readfolder = os.listdir(pathscans)

                totalofbaselist = drawinglist.count(drawinglist[k])

                if totalofbaselist == 1:
                    ncrtemp = []
                    ncrmaybe = "None"
                    indicesread = []
                    for i in range(len(readfolder)):
                        if jobnumber in readfolder[i]:
                            indicesread.append(i)

                    res = readfolder[indicesread[0]]
                    jobscans = "O:pwath way folders"
                    readfolderjob = os.listdir(jobscans)
                    lenreadfolderjob = len(readfolderjob)
                    checkforncrs = 0
                    while checkforncrs < lenreadfolderjob:
                        index2 = readfolderjob[checkforncrs][2]
                        if ncrcheck in index2:
                            ncrnumber = readfolderjob[checkforncrs][0:7]
                            ncrmaybe = "True"
                            ncrtemp.append(ncrnumber)
                        checkforncrs +=1
                    if "True" in ncrmaybe:
                        ncrlist.append(ncrtemp)
                    else:
                        ncrlist.append('')


                if totalofbaselist > 1:
                    ncrtemp = []
                    ncrmaybe = "None"
                    if "PO" in jobnumber:
                        indicesread = []
                        for i in range(len(readfolder)):
                            if jobnumber in readfolder[i]:
                                indicesread.append(i)

                        res = readfolder[indicesread[0]]
                        jobscans = '\\'.join([folderscans, chapterscans, partnumber, res])
                        readfolderjob = os.listdir(jobscans)
                        lenread = len(readfolderjob)
                        lentest = 0
                        while lentest < lenread:
                            lookup = readfolderjob[lentest]
                            lookup1 = lookup[2]
                            if ncrcheck in lookup1:
                                ncrmaybe = "True"
                                ncrnumber = readfolderjob[lentest][0:7]
                                ncrtemp.append(ncrnumber)
                            lentest +=1
                    if "True" in ncrmaybe:
                        ncrlist.append(ncrtemp)
                    else:
                        ncrlist.append('')

                k +=1

            self.textChanged.emit("Done Finding ones to complete")
            times = len(drawinglist)
            #error checking if drawinglist list has nothing in it
            if times == 1:
                    output1 = easygui.msgbox("Found none to Auto complete!", "DOOFUS", "OK")
                    exit()

            completionlist = []
            totalcompletion = 0

            while totalcompletion < times:
                combined = drawinglist[totalcompletion] + " " + joblist[totalcompletion]
                completionlist.append(combined)
                totalcompletion +=1

            tostringcompletionlist = '\n'.join(completionlist)

            self.dateChanged.emit(tostringcompletionlist)

            timesstring = str(times)
            self.numbertocompleteChanged.emit(timesstring)



    #this code creates a list for the ncrs

            #win32gui.SetForegroundWindow(enquiries)

            times = len(drawinglist)
            ncrtext = 0
            textncrlist = []
            valueofncr = "None"

            while ncrtext < times:
                checkforncr = len(ncrlist[ncrtext])
                if checkforncr >= 1:
                    joinedncrs = ", ".join(ncrlist[ncrtext])
                    ncrstring = drawinglist[ncrtext] + " " + joblist[ncrtext] + "; " + joinedncrs
                    valueofncr = "True"
                    textncrlist.append(ncrstring)
                ncrtext += 1
            if "None" in valueofncr:
                textncrlist.append("Detected no NCRs")


            stringncr = '\n'.join(textncrlist)

            self.ncrresultChanged.emit(stringncr)

            #outputting files found in certain folders to a txt file in case something happends to the program
            #and i need to revert back to it
            fileloc = usersdesktop + "conformityncrlist.txt"
            output_file = open(fileloc, 'w')
            for ncrs in textncrlist:
                output_file.write(ncrs + "\n")
            output_file.close()


            #the following takes the results from my lists, which is from reading the excel file and reading folders
            #and will hit buttons within a specific program based on those results
            times = len(drawinglist)
            checkmarkstuff = 0
            completedlist = []
            while checkmarkstuff < times:

                while self.pausevalue == False:
                    time.sleep(0)

                timeout_start = time.time() + 10
                drawingconformity = None
                while drawingconformity is None:
                    drawingconformity = pyautogui.locateOnScreen("O:path to screenshot", grayscale=False, confidence=.9)
                    if time.time() > timeout_start:
                        message = "Couldnt find drawing title column"
                        title = "error"
                        easygui.msgbox(message, title)
                        exit()
                p, o = pyautogui.center((drawingconformity))
                o1 = o + 15
                pyautogui.click((p, o1))
                pyautogui.write(drawinglist[checkmarkstuff])

                while self.pausevalue == False:
                    time.sleep(0)

                timeout_start = time.time() + 10
                arrowconformity = None
                while arrowconformity is None:
                    arrowconformity = pyautogui.locateOnScreen("O:path to screenshot",  grayscale=False, confidence=.8)
                    if time.time() > timeout_start:
                        message = "Couldnt find arrow column"
                        title = "error"
                        easygui.msgbox(message, title)
                        exit()
                x, m = pyautogui.center((arrowconformity))

                while self.pausevalue == False:
                    time.sleep(0)

                timeout_start = time.time() + 10
                jobconformity = None
                while jobconformity is None:
                    jobconformity = pyautogui.locateOnScreen("O:path to screenshot",  grayscale=False, confidence=.8)
                    if time.time() > timeout_start:
                        message = "Couldnt find job/po column"
                        title = "error"
                        easygui.msgbox(message, title)
                        exit()
                x, y = pyautogui.center((jobconformity))
                pyautogui.click((x, m))
                pyautogui.write(joblist[checkmarkstuff])

                while self.pausevalue == False:
                    time.sleep(0)

                timeout_start = time.time() + 10
                arrowconformity = None
                while arrowconformity is None:
                    arrowconformity = pyautogui.locateOnScreen("O:path to screenshot",  grayscale=False, confidence=.8)
                    if time.time() > timeout_start:
                        message = "Couldnt find arrow 2"
                        title = "error"
                        easygui.msgbox(message, title)
                        exit()
                x, m = pyautogui.center((arrowconformity))

                #this block of code gets exact location of line item in conformity

                checkforlot = ["JOB", "PO"]
                find = any(x in joblist[checkmarkstuff] for x in checkforlot)
                if find == False:

                    while self.pausevalue == False:
                        time.sleep(0)

                    timeout_start = time.time() + 10
                    naconformity = None
                    while naconformity is None:
                        naconformity = pyautogui.locateOnScreen("O:path to screenshot", grayscale=False, confidence=.8)
                        if time.time() > timeout_start:
                            message = "Couldnt find NA column"
                            title = "error"
                            easygui.msgbox(message, title)
                            exit()
                    x, y = pyautogui.center((naconformity))
                    pyautogui.click((x, m))

                    while self.pausevalue == False:
                        time.sleep(0)

                    timeout_start = time.time() + 10
                    commentconformity = None
                    while commentconformity is None:
                        commentconformity = pyautogui.locateOnScreen("O:path to screenshot",  grayscale=False, confidence=.8)
                        if time.time() > timeout_start:
                            message = "Couldnt find comment column"
                            title = "error"
                            easygui.msgbox(message, title)
                            exit()
                    x, y = pyautogui.center((commentconformity))
                    pyautogui.click((x, m))
                    pyautogui.write('duplicate')
                    pyautogui.click((p, o1))
                    time.sleep(3.5)
                else:

                    while self.pausevalue == False:
                        time.sleep(0)

                    timeout_start = time.time() + 10
                    conformedconformity = None
                    while conformedconformity is None:
                        conformedconformity = pyautogui.locateOnScreen("O:path to screenshot",  grayscale=False, confidence=.8)
                        if time.time() > timeout_start:
                            message = "Couldnt find conformed column"
                            title = "error"
                            easygui.msgbox(message, title)
                            exit()
                    x, y = pyautogui.center((conformedconformity))
                    pyautogui.click((x, m))

                    while self.pausevalue == False:
                        time.sleep(0)

                    timeout_start = time.time() + 10
                    haveconformity = None
                    while haveconformity is None:
                        haveconformity = pyautogui.locateOnScreen("O:path to screenshot",  grayscale=False, confidence=.8)
                        if time.time() > timeout_start:
                            message = "Couldnt find have column"
                            title = "error"
                            easygui.msgbox(message, title)
                            exit()
                    x, y = pyautogui.center((haveconformity))
                    pyautogui.click((x, m))
                    pyautogui.click((p, o1))
                    time.sleep(3.5)


                completedlist.append(completionlist[0])
                del completionlist[0]


                #output info to GUI on every iteration of loops to update GUI
                tostringcompletionlist = '\n'.join(completionlist)
                self.dateChanged.emit(tostringcompletionlist)

                tostringcompleted = '\n'.join(completedlist)
                self.completeChanged.emit(tostringcompleted)

                checkmarkstuff +=1

                stringcheckmark = str(checkmarkstuff)
                totalleft = times - checkmarkstuff
                stringtotalleft = str(totalleft)

                self.numbercompletedChanged.emit(stringcheckmark)
                self.numbertocompleteChanged.emit(stringtotalleft)


        except:
            easygui.msgbox(msg="ERROR 606", title="ERROR")
            drawinglist.clear()
            joblist.clear()
            idlist.clear()
            ncr.clear()

            #clear values if errors out


        drawinglist.clear()
        joblist.clear()
        idlist.clear()
        ncr.clear()

        #clear values when finished if want to re-run it again

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Actions()
    sys.exit(app.exec_())

