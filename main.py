#!/usr/bin/env python

from PyQt4 import QtGui # Import the PyQt4 module we'll need
from PyQt4 import QtCore
from PyQt4.QtCore import QSettings, QPoint, QSize
from PyQt4.QtCore import QThread, SIGNAL
from win32com.client import DispatchEx
import sys # We need sys so that we can pass argv to QApplication
import os
import re
import MainWindow # This file holds our MainWindow and all design related things

# ToDo
# - update path input when focus is lost

# C:\Python27\Lib\site-packages\PyQt4\pyuic4 MainWindow.ui  -o MainWindow.py
# pyinstaller --onefile --windowed --icon relink.ico  --name "Excel Refresh Links" "Excel Refresh Links FWW.spec" main.py

# Excel Constants


class MainAppWindow(QtGui.QMainWindow, MainWindow.Ui_MainWindow):
    
    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)
        
        self.excel = False
        self.max_history = 7
        self.msg = ""

        if getattr(sys, 'frozen', False):
            bundle_dir = sys._MEIPASS
        else:
            # we are running in a normal Python environment
            bundle_dir = os.path.dirname(os.path.abspath(__file__))

        self.setWindowIcon(QtGui.QIcon(os.path.join(bundle_dir,"relink.ico")))
        
        # work in INI File Stuff here
        QtCore.QCoreApplication.setOrganizationName("NRB")
        QtCore.QCoreApplication.setOrganizationDomain("northriverboats.com")
        QtCore.QCoreApplication.setApplicationName("Excel Refresh Links")
        self.settings = QSettings()
        
        # Initial window size/pos last saved. Use default values for first time
        self.resize(self.settings.value("size", QSize(900, 500)).toSize())
        self.move(self.settings.value("pos", QSize(0, 0)).toPoint())
        # need rules to keep window no larger than desktop
        # need rules to keep window pos visible on screen
        
        self.recent = []
        for i in range(self.max_history):
           self.recent.append(self.settings.value("recent" + str(unichr(48 + i)), "").toString())
        self.redrawMenu()
        
        #  do menu things
        # self.submenu2.menuAction().setVisible(False)
        # self.menuitem.setVisible(False)
        # self.menu_Recent.menuAction().setVisible(False)
        # self.actionRecent1.setText(self.recent[0])
        
        # self.actionRecent1.setVisible(False)

        # self.iniFolder = os.path.join(os.path.join(os.path.expanduser("~"), ".config"), "ExcelRelink")
        # self.iniFile = os.path.join(self.iniFolder, "config.ini")
        # if not os.path.exists(self.iniFolder):
        #   os.makedirs(self.iniFolder)	

        # generic slots and signals
        self.actionExit.triggered.connect(self.closeEvent)
        self.actionRelink.triggered.connect(self.doRelink)
        self.actionRecent1.triggered.connect(self.doRecent1)
        self.actionRecent2.triggered.connect(self.doRecent2)
        self.actionRecent3.triggered.connect(self.doRecent3)
        self.actionRecent4.triggered.connect(self.doRecent4)
        self.actionRecent5.triggered.connect(self.doRecent5)
        self.actionRecent6.triggered.connect(self.doRecent6)

        self.btnBrowse.clicked.connect(self.browseEvent)
        self.btnRelink.clicked.connect(self.doRelink)
        self.btnCancel.clicked.connect(self.update_done)
        app = QtGui.QApplication.instance()
        # app.focusChanged.connect(self.changed_focus)

        self.btnCancel.hide()
		
        # set status bar
        self.statusbar.showMessage("System Status | Idle")
        
    def closeEvent(self, e):        
        # Write window size and position to config file
        self.settings.setValue("size", self.size())
        self.settings.setValue("pos", self.pos())
        for i in range(self.max_history):
           self.settings.setValue("recent" + str(unichr(48 + i)), self.recent[i])
        sys.exit()
        
    def browseEvent(self):
        default_dir = self.recent[0] or os.path.join(os.path.expanduser("~"), "Desktop")
        my_dir = QtGui.QFileDialog.getExistingDirectory(self, "Open a folder", default_dir, QtGui.QFileDialog.ShowDirsOnly)
        if my_dir != "":
            self.changePath(my_dir)
        
    
    def changePath(self, my_dir):
        if my_dir in self.recent:
            self.recent.remove(my_dir)
        self.recent = [my_dir] + self.recent[:self.max_history - 1]
        self.redrawMenu()


    def redrawMenu(self):
        for i in range(1, self.max_history):
           getattr(self, "actionRecent" + str(unichr(48 + i))).setText(self.recent[i])
           getattr(self, "actionRecent" + str(unichr(48 + i))).setVisible(self.recent[i] != "")
        self.lePath.setText(self.recent[0])

    def doRecent1(self):
        self.changePath(self.recent[1])
        
    def doRecent2(self):
        self.changePath(self.recent[2])
        
    def doRecent3(self):
        self.changePath(self.recent[3])
        
    def doRecent4(self):
        self.changePath(self.recent[4])
        
    def doRecent5(self):
        self.changePath(self.recent[5])
        
    def doRecent6(self):
        self.changePath(self.recent[6])
        
    def doRelink(self):
        # if text NOT EQUAL self.recent[0] update
        # if self.recent[0] is blank THEN bail
        # if self.recent[0] is not a folder THEN bail
        
        
        self.btnRelink.hide()
        self.btnCancel.show()
        self.statusbar.showMessage("System Status | Starting Excel")
        print('Starting excel')
        self.excel = DispatchEx("Excel.Application")
        self.excel.Visible = 0
        self.msg = 'Completed'
        self.statusbar.showMessage("System Status | Excel Started")
        self.update_links_thread = update_links_thread(self.recent[0], self.excel)
        self.connect(self.update_links_thread, SIGNAL('update_statusbar(QString)'), self.update_statusbar)
        self.connect(self.update_links_thread, SIGNAL('update_progresssbar(int)'), self.update_progressbar)
        self.connect(self.update_links_thread, SIGNAL("finished()"), self.update_done)
        self.update_links_thread.start()
        self.btnCancel.clicked.connect(self.update_abort)
 

    def update_progressbar(self, int):
        self.progress_bar.setValue(int)
        
    def update_statusbar(self, message):
        self.statusbar.showMessage("System Status | " + message)

    def update_abort(self):
        self.msg = 'Canceled'
        self.update_links_thread.terminate()
        
    def update_done(self):
        if self.excel:
            self.btnCancel.hide()
            self.btnRelink.show()
            print("--" + self.msg)
            self.update_statusbar(self.msg)
            print('Stopping Excel')
            self.excel.Application.Quit()
            self.excel = False

class update_links_thread(QThread):

    def __init__(self, path, excel):
        QThread.__init__(self)
        self.path = path
        self.excel = excel

    def __del__(self):
        self.wait()

    def _file_filter(self, root, file):
        return (file.endswith(".xls") or file.endswith(".xlsx") or file.endswith(".xlsm"))

    def _file_scan(path):
        pass
        
    def run(self):
        self.file_list = []
        for root, dirs, files in os.walk(str(self.path)):
            for file in files:
                if self._file_filter(root, file):
                    self.file_list.append(os.path.join(root, file))
                    # print(os.path.join(root,file))
                    self.emit(SIGNAL('update_statusbar(QString)'), 'Files Found ' + str(len(self.file_list)))
        self.emit(SIGNAL('update_statusbar(QString)'), 'Scanning Completed ' + str(len(self.file_list)) + ' Files Found')
        self.emit(SIGNAL('update_progressbar(int)'), 50)
     
		


def main():
    app = QtGui.QApplication(sys.argv)  # A new instance of QApplication
    form = MainAppWindow()                 # We set the form to be our Main App Wehdiw (design)
    form.show()                         # Show the form
    app.exec_()                         # and execute the app


if __name__ == '__main__':              # if we're running file directly and not importing it
    main()                              # run the main function
