#!/usr/bin/env python

from PyQt4 import QtGui # Import the PyQt4 module we'll need
from PyQt4 import QtCore
from PyQt4.QtCore import QSettings, QPoint, QSize
from PyQt4.QtCore import QThread, SIGNAL
from win32com.client import DispatchEx
import pythoncom
import sys # We need sys so that we can pass argv to QApplication
import os
import re
import MainWindow # This file holds our MainWindow and all design related things

# ToDo
# - update path input when focus is lost
# - if path input empty do not allow relink
# DORELINK
# - if text NOT EQUAL self.recent[0] update
# - if self.recent[0] is blank THEN bail
# - if self.recent[0] is not a folder THEN bail
# - if exit or X incorner is clicked while running
# -   block it
# -   set flag and do abort
# -   durring done check for flag and need to run normal exit

# C:\Python27\Lib\site-packages\PyQt4\pyuic4 MainWindow.ui  -o MainWindow.py
# pyinstaller --onefile --windowed --icon relink.ico  --name "Excel Refresh Links" "Excel Refresh Links FWW.spec" main.py

# Excel Constants


class MainAppWindow(QtGui.QMainWindow, MainWindow.Ui_MainWindow):
    
    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)
        
        self.max_history = 7
        self.output_text = ""
        self.exit_flag = False

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

        # generic slots and signals
        self.actionExit.triggered.connect(self._closeEvent)
        self.actionRelink.triggered.connect(self.doRelink)
        self.actionRecent1.triggered.connect(self.doRecent1)
        self.actionRecent2.triggered.connect(self.doRecent2)
        self.actionRecent3.triggered.connect(self.doRecent3)
        self.actionRecent4.triggered.connect(self.doRecent4)
        self.actionRecent5.triggered.connect(self.doRecent5)
        self.actionRecent6.triggered.connect(self.doRecent6)

        self.btnBrowse.clicked.connect(self.browseEvent)
        self.btnRelink.clicked.connect(self.doRelink)
        app = QtGui.QApplication.instance()

        self.btnCancel.hide()
        
        # set status bar
        self.statusbar.showMessage("System Status | Idle")
        
    def closeEvent(self, e):
        self.exit_flag = True
        try:
            self.update_links_thread.running = False
            e.ignore()
        except:
            pass
 
    
    
    def _closeEvent(self, e): 
        # Write window size and position to config file
        self.settings.setValue("size", self.size())
        self.settings.setValue("pos", self.pos())
        # Write recent paths to config file
        for i in range(self.max_history):
           self.settings.setValue("recent" + str(unichr(48 + i)), self.recent[i])
        self.exit_flag = True
        # if relinking is taking place, 
        if self.update_links_thread.running:
            self.update_statusbar('Canceled')
            self.update_links_thread.running = False
        else:
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
        if not self.recent[0]:
            return
        self.btnRelink.hide()
        self.btnCancel.show()
        self.btnBrowse.setEnabled(False)
        self.lePath.setEnabled(False)
        self.actionRelink.setDisabled(True)
        self.menu_Recent.setDisabled(True)
        self.actionSave.setDisabled(True)
        self.update_links_thread = update_links_thread(self.recent[0])
        self.connect(self.update_links_thread, SIGNAL('update_statusbar(QString)'), self.update_statusbar)
        self.connect(self.update_links_thread, SIGNAL('update_progressbar(int)'), self.update_progressbar)
        self.connect(self.update_links_thread, SIGNAL('clear_textarea()'), self.clear_textarea)
        self.connect(self.update_links_thread, SIGNAL('update_textarea(QString)'), self.update_textarea)
        self.connect(self.update_links_thread, SIGNAL("finished()"), self.update_done)
        self.update_links_thread.start()
        self.btnCancel.clicked.connect(self.update_abort)
 
    def clear_textarea(self):
        self.textArea.clear()
        self.textArea.centerOnScroll = True
        
    def update_textarea(self, Qstring):
        self.textArea.append(Qstring)
        self.textArea.ensureCursorVisible()

    def update_progressbar(self, num):
        self.progressBar.setValue(num)
        
    def update_statusbar(self, message):
        self.statusbar.showMessage("System Status | " + message)

    def update_abort(self):
        self.update_links_thread.running = False
        
    def update_done(self):
        self.btnCancel.hide()
        self.btnRelink.show()
        self.btnBrowse.setEnabled(True)
        self.lePath.setEnabled(True)
        self.actionRelink.setDisabled(False)
        self.menu_Recent.setDisabled(False)
        self.actionSave.setDisabled(False)
        # If exit was requested close program
        if self.exit_flag:
            self._closeEvent(0)
        

class update_links_thread(QThread):

    def __init__(self, path):
        QThread.__init__(self)
        self.path = path

    def __del__(self):
        self.wait()

    def _file_filter(self, root, file):
        if file[:2] == "~$":
            return False
        return (file.endswith(".xls") or file.endswith(".xlsx") or file.endswith(".xlsm"))

    def _file_scan(path):
        pass
        
    def run(self):
        self.running = True
        self.file_list = []
        self.emit(SIGNAL('update_progressbar(int)'), 0)
        self.emit(SIGNAL('clear_textarea()'))
        for root, dirs, files in os.walk(str(self.path)):
            for file in files:
                if self._file_filter(root, file):
                    self.file_list.append(os.path.join(root, file))
                    self.emit(SIGNAL('update_statusbar(QString)'), 'Files Found ' + str(len(self.file_list)))
                    if not self.running:
                        self.emit(SIGNAL('update_statusbar(QString)'), 'Canceled')
                        return
        self.emit(SIGNAL('update_statusbar(QString)'), 'Scanning Completed ' + str(len(self.file_list)) + ' Files Found')
        total_files = len(self.file_list)
        current_count = 0
        pythoncom.CoInitialize()
        excel = DispatchEx("Excel.Application")
        excel.Visible = 0
        self.emit(SIGNAL('update_statusbar(QString)'), 'Starting Excel')
        for file in self.file_list:
            current_count += 1
            self.emit(SIGNAL('update_progressbar(int)'), int(float(current_count) / total_files * 100))
            self.emit(SIGNAL('update_statusbar(QString)'), 'Relinking %d of %d' % (current_count, total_files))
            self.emit(SIGNAL('update_textarea(QString)'), file[len(self.path) + 1:])
            excel.DisplayAlerts = False 
            workbook = excel.Workbooks.Open(file, UpdateLinks=3)
            excel.Workbooks.Close()
            if not self.running:
                break

        excel.Application.Quit()
        pythoncom.CoUninitialize()
        if current_count == total_files:
            self.emit(SIGNAL('update_statusbar(QString)'), 'Completed')
            self.emit(SIGNAL('update_progressbar(int)'), 100)
        else:
            self.emit(SIGNAL('update_statusbar(QString)'), 'Canceled')
            self.emit(SIGNAL('update_progressbar(int)'), 0)
        self.sleep(2)
        
     
		


def main():
    app = QtGui.QApplication(sys.argv)  # A new instance of QApplication
    form = MainAppWindow()                 # We set the form to be our Main App Wehdiw (design)
    form.show()                         # Show the form
    app.exec_()                         # and execute the app


if __name__ == '__main__':              # if we're running file directly and not importing it
    main()                              # run the main function
