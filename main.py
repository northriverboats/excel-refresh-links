#!/usr/bin/env python

from PyQt4 import QtGui # Import the PyQt4 module we'll need
from PyQt4 import QtCore
from PyQt4.QtCore import QSettings, QPoint, QSize
import sys # We need sys so that we can pass argv to QApplication
import os
import re
from win32com.client import DispatchEx
import MainWindow # This file holds our MainWindow and all design related things

# C:\Python27\Lib\site-packages\PyQt4\pyuic4 MainWindow.ui  -o MainWindow.py
# pyinstaller --onefile --windowed --icon relink.ico  --name "Excel Refresh Links" "Excel Refresh Links FWW.spec" main.py

# Excel Constants
class Foo(object):
    pass
  
c = Foo()
c.xlAscending = 1
c.xlDescendig = 2
c.xlPasteFormulaAndNumberFormats = 11
c.xlTextValues = 2
c.xlEdgeBottom = 9
c.xlContinuous = 1


class MainAppWindow(QtGui.QMainWindow, MainWindow.Ui_MainWindow):
    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)

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
        for i in range(6):
           self.recent.append(self.settings.value("recent" + str(unichr(49 + i)), "").toString())
           getattr(self, "actionRecent" + str(unichr(49 + i))).setText(self.recent[i])
           getattr(self, "actionRecent" + str(unichr(49 + i))).setVisible(self.recent[i] != "")

        self.lePath.setText(self.recent[0])
        
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
        self.actionRelink.triggered.connect(self.closeEvent)
        self.btnBrowse.clicked.connect(self.closeEvent)
        app = QtGui.QApplication.instance()
        # app.focusChanged.connect(self.changed_focus)

	    # Fire Up COM
        # self.oExcel = DispatchEx('Excel.Application')
        # self.oExcel.Visible = 0
		
        # set status bar
        self.statusbar.showMessage("System Status | " + self.recent[0])
        
    def closeEvent(self, e):        
        # Write window size and position to config file
        self.settings.setValue("size", self.size())
        self.settings.setValue("pos", self.pos())
        self.settings.setValue("recent1", self.recent[0])
        # self.settings.setValue("pos", self.pos())
        # self.settings.value("recent1", "")
        
        sys.exit()



		


def main():
    app = QtGui.QApplication(sys.argv)  # A new instance of QApplication
    form = MainAppWindow()                 # We set the form to be our Main App Wehdiw (design)
    form.show()                         # Show the form
    app.exec_()                         # and execute the app


if __name__ == '__main__':              # if we're running file directly and not importing it
    main()                              # run the main function
