#!/usr/bin/env python

from PyQt4 import QtGui # Import the PyQt4 module we'll need
from PyQt4 import QtCore
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
       

        # generic slots and signals
        app = QtGui.QApplication.instance()
        # app.focusChanged.connect(self.changed_focus)

	    # Fire Up COM
        # self.oExcel = DispatchEx('Excel.Application')
        # self.oExcel.Visible = 0
		
        # set status bar
        self.statusbar.showMessage("System Status | Normal")


		


def main():
    app = QtGui.QApplication(sys.argv)  # A new instance of QApplication
    form = MainAppWindow()                 # We set the form to be our Main App Wehdiw (design)
    form.show()                         # Show the form
    app.exec_()                         # and execute the app


if __name__ == '__main__':              # if we're running file directly and not importing it
    main()                              # run the main function
