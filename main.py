#!/usr/bin/env python

from PyQt4 import QtGui
from PyQt4 import QtCore
from PyQt4.QtCore import QSettings, QSize, QPoint # pylint: disable-msg=E0611
from PyQt4.QtCore import QThread, SIGNAL # pylint: disable-msg=E0611
from excel import ExcelDocument
from pathlib import Path
from time import sleep
from win32com.client import DispatchEx
from pythoncom import CoInitialize, CoUninitialize # pylint: disable-msg=E0611
import sys  # We need sys so that we can pass argv to QApplication
import os
import MainWindow  # This file holds our MainWindow and all design related things

r"""
Notes:
Lib\site-packages\PyQt4\pyuic4 MainWindow.ui  -o MainWindow.py
Scripts\pyinstaller.exe --onefile --windowed --icon options.ico  --name "Excel Refresh Links" "Excel Refresh Links FWW.spec" main.py

ToDo's
- Add Clear Text to the menu
- Don't clear text on every run
"""

# Excel Constants


class MainAppWindow(QtGui.QMainWindow, MainWindow.Ui_MainWindow):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.max_history = 7
        self.exit_flag = False

        if getattr(sys, 'frozen', False):
            bundle_dir = sys._MEIPASS # pylint: disable=no-member
        else:
            # we are running in a normal Python environment
            bundle_dir = os.path.dirname(os.path.abspath(__file__))

        # set program icon
        self.setWindowIcon(QtGui.QIcon(os.path.join(bundle_dir, "relink.ico")))

        # work in INI File Stuff here
        QtCore.QCoreApplication.setOrganizationName("NRB")
        QtCore.QCoreApplication.setOrganizationDomain("northriverboats.com")
        QtCore.QCoreApplication.setApplicationName("Excel Refresh Links")
        self.settings = QSettings()

        # Initial window size/pos last saved. Use default values for first time
        self.resize(self.settings.value("size", QSize(900, 500)))
        self.move(self.settings.value("pos", QPoint(0, 0)))
        # need rules to keep window no larger than desktop
        # need rules to keep window pos visible on screen

        self.recent = []
        for i in range(self.max_history):
            self.recent.append(self.settings.value("recent" + str(chr(48 + i)), ""))
        self.redrawMenu()

        # generic slots and signals
        self.actionAbout.triggered.connect(self.doAbout)
        self.actionSave.triggered.connect(self.doSave)
        self.actionClear_Output.triggered.connect(self.clear_textarea)
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

        # app = QtGui.QApplication.instance()

        self.btnCancel.hide()
        self.actionSave.setEnabled(False)
        self.actionClear_Output.setEnabled(False)

        # set status bar
        self.statusbar.showMessage("System Status | Idle")

    def closeEvent(self, e):
        self._closeEvent(0)

    def _closeEvent(self, e):
        # Write window size and position to config file
        self.settings.setValue("size", self.size())
        self.settings.setValue("pos", self.pos())
        # Write recent paths to config file
        for i in range(self.max_history):
            self.settings.setValue("recent" + str(chr(48 + i)), self.recent[i])
        self.exit_flag = True
        sys.exit(0)

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
            getattr(self, "actionRecent" + str(chr(48 + i))).setText(self.recent[i])
            getattr(self, "actionRecent" + str(chr(48 + i))).setVisible(self.recent[i] != "")
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
        # if path is not valid then bail
        if not os.path.isdir(self.recent[0]):
            self.update_statusbar('Path not valid')
        search = self.leSearch.text()
        replace = self.leReplace.text()
        partial = self.cbPartialMatch.checkState()
        if not search and replace:
            self.leReplace.setText("") 
        self.block_actions()
        self.update_links_thread = update_links_thread(self.recent[0], search, replace, partial)
        self.update_links_thread.update_statusbar.connect(self.update_statusbar)
        self.update_links_thread.update_progressbar.connect(self.update_progressbar)
        self.update_links_thread.clear_textarea.connect(self.clear_textarea)
        self.update_links_thread.update_textarea.connect(self.update_textarea)
        self.update_links_thread.finished.connect(self.update_done)
        self.update_links_thread.start()
        self.btnCancel.clicked.connect(self.update_abort)

    def doAbout(self, event):
        about_msg = "NRB Excel Relink Sheets\n©2019 North River Boats\nBy Fred Warren"
        _ = QtGui.QMessageBox.information(self, 'About',
                         about_msg, QtGui.QMessageBox.Ok)

    def doSave(self):
        text = self.textArea.toPlainText()
        if text == "":
            return
        name = QtGui.QFileDialog.getSaveFileName(self, 'Save File', '', 'TEXT (*.txt)')
        if name:
            file = open(name,'w')
            file.write(text)
            file.close()

    def clear_textarea(self):
        self.textArea.clear()
        self.textArea.centerOnScroll = True
        self.actionSave.setEnabled(False)

    def update_textarea(self, Qstring):
        self.textArea.append(Qstring)
        self.textArea.ensureCursorVisible()

    def update_progressbar(self, num):
        self.progressBar.setValue(num)

    def update_statusbar(self, message):
        self.statusbar.showMessage("System Status | " + message)
        if message == "Completed" or message == "Canceled":
            _ = QtGui.QMessageBox.information(self, 'About', message, QtGui.QMessageBox.Ok)
            self.actionSave.setEnabled(True)
            self.actionClear_Output.setEnabled(True)
            self.unblock_actions()

    def update_abort(self):
        self.update_links_thread.running = False
        self.btnCancel.setDisabled(True)

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

    def block_actions(self):
        self.btnRelink.hide()
        self.btnCancel.show()
        self.btnBrowse.setEnabled(False)
        self.lePath.setEnabled(False)
        self.actionRelink.setDisabled(True)
        self.menu_Recent.setDisabled(True)
        self.actionSave.setDisabled(True)
        
    def unblock_actions(self):
        self.btnRelink.show()
        self.btnCancel.hide()
        self.btnCancel.setEnabled(True)
        self.btnBrowse.setEnabled(True)
        self.lePath.setEnabled(True)
        self.actionRelink.setDisabled(False)
        self.menu_Recent.setDisabled(False)
        self.actionSave.setDisabled(False)

class update_links_thread(QThread):
    update_statusbar = QtCore.pyqtSignal(str)
    update_progressbar = QtCore.pyqtSignal(int)
    clear_textarea = QtCore.pyqtSignal()
    update_textarea = QtCore.pyqtSignal(str)
    finished = QtCore.pyqtSignal()

    def __init__(self, path, search, replace, partial):
        QThread.__init__(self)
        self.path = path
        self.search = search
        self.replace = replace
        self.partial = partial

    def __del__(self):
        self.wait()

    def _file_filter(self, root, file):
        if file[:2] == "~$":
            return False
        return (file.endswith(".xls") or file.endswith(".xlsx") or file.endswith(".xlsm"))

    def _file_scan(self, path):
        pass
        
    def excel_files(self, mypath):
        for root, _dirs, files in os.walk(str(mypath)):
            for file in files:
                if self._file_filter(root, file):
                    yield [root, file]

    def search_and_replace(self):
        if not self.search:
            return

    def run(self):
        self.running = True
        self.file_list = []
        self.update_progressbar.emit(0)
        self.clear_textarea.emit()
        self.update_textarea.emit("Relinking Files in {}\n".format(self.path))
        for root, file in self.excel_files(self.path):
            self.file_list.append(os.path.join(root, file))
            self.update_statusbar.emit('Files Found ' + str(len(self.file_list)))
            if not self.running:
                self.update_statusbar.emit('Canceled')
                return
        self.update_statusbar.emit('Scanning Completed ' + str(len(self.file_list)) + ' Files Found')
        total_files = len(self.file_list)
        current_count = 0
        CoInitialize()
        excel = ExcelDocument(visible=False)
        self.update_statusbar.emit('Starting Excel')
        for file in self.file_list:
            current_count += 1
            self.update_progressbar.emit(int(float(current_count) / total_files * 100))
            self.update_statusbar.emit('Relinking %d of %d' % (current_count, total_files))
            # Set text color
            self.update_textarea.emit(file[len(self.path) + 1:])
            excel.display_alerts(False)
            excel.open(file, updatelinks=3)
            self.search_and_replace()
            sleep(2) # give it some time to work it's magic
            excel.save()
            excel.close()
            if not self.running:
                break

        excel.quit()
        CoUninitialize()
        if current_count == total_files:
            self.update_statusbar.emit('Completed')
            self.update_progressbar.emit(100)
        else:
            self.update_statusbar.emit('Canceled')

        self.sleep(1)


def main():
    app = QtGui.QApplication(sys.argv)  # A new instance of QApplication
    form = MainAppWindow()                 # We set the form to be our Main App Wehdiw (design)
    form.show()                         # Show the form
    app.exec_()                         # and execute the app


if __name__ == '__main__':              # if we're running file directly and not importing it
    main()                              # run the main function
