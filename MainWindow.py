# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'MainWindow.ui'
#
# Created: Sat Apr 13 15:39:14 2019
#      by: PyQt4 UI code generator 4.10.3
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(529, 402)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.gridLayout = QtGui.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        self.horizontalLayout_8 = QtGui.QHBoxLayout()
        self.horizontalLayout_8.setObjectName(_fromUtf8("horizontalLayout_8"))
        spacerItem = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.horizontalLayout_8.addItem(spacerItem)
        self.btnRelink = QtGui.QPushButton(self.centralwidget)
        self.btnRelink.setObjectName(_fromUtf8("btnRelink"))
        self.horizontalLayout_8.addWidget(self.btnRelink)
        self.gridLayout.addLayout(self.horizontalLayout_8, 2, 0, 1, 1)
        self.horizontalLayout_7 = QtGui.QHBoxLayout()
        self.horizontalLayout_7.setObjectName(_fromUtf8("horizontalLayout_7"))
        self.lePath = QtGui.QLineEdit(self.centralwidget)
        self.lePath.setObjectName(_fromUtf8("lePath"))
        self.horizontalLayout_7.addWidget(self.lePath)
        self.btnBrowse = QtGui.QPushButton(self.centralwidget)
        self.btnBrowse.setObjectName(_fromUtf8("btnBrowse"))
        self.horizontalLayout_7.addWidget(self.btnBrowse)
        self.gridLayout.addLayout(self.horizontalLayout_7, 0, 0, 1, 1)
        self.teOutput = QtGui.QTextEdit(self.centralwidget)
        self.teOutput.setObjectName(_fromUtf8("teOutput"))
        self.gridLayout.addWidget(self.teOutput, 1, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 529, 21))
        self.menubar.setObjectName(_fromUtf8("menubar"))
        self.menuFile = QtGui.QMenu(self.menubar)
        self.menuFile.setObjectName(_fromUtf8("menuFile"))
        self.menu_Recent = QtGui.QMenu(self.menuFile)
        self.menu_Recent.setObjectName(_fromUtf8("menu_Recent"))
        self.menuHelp = QtGui.QMenu(self.menubar)
        self.menuHelp.setObjectName(_fromUtf8("menuHelp"))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtGui.QStatusBar(MainWindow)
        self.statusbar.setObjectName(_fromUtf8("statusbar"))
        MainWindow.setStatusBar(self.statusbar)
        self.actionSave = QtGui.QAction(MainWindow)
        self.actionSave.setObjectName(_fromUtf8("actionSave"))
        self.actionExit = QtGui.QAction(MainWindow)
        self.actionExit.setObjectName(_fromUtf8("actionExit"))
        self.actionAbout = QtGui.QAction(MainWindow)
        self.actionAbout.setObjectName(_fromUtf8("actionAbout"))
        self.actionRelink = QtGui.QAction(MainWindow)
        self.actionRelink.setObjectName(_fromUtf8("actionRelink"))
        self.actionRecent1 = QtGui.QAction(MainWindow)
        self.actionRecent1.setObjectName(_fromUtf8("actionRecent1"))
        self.actionRecent2 = QtGui.QAction(MainWindow)
        self.actionRecent2.setObjectName(_fromUtf8("actionRecent2"))
        self.actionRecent3 = QtGui.QAction(MainWindow)
        self.actionRecent3.setObjectName(_fromUtf8("actionRecent3"))
        self.actionRecent4 = QtGui.QAction(MainWindow)
        self.actionRecent4.setObjectName(_fromUtf8("actionRecent4"))
        self.actionRecent5 = QtGui.QAction(MainWindow)
        self.actionRecent5.setObjectName(_fromUtf8("actionRecent5"))
        self.actionRecent5_2 = QtGui.QAction(MainWindow)
        self.actionRecent5_2.setObjectName(_fromUtf8("actionRecent5_2"))
        self.actionRecent6 = QtGui.QAction(MainWindow)
        self.actionRecent6.setObjectName(_fromUtf8("actionRecent6"))
        self.menu_Recent.addAction(self.actionRecent1)
        self.menu_Recent.addAction(self.actionRecent2)
        self.menu_Recent.addAction(self.actionRecent3)
        self.menu_Recent.addAction(self.actionRecent4)
        self.menu_Recent.addAction(self.actionRecent5)
        self.menu_Recent.addAction(self.actionRecent6)
        self.menuFile.addAction(self.actionRelink)
        self.menuFile.addAction(self.menu_Recent.menuAction())
        self.menuFile.addAction(self.actionSave)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionExit)
        self.menuHelp.addAction(self.actionAbout)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow", None))
        self.btnRelink.setText(_translate("MainWindow", "Relink", None))
        self.btnBrowse.setText(_translate("MainWindow", "Browse", None))
        self.btnBrowse.setShortcut(_translate("MainWindow", "Ctrl+B", None))
        self.menuFile.setTitle(_translate("MainWindow", "&File", None))
        self.menu_Recent.setTitle(_translate("MainWindow", "&Recent", None))
        self.menuHelp.setTitle(_translate("MainWindow", "&Help", None))
        self.actionSave.setText(_translate("MainWindow", "&Save", None))
        self.actionSave.setShortcut(_translate("MainWindow", "Ctrl+S", None))
        self.actionExit.setText(_translate("MainWindow", "E&xit", None))
        self.actionExit.setShortcut(_translate("MainWindow", "Ctrl+X", None))
        self.actionAbout.setText(_translate("MainWindow", "&About", None))
        self.actionAbout.setShortcut(_translate("MainWindow", "Ctrl+A", None))
        self.actionRelink.setText(_translate("MainWindow", "&Relink", None))
        self.actionRelink.setShortcut(_translate("MainWindow", "Ctrl+R", None))
        self.actionRecent1.setText(_translate("MainWindow", "recent1", None))
        self.actionRecent2.setText(_translate("MainWindow", "recent2", None))
        self.actionRecent3.setText(_translate("MainWindow", "recent3", None))
        self.actionRecent4.setText(_translate("MainWindow", "recent4", None))
        self.actionRecent5.setText(_translate("MainWindow", "recent5", None))
        self.actionRecent5_2.setText(_translate("MainWindow", "recent5", None))
        self.actionRecent6.setText(_translate("MainWindow", "recent6", None))

