# -*- coding: utf-8 -*-

import sys
import os
from ChangeReport_UI import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox
from Data2Excel import Issue2Excel
from DataProcess import Issue


class MyMainWindow(QMainWindow, Ui_MainWindow):

    txtfilepath=""
    xlspath=""

    def __init__(self):
        super(MyMainWindow, self).__init__()
        self.setupUi(self)

    def findtxtreport(self):
        txtpath, ok = QFileDialog.getOpenFileName(self, "选择TXT报告", sys.path[0], "Text Files (*.txt)")
        self.lineEdit.setText(str(txtpath))
        self.txtfilepath = str(txtpath)
        self.xlspath = os.path.split(self.txtfilepath)[0]
        #print(self.xlspath)
        #print(self.txtfilepath)

    def txt2excle(self):
        reportdata = Issue.processtxtdata(self.txtfilepath)
        Issue2Excel.issues2excel(reportdata, self.xlspath + '\\issue.xls')
        self.successMessage()


    def successMessage(self):
        button = QMessageBox.question(self,"报告转换","报告转换成功！是否打开文件所在目录？",
                                      QMessageBox.Ok|QMessageBox.Cancel,QMessageBox.Ok)

        if button == QMessageBox.Ok:
            command = 'start ' + self.xlspath
            os.system(command)
        else:
            return
