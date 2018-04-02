# -*- coding: utf-8 -*-

import ChangeReport
import sys
from PyQt5 import QtWidgets

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    mainWindows = ChangeReport.MyMainWindow()
    mainWindows.show()
    sys.exit(app.exec_())